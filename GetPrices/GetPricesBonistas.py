import re
import time
from pathlib import Path
from urllib.parse import quote

import pandas as pd
import requests


# =============================================================================
# PATHS
# =============================================================================

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = SCRIPT_DIR.parent

INPUT_DIR = SCRIPT_DIR / "output"
OUTPUT_DIR = SCRIPT_DIR / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

TICKERS_INPUT_PATH = INPUT_DIR / "tickers_cierre.xlsx"
BONISTAS_OUTPUT_PATH = OUTPUT_DIR / "bonistas_precios.xlsx"


# =============================================================================
# CONFIG
# =============================================================================

BONISTAS_HOME = "https://bonistas.com"
BONISTAS_PAGE_TEMPLATE = "https://bonistas.com/bono-cotizacion-rendimiento-precio-hoy/{ticker}"

BONISTAS_DATA_TEMPLATE = (
    "https://bonistas.com/_next/data/{build_id}/"
    "bono-cotizacion-rendimiento-precio-hoy/{ticker}.json?bondId={ticker}"
)

REQUEST_SLEEP_SECONDS = 0.15
TIMEOUT = 25

BONISTAS_SHEETS = [
    "LECAPS",
    "BONCER",
    "DUALES",
    "DUALESCER",
    "TAMAR",
    "DLK",
    "BOTE",
    "HD_Sob",
]

# Tickers que se llaman distinto en Alquimia vs Bonistas.
# Clave: como viene en Alquimia. Valor: como hay que pedirlo a Bonistas.
TICKER_MAP_ALQUIMIA_TO_BONISTAS = {
    "TY30PPUT": "TY30P_PUT",
}


# =============================================================================
# HTTP / BONISTAS
# =============================================================================

def get_session():
    session = requests.Session()
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0 Safari/537.36"
        ),
        "Accept": "application/json,text/html,*/*",
    })
    return session


def get_build_id(session):
    urls_to_try = [
        BONISTAS_HOME,
        BONISTAS_PAGE_TEMPLATE.format(ticker="AL30D"),
        BONISTAS_PAGE_TEMPLATE.format(ticker="GD30D"),
    ]

    for url in urls_to_try:
        print(f"Buscando buildId en: {url}")
        response = session.get(url, timeout=TIMEOUT)
        response.raise_for_status()
        html = response.text

        match = re.search(r'"buildId"\s*:\s*"([^"]+)"', html)
        if match:
            return match.group(1)

        match = re.search(r"/_next/static/([^/]+)/", html)
        if match:
            return match.group(1)

    raise RuntimeError("No pude detectar el buildId de Bonistas.")


def extract_bond_data(json_data):
    try:
        return json_data["pageProps"]["bondData"]["bond"] or {}
    except Exception:
        return {}


def alquimia_to_bonistas_ticker(ticker_alquimia: str) -> str:
    """Traduce un ticker de Alquimia al nombre que usa Bonistas, si está mapeado."""
    return TICKER_MAP_ALQUIMIA_TO_BONISTAS.get(ticker_alquimia, ticker_alquimia)


def fetch_bonistas_ticker(session, build_id, ticker_alquimia):
    """
    Consulta Bonistas usando el nombre traducido.
    En el resultado devuelve el ticker ORIGINAL de Alquimia, para que el merge
    contra el Excel de Alquimia funcione en PriceComparison.
    """
    ticker_alquimia = str(ticker_alquimia).strip().upper()
    ticker_bonistas = alquimia_to_bonistas_ticker(ticker_alquimia)
    ticker_url = quote(ticker_bonistas, safe="")

    url = BONISTAS_DATA_TEMPLATE.format(
        build_id=build_id,
        ticker=ticker_url,
    )

    try:
        response = session.get(url, timeout=TIMEOUT)

        if response.status_code == 404:
            return {
                "ticker": ticker_alquimia,
                "ticker_bonistas": ticker_bonistas,
                "precio_bonistas": None,
                "status": "NO_ENCONTRADO",
                "error": "404",
                "url": url,
            }

        response.raise_for_status()
        data = response.json()
        bond = extract_bond_data(data)

        if not bond:
            return {
                "ticker": ticker_alquimia,
                "ticker_bonistas": ticker_bonistas,
                "precio_bonistas": None,
                "status": "SIN_BOND_DATA",
                "error": "No existe pageProps.bondData.bond",
                "url": url,
            }

        return {
            # Dejamos el ticker como viene de Alquimia para el merge posterior
            "ticker": ticker_alquimia,
            "ticker_bonistas": ticker_bonistas,
            "precio_bonistas": bond.get("last_price"),
            "last_close": bond.get("last_close"),
            "last_open": bond.get("last_open"),
            "last_min": bond.get("last_min"),
            "last_max": bond.get("last_max"),
            "tir_bonistas": bond.get("tir"),
            "md_bonistas": bond.get("modified_duration"),
            "paridad_bonistas": bond.get("parity"),
            "volumen_bonistas": bond.get("volume"),
            "settlement": bond.get("settlement"),
            "familia": bond.get("bond_family"),
            "familia_label": bond.get("bond_family_label"),
            "vencimiento": bond.get("end_date"),
            "status": "OK",
            "error": None,
            "url": url,
        }

    except Exception as e:
        return {
            "ticker": ticker_alquimia,
            "ticker_bonistas": ticker_bonistas,
            "precio_bonistas": None,
            "status": "ERROR",
            "error": str(e),
            "url": url,
        }


# =============================================================================
# MAIN
# =============================================================================

def main():
    print("========================================")
    print("GET PRICES BONISTAS")
    print(f"Hojas a consultar: {BONISTAS_SHEETS}")
    print("========================================")

    print(f"PROJECT_DIR: {PROJECT_DIR}")
    print(f"TICKERS_INPUT_PATH: {TICKERS_INPUT_PATH}")
    print(f"BONISTAS_OUTPUT_PATH: {BONISTAS_OUTPUT_PATH}")

    if not TICKERS_INPUT_PATH.exists():
        raise FileNotFoundError(
            f"No encontré {TICKERS_INPUT_PATH}. "
            "Corré primero GetPricesAlquimia.py."
        )

    tickers_df = pd.read_excel(
        TICKERS_INPUT_PATH,
        sheet_name="TICKERS_CIERRE",
        engine="openpyxl",
    )

    required_cols = {"ticker", "hoja_origen"}
    missing = required_cols - set(tickers_df.columns)
    if missing:
        raise ValueError(f"Faltan columnas en tickers_cierre.xlsx: {missing}")

    tickers_df["ticker"] = (
        tickers_df["ticker"].astype(str).str.strip().str.upper()
    )
    tickers_df["hoja_origen"] = (
        tickers_df["hoja_origen"].astype(str).str.strip()
    )

    print("\nHojas detectadas en tickers_cierre.xlsx:")
    print(sorted(tickers_df["hoja_origen"].dropna().unique().tolist()))

    bonos_df = tickers_df[tickers_df["hoja_origen"].isin(BONISTAS_SHEETS)].copy()

    print(f"\nTotal filas en tickers_cierre.xlsx: {len(tickers_df)}")
    print(f"Filas después de filtrar por hojas de bonos: {len(bonos_df)}")

    if bonos_df.empty:
        print("\n[WARN] No encontré tickers de bonos en tickers_cierre.xlsx.")
        print("Revisá que GetPricesAlquimia.py esté extrayendo las hojas:")
        print(BONISTAS_SHEETS)

        bonistas_df = pd.DataFrame(
            columns=["ticker", "ticker_bonistas", "precio_bonistas", "status", "error", "url"]
        )

        with pd.ExcelWriter(BONISTAS_OUTPUT_PATH, engine="openpyxl") as writer:
            bonistas_df.to_excel(writer, sheet_name="BONISTAS_RAW", index=False)

        print(f"\nExcel vacío generado: {BONISTAS_OUTPUT_PATH}")
        return

    tickers = (
        bonos_df["ticker"].dropna().drop_duplicates().sort_values().tolist()
    )

    print(f"Tickers únicos a consultar: {len(tickers)}")

    if TICKER_MAP_ALQUIMIA_TO_BONISTAS:
        print("\nMapeos Alquimia -> Bonistas activos:")
        for k, v in TICKER_MAP_ALQUIMIA_TO_BONISTAS.items():
            print(f"  {k} -> {v}")

    session = get_session()

    print("\nDetectando buildId de Bonistas...")
    build_id = get_build_id(session)
    print(f"buildId detectado: {build_id}")

    rows = []

    print("\nConsultando tickers en Bonistas...")

    for i, ticker in enumerate(tickers, start=1):
        ticker_bonistas = alquimia_to_bonistas_ticker(ticker)
        if ticker != ticker_bonistas:
            print(f"[{i}/{len(tickers)}] {ticker} -> {ticker_bonistas}")
        else:
            print(f"[{i}/{len(tickers)}] {ticker}")

        row = fetch_bonistas_ticker(session, build_id, ticker)
        rows.append(row)
        time.sleep(REQUEST_SLEEP_SECONDS)

    bonistas_df = pd.DataFrame(rows)

    if "precio_bonistas" in bonistas_df.columns:
        bonistas_df["precio_bonistas"] = pd.to_numeric(
            bonistas_df["precio_bonistas"], errors="coerce"
        )

    with pd.ExcelWriter(BONISTAS_OUTPUT_PATH, engine="openpyxl") as writer:
        bonistas_df.to_excel(writer, sheet_name="BONISTAS_RAW", index=False)

    print("\n========================================")
    print("PROCESO TERMINADO")
    print("========================================")
    print(f"Excel generado: {BONISTAS_OUTPUT_PATH}")

    if "status" in bonistas_df.columns:
        print("\nResumen status:")
        print(bonistas_df["status"].value_counts(dropna=False).to_string())


if __name__ == "__main__":
    main()