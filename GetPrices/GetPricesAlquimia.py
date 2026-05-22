import io
import re
import os
import sys
from pathlib import Path

import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload


# =============================================================================
# PATHS
# =============================================================================

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = SCRIPT_DIR.parent

SERVICE_ACCOUNT_FILE = Path(
    os.getenv("SERVICE_ACCOUNT_FILE", PROJECT_DIR / "service_account.json")
)

OUTPUT_DIR = SCRIPT_DIR / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

LOCAL_CIERRE_PATH = OUTPUT_DIR / "ultimo_cierre_drive.xlsx"
TICKERS_OUTPUT_PATH = OUTPUT_DIR / "tickers_cierre.xlsx"

NOMBRE_ARCHIVO_DRIVE = "ultimo_cierre.xlsx"


# =============================================================================
# CONFIG EXCEL
# =============================================================================

# Hojas que SÍ leemos del Excel de Alquimia.
# CEDEAR queda AFUERA: no se compara contra ninguna fuente externa.
SHEETS_TO_READ = [
    # Bonos AR (comparan contra Bonistas)
    "BONCER",
    "DLK",
    "TAMAR",
    "DUALES",
    "DUALESCER",
    "BOTE",
    "HD_Sob",
    "LECAPS",
    # Yahoo (ADR + NYSE whitelist + INDICES + US BONDS + COMMODITIES)
    "ADR",
    "NYSE",
    "INDICES US",
    "INDICES AM",
    "INDICES EU",
    "INDICES ASIA",
    "US BONDS",
    "COMMODITIES",
]

TICKER_CANDIDATES = [
    "Ticker",
    "ticker",
    "Especie",
    "especie",
    "Symbol",
    "symbol",
]

PRICE_CANDIDATES_BY_SHEET = {
    # Bonos
    "BONCER": ["Precio", "precio"],
    "DLK": ["Precio ARS", "Precio", "precio_ars", "precio"],
    "TAMAR": ["Precio", "precio"],
    "DUALES": ["Precio", "precio"],
    "DUALESCER": ["Precio", "precio"],
    "BOTE": ["Precio Sucio", "Precio", "precio_sucio", "precio"],
    "HD_Sob": ["Precio", "precio"],
    "LECAPS": ["Precio Dirty", "Precio", "precio"],
    # Yahoo
    "ADR": ["Precio USD", "Precio", "precio_usd", "precio"],
    "NYSE": ["Precio USD", "Precio", "precio_usd", "precio"],
    "INDICES US": ["Precio", "precio"],
    "INDICES AM": ["Precio", "precio"],
    "INDICES EU": ["Precio", "precio"],
    "INDICES ASIA": ["Precio", "precio"],
    "US BONDS": ["Precio", "precio"],
    "COMMODITIES": ["Precio", "precio"],
}

# Whitelist NYSE: solo estos tickers entran a tickers_cierre.xlsx.
# El resto que aparezca en la hoja NYSE del Excel se descarta acá mismo.
NYSE_WHITELIST = {
    # 1/3
    "AAPL", "ABEV", "ADBE", "AMD", "AMZN", "BABA", "BAC", "BBD",
    "BRK-B", "BRK.B", "BRKB",
    "BSBR", "CIG", "COST", "CRM", "CSCO", "CVX", "DIS",
    "EMBJ", "GGB", "GOOG", "GOOGL", "HD", "IBM", "INTC", "ITUB",
    "JNJ", "JPM",
    # 2/3
    "KO", "MA", "MCD", "META", "MRK", "MSFT", "NFLX", "NKE", "NU",
    "NVDA", "ORCL", "PAGS", "PBR", "PEP", "PFE", "PG", "QCOM",
    "SBUX", "SID", "STNE", "T", "TIMB", "TME", "TSLA", "TSM", "UNH",
    # 3/3
    "V", "VALE", "VIV", "VZ", "WFC", "WMT", "XOM", "XP",
}

SHEET_TICKER_WHITELISTS = {
    "NYSE": NYSE_WHITELIST,
}


# =============================================================================
# GOOGLE DRIVE
# =============================================================================

def get_drive_service():
    if not SERVICE_ACCOUNT_FILE.exists():
        raise FileNotFoundError(
            f"No encontré el service account en: {SERVICE_ACCOUNT_FILE}"
        )

    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=["https://www.googleapis.com/auth/drive.readonly"],
    )

    return build("drive", "v3", credentials=creds)


def find_latest_ultimo_cierre(service):
    query = (
        f"name = '{NOMBRE_ARCHIVO_DRIVE}' "
        f"and trashed = false"
    )

    response = service.files().list(
        q=query,
        fields="files(id, name, modifiedTime, size, owners(displayName, emailAddress), parents)",
        orderBy="modifiedTime desc",
        pageSize=20,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()

    files = response.get("files", [])

    if not files:
        raise FileNotFoundError(
            f"No encontré ningún archivo llamado {NOMBRE_ARCHIVO_DRIVE} "
            "accesible para este service account."
        )

    print("\nArchivos encontrados en Drive:")
    for i, f in enumerate(files, start=1):
        owners = f.get("owners") or []
        owner = ""
        if owners:
            owner = owners[0].get("displayName") or owners[0].get("emailAddress") or ""

        print(
            f"{i}. {f.get('name')} | "
            f"modifiedTime={f.get('modifiedTime')} | "
            f"size={f.get('size')} | "
            f"owner={owner} | "
            f"id={f.get('id')}"
        )

    latest = files[0]

    print("\nArchivo elegido:")
    print(
        f"{latest.get('name')} | "
        f"modifiedTime={latest.get('modifiedTime')} | "
        f"size={latest.get('size')} | "
        f"id={latest.get('id')}"
    )

    return latest


def download_drive_file(service, file_id, output_path):
    request = service.files().get_media(
        fileId=file_id,
        supportsAllDrives=True,
    )

    with io.FileIO(output_path, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)

        done = False
        while not done:
            status, done = downloader.next_chunk()
            if status:
                print(f"Descargando cierre... {int(status.progress() * 100)}%")

    return output_path


# =============================================================================
# EXCEL READER
# =============================================================================

def normalize_ticker(value):
    if pd.isna(value):
        return None

    ticker = str(value).strip().upper()
    ticker = re.sub(r"\s+", "", ticker)

    if ticker in {"", "NAN", "NONE", "NULL"}:
        return None

    return ticker


def find_header_row(raw_df):
    max_rows = min(len(raw_df), 25)

    for idx in range(max_rows):
        row_values = [str(x).strip() for x in raw_df.iloc[idx].tolist()]
        if any(x in TICKER_CANDIDATES for x in row_values):
            return idx

    return None


def pick_column(df, candidates):
    normalized_map = {str(c).strip(): c for c in df.columns}

    for candidate in candidates:
        if candidate in normalized_map:
            return normalized_map[candidate]

    return None


def read_sheet_tickers(excel_path, sheet_name):
    raw = pd.read_excel(
        excel_path,
        sheet_name=sheet_name,
        header=None,
        engine="openpyxl",
    )

    header_row = find_header_row(raw)

    if header_row is None:
        print(f"[WARN] No encontré header en hoja {sheet_name}. La salteo.")
        return pd.DataFrame()

    df = pd.read_excel(
        excel_path,
        sheet_name=sheet_name,
        header=header_row,
        engine="openpyxl",
    )

    df = df.dropna(how="all")

    ticker_col = pick_column(df, TICKER_CANDIDATES)
    price_col = pick_column(
        df,
        PRICE_CANDIDATES_BY_SHEET.get(sheet_name, ["Precio", "precio"]),
    )

    if ticker_col is None:
        print(f"[WARN] No encontré columna de ticker en {sheet_name}.")
        print(f"Columnas disponibles: {list(df.columns)}")
        return pd.DataFrame()

    if price_col is None:
        print(f"[WARN] No encontré columna de precio en {sheet_name}.")
        print(f"Columnas disponibles: {list(df.columns)}")
        return pd.DataFrame()

    out = df[[ticker_col, price_col]].copy()
    out.columns = ["ticker", "precio_alquimia"]

    out["ticker"] = out["ticker"].apply(normalize_ticker)
    out["precio_alquimia"] = pd.to_numeric(out["precio_alquimia"], errors="coerce")
    out["hoja_origen"] = sheet_name

    out = out.dropna(subset=["ticker"])
    out = out[out["precio_alquimia"].notna()]

    # Aplicar whitelist si corresponde
    if sheet_name in SHEET_TICKER_WHITELISTS:
        whitelist = SHEET_TICKER_WHITELISTS[sheet_name]
        before = len(out)
        out = out[out["ticker"].isin(whitelist)].copy()
        after = len(out)
        print(
            f"   [Whitelist {sheet_name}] "
            f"Antes: {before} | Después: {after} | Excluidos: {before - after}"
        )

    return out[["hoja_origen", "ticker", "precio_alquimia"]]


# =============================================================================
# MAIN
# =============================================================================

def main():
    print("========================================")
    print("GET PRICES ALQUIMIA")
    print("========================================")

    print(f"PROJECT_DIR: {PROJECT_DIR}")
    print(f"SERVICE_ACCOUNT_FILE: {SERVICE_ACCOUNT_FILE}")
    print(f"OUTPUT_DIR: {OUTPUT_DIR}")

    service = get_drive_service()

    print("\nBuscando último ultimo_cierre.xlsx accesible...")
    latest_file = find_latest_ultimo_cierre(service)

    print("\nDescargando archivo elegido...")
    download_drive_file(
        service=service,
        file_id=latest_file["id"],
        output_path=LOCAL_CIERRE_PATH,
    )

    print(f"\nArchivo descargado en: {LOCAL_CIERRE_PATH}")

    xl = pd.ExcelFile(LOCAL_CIERRE_PATH, engine="openpyxl")
    existing_sheets = set(xl.sheet_names)

    print("\nSolapas encontradas:")
    print(xl.sheet_names)

    all_rows = []

    print("\nLeyendo solapas objetivo...")

    for sheet in SHEETS_TO_READ:
        if sheet not in existing_sheets:
            print(f"[WARN] No existe la hoja {sheet}. La salteo.")
            continue

        rows = read_sheet_tickers(LOCAL_CIERRE_PATH, sheet)

        if rows.empty:
            print(f"[WARN] {sheet}: sin tickers válidos.")
            continue

        print(f"[OK] {sheet}: {len(rows)} tickers.")
        all_rows.append(rows)

    if not all_rows:
        raise RuntimeError("No se extrajo ningún ticker de las solapas indicadas.")

    final = pd.concat(all_rows, ignore_index=True)

    final = final.drop_duplicates(
        subset=["hoja_origen", "ticker"],
        keep="first",
    )

    final = final.sort_values(["hoja_origen", "ticker"])

    with pd.ExcelWriter(TICKERS_OUTPUT_PATH, engine="openpyxl") as writer:
        final.to_excel(writer, sheet_name="TICKERS_CIERRE", index=False)

    print("\n========================================")
    print("PROCESO TERMINADO")
    print("========================================")
    print(f"Excel generado: {TICKERS_OUTPUT_PATH}")
    print(f"Total filas: {len(final)}")
    print(f"Total tickers únicos: {final['ticker'].nunique()}")

    print("\nResumen por hoja:")
    print(
        final.groupby("hoja_origen")
        .size()
        .reset_index(name="cantidad")
        .to_string(index=False)
    )


if __name__ == "__main__":
    main()