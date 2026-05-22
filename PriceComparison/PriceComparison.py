import shutil
from pathlib import Path

import pandas as pd


# =============================================================================
# PATHS
# =============================================================================

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = SCRIPT_DIR.parent

ALQUIMIA_INPUT_PATH = PROJECT_DIR / "GetPrices" / "output" / "tickers_cierre.xlsx"
BONISTAS_INPUT_PATH = PROJECT_DIR / "GetPrices" / "output" / "bonistas_precios.xlsx"
YAHOO_INPUT_PATH = PROJECT_DIR / "GetPrices" / "output" / "yahoo_precios.xlsx"

OUTPUT_DIR = SCRIPT_DIR / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

COMPARACION_OUTPUT_PATH = OUTPUT_DIR / "comparacion_precios.xlsx"


# =============================================================================
# CONFIG
# =============================================================================

TOLERANCIA_BONOS_DIF_PCT = 0.005
TOLERANCIA_BONOS_DIF_ABS = 0.10

TOLERANCIA_YAHOO_DIF_PCT = 0.005
TOLERANCIA_YAHOO_DIF_ABS = 0.03

YAHOO_SHEETS = [
    "ADR",
    "NYSE",
    "CEDEAR",
    "INDICES US",
    "INDICES AM",
    "INDICES EU",
    "INDICES ASIA",
    "US BONDS",
    "COMMODITIES",
]

SHEET_ALIASES = {
    "ADR": "ADR",
    "ADRS": "ADR",

    "NYSE": "NYSE",

    "CEDEAR": "CEDEAR",
    "CEDEARS": "CEDEAR",

    "INDICES US": "INDICES US",
    "INDICES_US": "INDICES US",
    "ÍNDICES US": "INDICES US",
    "ÍNDICES_US": "INDICES US",

    "INDICES AM": "INDICES AM",
    "INDICES_AM": "INDICES AM",
    "ÍNDICES AM": "INDICES AM",
    "ÍNDICES_AM": "INDICES AM",

    "INDICES EU": "INDICES EU",
    "INDICES_EU": "INDICES EU",
    "ÍNDICES EU": "INDICES EU",
    "ÍNDICES_EU": "INDICES EU",

    "INDICES ASIA": "INDICES ASIA",
    "INDICES_ASIA": "INDICES ASIA",
    "ÍNDICES ASIA": "INDICES ASIA",
    "ÍNDICES_ASIA": "INDICES ASIA",

    "US BONDS": "US BONDS",
    "US_BONDS": "US BONDS",

    "COMMODITIES": "COMMODITIES",
}


# =============================================================================
# HELPERS
# =============================================================================

def normalize_ticker_series(series):
    return (
        series
        .astype(str)
        .str.strip()
        .str.upper()
    )


def normalize_sheet_name(value):
    raw = str(value).strip().upper()
    raw = raw.replace("-", " ")
    raw = " ".join(raw.split())
    raw_underscore = raw.replace(" ", "_")

    if raw in SHEET_ALIASES:
        return SHEET_ALIASES[raw]

    if raw_underscore in SHEET_ALIASES:
        return SHEET_ALIASES[raw_underscore]

    return raw


def normalize_sheet_series(series):
    return series.apply(normalize_sheet_name)


def classify_bonos(row):
    status = row.get("status")

    if status != "OK":
        return status if pd.notna(status) else "NO_CONSULTADO"

    if pd.isna(row.get("precio_alquimia")):
        return "SIN_PRECIO_ALQUIMIA"

    if pd.isna(row.get("precio_bonistas")):
        return "SIN_PRECIO_BONISTAS"

    dif_abs = abs(row.get("dif_abs"))
    dif_pct = abs(row.get("dif_pct"))

    if dif_abs > TOLERANCIA_BONOS_DIF_ABS and dif_pct > TOLERANCIA_BONOS_DIF_PCT:
        return "ALERTA"

    return "OK"


def classify_yahoo(row):
    status = row.get("status_yahoo")

    if status != "OK":
        return status if pd.notna(status) else "NO_CONSULTADO"

    if pd.isna(row.get("precio_alquimia")):
        return "SIN_PRECIO_ALQUIMIA"

    if pd.isna(row.get("precio_yahoo")):
        return "SIN_PRECIO_YAHOO"

    dif_abs = abs(row.get("dif_abs"))
    dif_pct = abs(row.get("dif_pct"))

    if dif_abs > TOLERANCIA_YAHOO_DIF_ABS and dif_pct > TOLERANCIA_YAHOO_DIF_PCT:
        return "ALERTA"

    return "OK"


def autosize_columns(writer, sheet_name, df):
    worksheet = writer.sheets[sheet_name]

    for idx, col in enumerate(df.columns, start=1):
        series = df[col].astype(str)

        max_len = max(
            [len(str(col))]
            + [len(x) for x in series.head(500).tolist()]
        )

        adjusted_width = min(max(max_len + 2, 10), 35)

        worksheet.column_dimensions[
            worksheet.cell(row=1, column=idx).column_letter
        ].width = adjusted_width


def check_output_writable(path: Path):
    """
    Chequea que podamos escribir el archivo de salida.
    Si está abierto en Excel/OneDrive, falla con un mensaje claro
    en lugar de un traceback de openpyxl.
    """
    if not path.exists():
        return

    try:
        # Abrir en modo append binario es no destructivo pero falla si está lockeado
        with open(path, "a+b"):
            pass
    except PermissionError:
        raise PermissionError(
            f"\n\n"
            f"========================================\n"
            f"NO PUEDO ESCRIBIR EL ARCHIVO DE SALIDA\n"
            f"========================================\n"
            f"Path: {path}\n"
            f"Probablemente lo tenés abierto en Excel.\n"
            f"Cerralo y volvé a correr el script.\n"
            f"========================================\n"
        )


def write_excel_safely(path: Path, write_fn):
    """
    Escribe el Excel a un archivo temporal y después lo renombra al destino final.
    Si el destino está lockeado, al menos queda el .tmp con los datos del run.

    write_fn recibe el ExcelWriter abierto y se encarga de escribir las hojas.
    """
    tmp_path = path.with_suffix(path.suffix + ".tmp")

    with pd.ExcelWriter(tmp_path, engine="openpyxl") as writer:
        write_fn(writer)

    try:
        # shutil.move sobrescribe si el destino existe y está libre
        shutil.move(str(tmp_path), str(path))
    except PermissionError:
        raise PermissionError(
            f"\n\n"
            f"========================================\n"
            f"NO PUDE REEMPLAZAR EL ARCHIVO FINAL\n"
            f"========================================\n"
            f"Generé los datos en: {tmp_path}\n"
            f"Pero no pude moverlos a: {path}\n"
            f"(¿lo abriste en Excel mientras corría el script?)\n"
            f"Cerralo y renombrá manualmente el .tmp, o volvé a correr.\n"
            f"========================================\n"
        )


# =============================================================================
# COMPARACIONES
# =============================================================================

def build_bonos_comparison(alquimia):
    if not BONISTAS_INPUT_PATH.exists():
        raise FileNotFoundError(
            f"No encontré {BONISTAS_INPUT_PATH}. "
            "Corré primero GetPricesBonistas.py."
        )

    bonistas = pd.read_excel(
        BONISTAS_INPUT_PATH,
        sheet_name="BONISTAS_RAW",
        engine="openpyxl",
    )

    bonos = alquimia[
        ~alquimia["hoja_origen"].isin(YAHOO_SHEETS)
    ].copy()

    if bonos.empty:
        return pd.DataFrame(
            columns=[
                "hoja_origen",
                "ticker",
                "precio_alquimia",
                "precio_bonistas",
                "dif_abs",
                "dif_pct",
                "estado",
            ]
        )

    required_bonistas = {"ticker", "precio_bonistas", "status"}
    missing_bonistas = required_bonistas - set(bonistas.columns)

    if missing_bonistas:
        raise ValueError(f"Faltan columnas en Bonistas: {missing_bonistas}")

    bonistas = bonistas[["ticker", "precio_bonistas", "status"]].copy()

    bonos["ticker"] = normalize_ticker_series(bonos["ticker"])
    bonistas["ticker"] = normalize_ticker_series(bonistas["ticker"])

    bonos["precio_alquimia"] = pd.to_numeric(
        bonos["precio_alquimia"],
        errors="coerce",
    )

    bonistas["precio_bonistas"] = pd.to_numeric(
        bonistas["precio_bonistas"],
        errors="coerce",
    )

    bonistas = bonistas.drop_duplicates(subset=["ticker"], keep="first")

    comp = bonos.merge(
        bonistas,
        on="ticker",
        how="left",
    )

    comp["dif_abs"] = comp["precio_bonistas"] - comp["precio_alquimia"]

    comp["dif_pct"] = (
        comp["precio_bonistas"] / comp["precio_alquimia"] - 1
    )

    comp["estado"] = comp.apply(classify_bonos, axis=1)

    comp = comp[
        [
            "hoja_origen",
            "ticker",
            "precio_alquimia",
            "precio_bonistas",
            "dif_abs",
            "dif_pct",
            "estado",
        ]
    ]

    return comp


def build_yahoo_comparison(alquimia):
    if not YAHOO_INPUT_PATH.exists():
        raise FileNotFoundError(
            f"No encontré {YAHOO_INPUT_PATH}. "
            "Corré primero GetPricesYahoo.py."
        )

    yahoo = pd.read_excel(
        YAHOO_INPUT_PATH,
        sheet_name="YAHOO_RAW",
        engine="openpyxl",
    )

    yahoo_assets = alquimia[
        alquimia["hoja_origen"].isin(YAHOO_SHEETS)
    ].copy()

    if yahoo_assets.empty:
        return pd.DataFrame(
            columns=[
                "hoja_origen",
                "ticker",
                "precio_alquimia",
                "precio_yahoo",
                "dif_abs",
                "dif_pct",
                "estado",
            ]
        )

    required_yahoo = {
        "hoja_origen",
        "ticker",
        "precio_yahoo",
        "status_yahoo",
    }

    missing_yahoo = required_yahoo - set(yahoo.columns)

    if missing_yahoo:
        raise ValueError(f"Faltan columnas en Yahoo: {missing_yahoo}")

    yahoo = yahoo[
        [
            "hoja_origen",
            "ticker",
            "precio_yahoo",
            "status_yahoo",
        ]
    ].copy()

    yahoo["hoja_origen"] = normalize_sheet_series(yahoo["hoja_origen"])
    yahoo_assets["hoja_origen"] = normalize_sheet_series(yahoo_assets["hoja_origen"])

    yahoo_assets["ticker"] = normalize_ticker_series(yahoo_assets["ticker"])
    yahoo["ticker"] = normalize_ticker_series(yahoo["ticker"])

    yahoo_assets["precio_alquimia"] = pd.to_numeric(
        yahoo_assets["precio_alquimia"],
        errors="coerce",
    )

    yahoo["precio_yahoo"] = pd.to_numeric(
        yahoo["precio_yahoo"],
        errors="coerce",
    )

    yahoo = yahoo.drop_duplicates(
        subset=["hoja_origen", "ticker"],
        keep="first",
    )

    comp = yahoo_assets.merge(
        yahoo,
        on=["hoja_origen", "ticker"],
        how="left",
    )

    comp["dif_abs"] = comp["precio_yahoo"] - comp["precio_alquimia"]

    comp["dif_pct"] = (
        comp["precio_yahoo"] / comp["precio_alquimia"] - 1
    )

    comp["estado"] = comp.apply(classify_yahoo, axis=1)

    comp = comp[
        [
            "hoja_origen",
            "ticker",
            "precio_alquimia",
            "precio_yahoo",
            "dif_abs",
            "dif_pct",
            "estado",
        ]
    ]

    return comp


# =============================================================================
# MAIN
# =============================================================================

def main():
    print("========================================")
    print("PRICE COMPARISON")
    print("========================================")

    # Chequeo upfront: si el archivo de salida está abierto en Excel,
    # avisamos antes de hacer todo el trabajo.
    check_output_writable(COMPARACION_OUTPUT_PATH)

    if not ALQUIMIA_INPUT_PATH.exists():
        raise FileNotFoundError(
            f"No encontré {ALQUIMIA_INPUT_PATH}. "
            "Corré primero GetPricesAlquimia.py."
        )

    alquimia = pd.read_excel(
        ALQUIMIA_INPUT_PATH,
        sheet_name="TICKERS_CIERRE",
        engine="openpyxl",
    )

    if "precio_alquimia" not in alquimia.columns:
        if "precio_nuestro" in alquimia.columns:
            alquimia = alquimia.rename(
                columns={"precio_nuestro": "precio_alquimia"}
            )
        else:
            raise ValueError(
                "El archivo de Alquimia no tiene columna 'precio_alquimia' "
                "ni 'precio_nuestro'."
            )

    required_alquimia = {"hoja_origen", "ticker", "precio_alquimia"}
    missing_alquimia = required_alquimia - set(alquimia.columns)

    if missing_alquimia:
        raise ValueError(f"Faltan columnas en Alquimia: {missing_alquimia}")

    alquimia = alquimia[
        [
            "hoja_origen",
            "ticker",
            "precio_alquimia",
        ]
    ].copy()

    alquimia["hoja_origen"] = normalize_sheet_series(alquimia["hoja_origen"])
    alquimia["ticker"] = normalize_ticker_series(alquimia["ticker"])

    alquimia["precio_alquimia"] = pd.to_numeric(
        alquimia["precio_alquimia"],
        errors="coerce",
    )

    bonos_comp = build_bonos_comparison(alquimia)
    yahoo_comp = build_yahoo_comparison(alquimia)

    # Filtrar DataFrames vacíos antes del concat (evita FutureWarning de pandas)
    frames = []
    if not bonos_comp.empty:
        frames.append(
            bonos_comp.rename(columns={"precio_bonistas": "precio_externo"})
        )
    if not yahoo_comp.empty:
        frames.append(
            yahoo_comp.rename(columns={"precio_yahoo": "precio_externo"})
        )

    if frames:
        total_comp = pd.concat(frames, ignore_index=True)
    else:
        total_comp = pd.DataFrame(
            columns=[
                "hoja_origen",
                "ticker",
                "precio_alquimia",
                "precio_externo",
                "dif_abs",
                "dif_pct",
                "estado",
            ]
        )

    resumen_estado = (
        total_comp
        .groupby("estado", dropna=False)
        .size()
        .reset_index(name="cantidad")
        .sort_values("cantidad", ascending=False)
    )

    errores = total_comp[total_comp["estado"] != "OK"].copy()

    def write_all_sheets(writer):
        resumen_estado.to_excel(
            writer,
            sheet_name="RESUMEN_ESTADO",
            index=False,
        )

        errores.to_excel(
            writer,
            sheet_name="ERRORES",
            index=False,
        )

        bonos_comp.to_excel(
            writer,
            sheet_name="BONOS_BONISTAS",
            index=False,
        )

        yahoo_comp.to_excel(
            writer,
            sheet_name="YAHOO",
            index=False,
        )

        total_comp.to_excel(
            writer,
            sheet_name="COMPARACION_TOTAL",
            index=False,
        )

        sheets_to_resize = {
            "RESUMEN_ESTADO": resumen_estado,
            "ERRORES": errores,
            "BONOS_BONISTAS": bonos_comp,
            "YAHOO": yahoo_comp,
            "COMPARACION_TOTAL": total_comp,
        }

        for sheet_name, df_resize in sheets_to_resize.items():
            autosize_columns(writer, sheet_name, df_resize)

    write_excel_safely(COMPARACION_OUTPUT_PATH, write_all_sheets)

    print("\n========================================")
    print("PROCESO TERMINADO")
    print("========================================")
    print(f"Excel generado: {COMPARACION_OUTPUT_PATH}")

    print("\nResumen por estado:")
    print(resumen_estado.to_string(index=False))


if __name__ == "__main__":
    main()