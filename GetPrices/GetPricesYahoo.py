from pathlib import Path

import pandas as pd
import yfinance as yf


# =============================================================================
# PATHS
# =============================================================================

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = SCRIPT_DIR.parent

INPUT_DIR = SCRIPT_DIR / "output"
OUTPUT_DIR = SCRIPT_DIR / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

TICKERS_INPUT_PATH = INPUT_DIR / "tickers_cierre.xlsx"
YAHOO_OUTPUT_PATH = OUTPUT_DIR / "yahoo_precios.xlsx"


# =============================================================================
# CONFIG
# =============================================================================

# Hojas que SÍ vamos a consultar en Yahoo Finance.
# OJO: CEDEAR queda afuera a propósito (no comparamos vs Yahoo).
YAHOO_SHEETS_CANONICAL = [
    "ADR",
    "NYSE",
    "INDICES US",
    "INDICES AM",
    "INDICES EU",
    "INDICES ASIA",
    "US BONDS",
    "COMMODITIES",
]

# Whitelist NYSE: solo estos tickers se consultan en Yahoo.
# El resto que aparezca en la hoja NYSE del Excel se IGNORA.
# Lista tomada de las slides "NYSE - US LARGE CAPS" (1/3, 2/3, 3/3).
NYSE_WHITELIST = {
    # 1/3
    "AAPL", "ABEV", "ADBE", "AMD", "AMZN", "BABA", "BAC", "BBD",
    "BRK-B", "BSBR", "CIG", "COST", "CRM", "CSCO", "CVX", "DIS",
    "EMBJ", "GGB", "GOOG", "GOOGL", "HD", "IBM", "INTC", "ITUB",
    "JNJ", "JPM",
    # 2/3
    "KO", "MA", "MCD", "META", "MRK", "MSFT", "NFLX", "NKE", "NU",
    "NVDA", "ORCL", "PAGS", "PBR", "PEP", "PFE", "PG", "QCOM",
    "SBUX", "SID", "STNE", "T", "TIMB", "TME", "TSLA", "TSM", "UNH",
    # 3/3
    "V", "VALE", "VIV", "VZ", "WFC", "WMT", "XOM", "XP",
    # Variantes equivalentes de BRK-B (por si vienen sin guión en el Excel)
    "BRK.B", "BRKB",
}

# Hojas en las que filtramos por whitelist.
# Si una hoja está acá, solo se consultan los tickers presentes en su whitelist.
SHEET_WHITELISTS = {
    "NYSE": NYSE_WHITELIST,
}

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

# Hojas que requieren sufijo .BA para Yahoo (mercado argentino).
# CEDEAR quedó fuera de YAHOO_SHEETS_CANONICAL, así que esto ya no aplica,
# pero lo dejo por consistencia y por si en el futuro se vuelve a sumar.
ADD_BA_SUFFIX_SHEETS = ["CEDEAR"]


# =============================================================================
# HELPERS
# =============================================================================

def normalize_text(value):
    return str(value).strip().upper()


def normalize_sheet_name(value):
    raw = normalize_text(value)
    raw = raw.replace("-", " ")
    raw = " ".join(raw.split())
    raw_underscore = raw.replace(" ", "_")

    if raw in SHEET_ALIASES:
        return SHEET_ALIASES[raw]

    if raw_underscore in SHEET_ALIASES:
        return SHEET_ALIASES[raw_underscore]

    return raw


def yahoo_symbol_from_ticker(ticker: str, hoja_origen: str) -> str:
    """
    Devuelve el símbolo Yahoo para un ticker dado.

    Los tickers en INDICES, US BONDS y COMMODITIES del archivo de Alquimia
    YA vienen con sintaxis Yahoo (^VIX, ^N225, SI=F, 2YY=F, etc.).
    Para esos casos, el manual_map solo cubre eventuales variantes por nombre.

    Para US BONDS: los símbolos ^IRX, ^FVX, ^TNX, ^TYX y 2YY=F devuelven
    YIELDS directamente desde yfinance (no precios), que es lo que Alquimia
    también está exportando en su columna "Precio".
    """
    ticker = normalize_text(ticker)
    hoja_origen = normalize_sheet_name(hoja_origen)

    manual_map = {
        # Acciones / ADRs / NYSE
        "BRK.B": "BRK-B",
        "BRKB": "BRK-B",
        "BF.B": "BF-B",
        "BFB": "BF-B",

        # Índices (por si vienen como nombre en vez de símbolo Yahoo)
        "SPX": "^GSPC",
        "S&P500": "^GSPC",
        "S&P 500": "^GSPC",
        "NASDAQ": "^IXIC",
        "DOW": "^DJI",
        "DOW 30": "^DJI",
        "VIX": "^VIX",
        "RUSSELL2000": "^RUT",
        "RUSSELL 2000": "^RUT",

        # América
        "MERVAL": "^MERV",
        "MERV": "^MERV",
        "IBOV": "^BVSP",
        "IBOVESPA": "^BVSP",

        # Europa
        "FTSE": "^FTSE",
        "DAX": "^GDAXI",
        "CAC": "^FCHI",
        "CAC40": "^FCHI",
        "CAC 40": "^FCHI",
        "EUROSTOXX50": "^STOXX50E",
        "EURO STOXX 50": "^STOXX50E",
        "IBEX": "^IBEX",

        # Asia
        "NIKKEI": "^N225",
        "NIKKEI225": "^N225",
        "NIKKEI 225": "^N225",
        "HANGSENG": "^HSI",
        "HANG SENG": "^HSI",
        "KOSPI": "^KS11",
        "ASX": "^AXJO",
        "SSE": "000001.SS",
        "SSE COMPOSITE": "000001.SS",

        # US Bonds - yields
        "US 2Y": "2YY=F",
        "US2Y": "2YY=F",
        "US 5Y": "^FVX",
        "US5Y": "^FVX",
        "US 10Y": "^TNX",
        "US10Y": "^TNX",
        "US 30Y": "^TYX",
        "US30Y": "^TYX",
        "13-WK T-BILL": "^IRX",
        "13WK": "^IRX",

        # Commodities (por si vienen como nombre)
        "SILVER": "SI=F",
        "GOLD": "GC=F",
        "COPPER": "HG=F",
        "WTI": "CL=F",
        "BRENT": "BZ=F",
        "NATURAL GAS": "NG=F",
    }

    if ticker in manual_map:
        return manual_map[ticker]

    if hoja_origen in ADD_BA_SUFFIX_SHEETS:
        if ticker.endswith(".BA"):
            return ticker
        return f"{ticker}.BA"

    return ticker


def get_last_price_from_yahoo(symbol: str):
    ticker_obj = yf.Ticker(symbol)

    try:
        fast_info = ticker_obj.fast_info
        price = fast_info.get("last_price")

        if price is not None:
            return float(price)
    except Exception:
        pass

    try:
        hist = ticker_obj.history(
            period="5d",
            interval="1d",
            auto_adjust=False,
        )

        if hist is not None and not hist.empty:
            close = hist["Close"].dropna()
            if not close.empty:
                return float(close.iloc[-1])
    except Exception:
        pass

    return None


def fetch_yahoo_ticker(ticker: str, hoja_origen: str):
    hoja_canonica = normalize_sheet_name(hoja_origen)
    ticker_clean = normalize_text(ticker)

    yahoo_symbol = yahoo_symbol_from_ticker(
        ticker=ticker_clean,
        hoja_origen=hoja_canonica,
    )

    try:
        price = get_last_price_from_yahoo(yahoo_symbol)

        if price is None:
            return {
                "hoja_origen": hoja_canonica,
                "ticker": ticker_clean,
                "yahoo_symbol": yahoo_symbol,
                "precio_yahoo": None,
                "status_yahoo": "SIN_PRECIO",
                "error_yahoo": None,
            }

        return {
            "hoja_origen": hoja_canonica,
            "ticker": ticker_clean,
            "yahoo_symbol": yahoo_symbol,
            "precio_yahoo": price,
            "status_yahoo": "OK",
            "error_yahoo": None,
        }

    except Exception as e:
        return {
            "hoja_origen": hoja_canonica,
            "ticker": ticker_clean,
            "yahoo_symbol": yahoo_symbol,
            "precio_yahoo": None,
            "status_yahoo": "ERROR",
            "error_yahoo": str(e),
        }


# =============================================================================
# MAIN
# =============================================================================

def main():
    print("========================================")
    print("GET PRICES YAHOO")
    print("ADR + NYSE (whitelist) + INDICES + US BONDS + COMMODITIES")
    print("CEDEAR: NO se consulta")
    print("========================================")

    print(f"TICKERS_INPUT_PATH: {TICKERS_INPUT_PATH}")
    print(f"YAHOO_OUTPUT_PATH: {YAHOO_OUTPUT_PATH}")

    if not TICKERS_INPUT_PATH.exists():
        raise FileNotFoundError(
            f"No encontré {TICKERS_INPUT_PATH}. "
            "Corré primero GetPricesAlquimia.py."
        )

    df = pd.read_excel(
        TICKERS_INPUT_PATH,
        sheet_name="TICKERS_CIERRE",
        engine="openpyxl",
    )

    required_cols = {"hoja_origen", "ticker"}
    missing = required_cols - set(df.columns)

    if missing:
        raise ValueError(f"Faltan columnas en tickers_cierre.xlsx: {missing}")

    df["hoja_origen"] = df["hoja_origen"].apply(normalize_sheet_name)
    df["ticker"] = df["ticker"].apply(normalize_text)

    print("\nHojas detectadas en tickers_cierre.xlsx:")
    print(sorted(df["hoja_origen"].dropna().unique().tolist()))

    # Filtro 1: solo hojas Yahoo (excluye bonos AR y CEDEAR)
    yahoo_df = df[df["hoja_origen"].isin(YAHOO_SHEETS_CANONICAL)].copy()

    # Filtro 2: para hojas con whitelist (ej. NYSE), quedarse solo con los tickers de la whitelist
    if not yahoo_df.empty:
        keep_mask = pd.Series(True, index=yahoo_df.index)

        for sheet, whitelist in SHEET_WHITELISTS.items():
            in_sheet = yahoo_df["hoja_origen"] == sheet
            in_whitelist = yahoo_df["ticker"].isin(whitelist)
            # Mantengo si: (no es esta hoja) o (es esta hoja y está en la whitelist)
            keep_mask &= (~in_sheet) | in_whitelist

            n_excluidos = (in_sheet & ~in_whitelist).sum()
            n_incluidos = (in_sheet & in_whitelist).sum()
            print(
                f"\n[Whitelist {sheet}] "
                f"Incluidos: {n_incluidos} | Excluidos: {n_excluidos}"
            )

        yahoo_df = yahoo_df[keep_mask].copy()

    if yahoo_df.empty:
        print("\n[WARN] No encontré tickers de hojas Yahoo en tickers_cierre.xlsx.")

        out = pd.DataFrame(
            columns=[
                "hoja_origen",
                "ticker",
                "yahoo_symbol",
                "precio_yahoo",
                "status_yahoo",
                "error_yahoo",
            ]
        )

    else:
        yahoo_df = yahoo_df[["hoja_origen", "ticker"]].drop_duplicates()
        yahoo_df = yahoo_df.sort_values(["hoja_origen", "ticker"])

        print(f"\nTickers a consultar en Yahoo Finance: {len(yahoo_df)}")

        rows = []

        for i, row in enumerate(yahoo_df.itertuples(index=False), start=1):
            hoja_origen = row.hoja_origen
            ticker = row.ticker
            yahoo_symbol = yahoo_symbol_from_ticker(ticker, hoja_origen)

            print(
                f"[{i}/{len(yahoo_df)}] "
                f"{hoja_origen} - {ticker} -> {yahoo_symbol}"
            )

            rows.append(
                fetch_yahoo_ticker(
                    ticker=ticker,
                    hoja_origen=hoja_origen,
                )
            )

        out = pd.DataFrame(rows)

    with pd.ExcelWriter(YAHOO_OUTPUT_PATH, engine="openpyxl") as writer:
        out.to_excel(writer, sheet_name="YAHOO_RAW", index=False)

    print("\n========================================")
    print("PROCESO TERMINADO")
    print("========================================")
    print(f"Excel generado: {YAHOO_OUTPUT_PATH}")

    if "status_yahoo" in out.columns:
        print("\nResumen status:")
        print(out["status_yahoo"].value_counts(dropna=False).to_string())

    if "hoja_origen" in out.columns and "status_yahoo" in out.columns and not out.empty:
        print("\nResumen por hoja:")
        print(
            out.groupby(["hoja_origen", "status_yahoo"])
            .size()
            .reset_index(name="cantidad")
            .to_string(index=False)
        )


if __name__ == "__main__":
    main()