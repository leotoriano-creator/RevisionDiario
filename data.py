import yfinance as yf
import pandas as pd


# =========================
# TICKERS + SECTORES (ACTUALIZADO)
# =========================
TICKERS = {
    "TX": "Insumos Básicos",
    "TGS": "Energía",
    "YPF": "Oil & Gas",
    "CEPU": "Energía",
    "PAM": "Energía",

    "BBAR": "Finanzas",
    "BMA": "Finanzas",
    "GGAL": "Finanzas",
    "SUPV": "Finanzas",
    "CRESY": "Finanzas",
    "IRS": "Finanzas",

    "LOMA": "Industrial",
    "EDN": "Utilities",

    "VIST": "Oil & Gas",

    "TEO": "Comunicaciones",
    "MELI": "Comunicaciones",

    "GLOB": "Tecnología",

    "CAAP": "Consumo",
    "BIOX": "Consumo"
}


# =========================
# EVENTOS
# =========================
EVENTS = {
    "PASO19": ("2019-08-09", "2019-08-12"),
    "KIC25": ("2025-09-05", "2025-09-08"),
    "MILEI23": ("2023-11-17", "2023-11-21"),
    "MILEI25": ("2025-10-24", "2025-10-27")
}

UPSIDE_EVENTS = ["MILEI23", "MILEI25"]
DOWNSIDE_EVENTS = ["PASO19", "KIC25"]


# =========================
# DATA
# =========================
def download_data(tickers):
    return yf.download(
        list(tickers),
        start="2017-01-01",
        end="2026-12-31",
        auto_adjust=True,
        progress=False
    )


def get_price(series, date):
    s = series.loc[:date].dropna()
    return s.iloc[-1] if not s.empty else None


# =========================
# CALCULO DE BETAS
# =========================
def compute_betas(data):

    results = []

    # 1️⃣ Calcular el shock de mercado por evento
    event_shock = {}

    for event, (t0, t1) in EVENTS.items():
        returns = []

        for ticker in TICKERS:
            try:
                s = data["Close"][ticker]
                r = get_price(s, t1) / get_price(s, t0) - 1
                returns.append(r)
            except:
                continue

        event_shock[event] = sum(returns) / len(returns)

    # 2️⃣ Calcular betas por activo
    for ticker, sector in TICKERS.items():

        row = {"Ticker": ticker, "Sector": sector}

        betas_up = []
        betas_down = []

        for event, (t0, t1) in EVENTS.items():
            try:
                s = data["Close"][ticker]

                r = get_price(s, t1) / get_price(s, t0) - 1
                shock = event_shock[event]

                beta = r / shock if shock != 0 else None

                row[f"Return_{event}"] = r
                row[f"Beta_{event}"] = beta

                if event in UPSIDE_EVENTS:
                    betas_up.append(beta)
                else:
                    betas_down.append(beta)

            except:
                row[f"Return_{event}"] = None
                row[f"Beta_{event}"] = None

        # 3️⃣ Métricas finales
        row["Beta_Upside"] = sum(betas_up) / len(betas_up)
        row["Beta_Downside"] = sum(betas_down) / len(betas_down)
        row["Convexidad"] = row["Beta_Upside"] - abs(row["Beta_Downside"])

        results.append(row)

    return pd.DataFrame(results)


# =========================
# MAIN
# =========================
def main():

    print("Descargando datos...")
    data = download_data(TICKERS.keys())

    print("Calculando sensibilidad política...")
    df = compute_betas(data)

    print(df)

    df.to_excel("political_betas.xlsx", index=False)

    # Agregado por sector
    sector_df = df.groupby("Sector")[["Beta_Upside", "Beta_Downside", "Convexidad"]].mean()
    sector_df.to_excel("political_betas_sector.xlsx")

    print("Archivos generados:")
    print("- political_betas.xlsx")
    print("- political_betas_sector.xlsx")


if __name__ == "__main__":
    main()