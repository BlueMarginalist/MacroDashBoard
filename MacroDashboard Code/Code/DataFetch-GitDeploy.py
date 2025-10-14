import pandas as pd
from fredapi import Fred
from pathlib import Path

# Set up FRED API
api = "958ccd9c67808caf9f941367daf6e812"
fred = Fred(api_key=api)

# Input tickers
# Have to make sure all tickers on the template are included here so later retrieval won't have error
monthly_tickers = [
    "SAHMREALTIME",
    "RECPROUSM156N",
    "PCEDGC96",
    'PCENDC96',
    "PCESC96",
    "RSAFS",
    "DSPIC96",
    "PSAVERT",
    "TOTALSA",
    "REVOLSL",
    "NONREVSL",
    "DGORDER",
    'ADXDNO',
    "INDPRO",
    "TCU",
    "HOUST",
    "PERMIT",
    "HSN1F",
    "EXHOSLUSM495S",
    "PAYEMS",
    "ADPMNUSNERSA",
    "UNRATE",
    "U6RATE",
    "CIVPART",
    "EMRATIO",
    "JTSJOR",
    "JTSHIR",
    "JTSTSR",
    "AWHAETP",
    "CPIAUCSL",
    "CPILFESL",
    "PPIFIS",
    "PCEPI",
    "PCEPILFE",
    "UMCSENT",
    "T5YIFR",
    'CES0500000003',
    "AHETPI",
    "CSUSHPINSA",
    "DTWEXBGS",
    "IQ",
    "IR",
    'BOPTEXP',
    'BOPTIMP',
    'BOPSTB'
]

quarterly_tickers = [
    "GDPC1",
    "NGDPSAXDCUSQ",
    "GDPNOW",
    "PCECC96",
    "PNFIC1",
    "PRFIC1",
    "GCE",
    "ECIWAG",
    'PRS85006092'
]

weekly_tickers = [
    "ICSA",
    "CCSA",
    "MORTGAGE30US"
]

daily_tickers = [
    "T10YIE",
    "DTWEXBGS",
    "FEDFUNDS",
    "DGS2",
    "DGS5",
    "DGS10",
    "DBAA"
]

# Fetch Data and Align Index
def Fetch(ticker_list, fred, freq):
    values_dfs = []
    dates_dfs = []

    # Set Frequency
    f = str(freq).strip().upper()
    if f in ("MONTHLY", "M"):
        period_code = "M"
    elif f in ("QUARTERLY", "Q"):
        period_code = "Q"
    elif f in ("WEEKLY", "W"):
        period_code = "W"
    elif f in ("ANNUAL", "A", "Y"):
        period_code = "A"
    elif f in ("DAILY", "D"):
        period_code = "D"
    else:
        raise ValueError(f"Unsupported frequency: {freq}")

    for ticker in ticker_list:
        try:
            s = fred.get_series(ticker)
        except Exception as e:
            print(f"Warning: couldn't fetch {ticker}: {e}")
            continue

        s.index = pd.to_datetime(s.index)
        df = s.reset_index()
        df.columns = ["release_date", ticker]

        if period_code == "D":
            df["Time"] = df["release_date"].dt.floor("D")
        else:
            df["Time"] = df["release_date"].dt.to_period(period_code).dt.to_timestamp()

        grouped = df.groupby("Time").agg({ticker: "last", "release_date": "last"})
        values_dfs.append(grouped[[ticker]])
        dates_dfs.append(grouped[["release_date"]].rename(columns={"release_date": ticker}))

    if not values_dfs:
        return pd.DataFrame(), pd.DataFrame()

    values = pd.concat(values_dfs, axis=1, sort=True)
    dates = pd.concat(dates_dfs, axis=1, sort=True)
    values.index.name = "Time"
    dates.index.name = "Time"

    return values, dates

monthly_values, monthly_dates = Fetch(monthly_tickers, fred, "M")
quarterly_values, quarterly_dates = Fetch(quarterly_tickers, fred, "Q")
weekly_values, weekly_dates = Fetch(weekly_tickers, fred, "W")
daily_values, daily_dates = Fetch(daily_tickers, fred, "D")

# Save the data table to the folder "Raw Data"
base_dir = Path.cwd().parent.parent
print(base_dir)
raw_location = base_dir / "Raw Data"
raw_location.mkdir(parents=True, exist_ok=True)

monthly_values.to_csv(raw_location / "Monthly_Values.csv")
monthly_dates.to_csv(raw_location / "Monthly_Dates.csv")
quarterly_values.to_csv(raw_location / "Quarterly_Values.csv")
quarterly_dates.to_csv(raw_location / "Quarterly_Dates.csv")
weekly_values.to_csv(raw_location / "Weekly_Values.csv")
weekly_dates.to_csv(raw_location / "Weekly_Dates.csv")
daily_values.to_csv(raw_location / "Daily_Values.csv")
daily_dates.to_csv(raw_location / "Daily_Dates.csv")