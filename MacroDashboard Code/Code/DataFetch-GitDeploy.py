import pandas as pd
from fredapi import Fred
from pathlib import Path

# Set up FRED API
api = "958ccd9c67808caf9f941367daf6e812"
fred = Fred(api_key=api)

# Input tickers
# Have to make sure all tickers on the template are included here so later retrieval won't have error
monthly_tickers = [
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
    'CES0500000003',
    "AHETPI",
    "CSUSHPINSA",
    "IQ",
    "IR",
    'BOPTEXP',
    'BOPTIMP',
    'BOPSTB',
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
    "T5YIFR",
    "T10YIE",
    "DTWEXBGS",
    "DFF",
    "DGS2",
    "DGS5",
    "DGS10",
    "DBAA"
]

# Fetch Data and Align Index
def Fetch(ticker_list, fred, freq,warning_list=[]):
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

    # Get Data through Fred. throw an error if unable to fetch
    for ticker in [t for t in ticker_list if t not in warning_list]:
        try:
            s = fred.get_series_all_releases(ticker)
        except Exception as e:
            print(f"Warning: couldn't fetch {ticker}: {e}")
            continue
        s["date"] = pd.to_datetime(s["date"])
        s["realtime_start"] = pd.to_datetime(s["realtime_start"])

        # Normalize timestamps by converting each release_date to the start of its time period
        # (e.g., start of day, month, quarter, etc.) for consistent grouping and comparison
        if period_code == "D":
            s["Time"] = s["date"].dt.floor("D")
        else:
            s["Time"] = s["date"].dt.to_period(period_code).dt.to_timestamp()

        # Group by the period and take the last (i.e., the latest release for that period)
        # ticker is the value column and real_time_start is the release date column
        grouped = s.groupby("Time").agg({"value": "last", "realtime_start": "last"})
        values_dfs.append(grouped[["value"]].rename(columns={"value": ticker}))
        dates_dfs.append(grouped[["realtime_start"]].rename(columns={"realtime_start": ticker}))

    # deal with the variables in
    for ticker in warning_list:
        try:
            s = fred.get_series(ticker)
        except Exception as e:
            print(f"Warning: couldn't fetch {ticker}: {e}")
            continue

        # Change to datetime
        s.index = pd.to_datetime(s.index)
        df = s.reset_index()
        df.columns = ["release_date", ticker]

        # Normalize timestamps by converting each release_date to the start of its time period
        # (e.g., start of day, month, quarter, etc.) for consistent grouping and comparison
        # here because release data is temporarily not available because the series is not available in fred. I just keep the release date blank
        if period_code == "D":
            df["Time"] = df["release_date"].dt.floor("D")
        else:
            df["Time"] = df["release_date"].dt.to_period(period_code).dt.to_timestamp()

        grouped = df.groupby("Time").agg({ticker: "last", "release_date": "last"})
        values_dfs.append(grouped[[ticker]])
        dates_dfs.append(grouped[["release_date"]].rename(columns={"release_date": ticker}))

    if not values_dfs:
        return pd.DataFrame(), pd.DataFrame()

    # concat the series of each economic data into a table
    values = pd.concat(values_dfs, axis=1, sort=True)
    dates = pd.concat(dates_dfs, axis=1, sort=True)
    values.index.name = "Time"
    dates.index.name = "Time"

    return values, dates

# Existing Home Sale is not present in ALFRED database, and daily measures have too many vintage dates to be fetched through api
monthly_values, monthly_dates = Fetch(monthly_tickers, fred, "M",['EXHOSLUSM495S'])
quarterly_values, quarterly_dates = Fetch(quarterly_tickers, fred, "Q")
weekly_values, weekly_dates = Fetch(weekly_tickers, fred, "W")
daily_values, daily_dates = Fetch(daily_tickers, fred, "D",daily_tickers)

def AddRealAvgEarning():
    monthly_values['RCES0500000003*']=monthly_values['CES0500000003']/(monthly_values['PCEPI']/100)
    monthly_dates['RCES0500000003*']=monthly_dates['CES0500000003']

AddRealAvgEarning()
# Save the data table to the folder "Raw Data"
base_dir = Path(__file__).resolve().parent.parent.parent
print(base_dir)
raw_location = base_dir / "Raw Data"
raw_location.mkdir(parents=True, exist_ok=True)

monthly_values.to_csv(raw_location / "M_Values.csv")
monthly_dates.to_csv(raw_location / "M_Dates.csv")
quarterly_values.to_csv(raw_location / "Q_Values.csv")
quarterly_dates.to_csv(raw_location / "Q_Dates.csv")
weekly_values.to_csv(raw_location / "W_Values.csv")
weekly_dates.to_csv(raw_location / "W_Dates.csv")
daily_values.to_csv(raw_location / "D_Values.csv")
daily_dates.to_csv(raw_location / "D_Dates.csv")