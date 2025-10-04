import pandas as pd
from fredapi import Fred

# Set up FRED API
api="958ccd9c67808caf9f941367daf6e812"
fred = Fred(api_key=api)

# Input Tickers
Monthly_Tickers=[
    "SAHMREALTIME",
    "RECPROUSM156N",
    "PCEDGC96",
    "PCESC96",
    "RSAFS",
    "DSPIC96",
    "PSAVERT",
    "TOTALSA",
    "REVOLSL",
    "NONREVSL",
    "DGORDER",
    "INDPRO",
    "TCU",
    "HOUST",
    "PERMIT",
    "HSN1F",
    "EXHOSLUSM495S",
    "EXPGS",
    "IMPGS",
    "NETEXP",
    "PAYEMS",
    "ADPMNUSNERSA",
    "UNRATE",
    "U6RATE",
    "CIVPART",
    "EMRATIO",
    "JTSJOR",
    "JTSHIR",
    "JTSTSR",
    "CPIAUCSL",
    "CPILFESL",
    "PPIFIS",
    "PCEPI",
    "PCEPILFE",
    "UMCSENT",
    "T5YIFR",
    "AHETPI",
    "CSUSHPINSA",
    "DTWEXBGS",
    "IQ",
    "IR"
]

Quarterly_Tickers=[
    "GDPC1",
    "NGDPSAXDCUSQ",
    "GDPNOW",
    "PCECC96",
    "PCNDGC96",
    "PNFIC1",
    "PRFIC1",
    "GCE",
    "ECIWAG"
]

Weekly_Tickers=[
    "ICSA",
    "CCSA",
    "MORTGAGE30US"
]

Daily_Tickers = [
    "T10YIE",
    "DTWEXBGS",
    "FEDFUNDS",
    "DGS2",
    "DGS5",
    "DGS10",
    "DBAA"
]

#Fetch Data and Align Index
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

    Values = pd.concat(values_dfs, axis=1, sort=True)
    Dates = pd.concat(dates_dfs, axis=1, sort=True)
    Values.index.name = "Time"
    Dates.index.name = "Time"

    return Values, Dates

Monthly_Values,Monthly_Dates = Fetch(Monthly_Tickers,fred,"M")
Quarterly_Values,Quarterly_Dates = Fetch(Quarterly_Tickers,fred,"Q")
Weekly_Values,Weekly_Dates = Fetch(Weekly_Tickers,fred,"W")
Daily_Values,Daily_Dates = Fetch(Daily_Tickers,fred,"D")