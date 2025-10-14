import pandas as pd
import os
from openpyxl import load_workbook
from typing import List, Union
from pathlib import Path

# Set Location
base_dir = Path(__file__).resolve().parent.parent.parent
print(base_dir)
raw_location = base_dir / "Raw Data"
raw_location.mkdir(parents=True, exist_ok=True)

# GetData from the csv files
def GetData(Ticker: str, Freq: str, n_lags: int = 4):
    if Freq in ("MONTHLY", "M"):
        period_code = "Monthly"
    elif Freq in ("QUARTERLY", "Q"):
        period_code = "Quarterly"
    elif Freq in ("WEEKLY", "W"):
        period_code = "Weekly"
    elif Freq in ("ANNUAL", "A", "Y"):
        period_code = "Annual"
    elif Freq in ("DAILY", "D"):
        period_code = "Daily"
    else:
        raise ValueError(f"Unsupported frequency: {Freq}")

    base = base_dir / "Raw Data"
    dates_path = os.path.join(base, f"{period_code}_Dates.csv")
    values_path = os.path.join(base, f"{period_code}_Values.csv")

    dates = pd.read_csv(dates_path)
    values = pd.read_csv(values_path)

    if Ticker not in values.columns:
        raise KeyError(f"Ticker '{Ticker}' not found in values file.")

    last_idx = values[Ticker].last_valid_index()
    last_pos = values.index.get_loc(last_idx)

    start_pos = max(0, last_pos - n_lags)
    latest_values = values[Ticker].iloc[start_pos: last_pos + 1].tolist()
    latest_date = dates[Ticker].iloc[last_pos]
    return latest_date, latest_values

# Get the template and location for the updated version
folder = base_dir / "MacroDashboard Versions"
input_xlsx = folder / "Economic Dashboard V1-Template.xlsx"
output_xlsx = folder / "Economic Dashboard V1-Template-UPDATE.xlsx"

# Current csv file is in Timestamp format, so change it to datetime
def _to_python_datetime(dt: Union[pd.Timestamp, str, None]):
    """
    Convert pandas.Timestamp -> python datetime for openpyxl.
    If dt is None or unparseable, return it unchanged.
    """
    if isinstance(dt, pd.Timestamp):
        return dt.to_pydatetime()
    return dt

# Load excel template and select the range of rows to update data
wb = load_workbook(input_xlsx)
ws = wb.active
start_row = 3
max_row = ws.max_row

# write data into the rows, one row each time
def write_panel_for_row(
    row: int,
    ticker_col: str,
    freq_col: str,
    date_col: str,
    present_col: str,
    lag_cols: List[str],
    n_lags: int = 4,
) -> None:
    """
    Write one row for a panel (left or right).
    - ticker_col, freq_col, date_col, present_col: column letters (e.g. "B", "E", "C", "F")
    - lag_cols: list of lag column letters in order [Lag1_col, Lag2_col, ...]
    - n_lags: number of lags expected (defaults to 4)
    """
    # read ticker and freq values from the worksheet
    raw_ticker = ws[f"{ticker_col}{row}"].value
    raw_freq = ws[f"{freq_col}{row}"].value

    # If ticker cell is empty -> skip row
    if raw_ticker is None or str(raw_ticker).strip() == "" or str(raw_ticker).strip() == "Freq":
        return

    # normalize strings safely
    ticker = str(raw_ticker).strip()
    freq = str(raw_freq).strip() if raw_freq is not None else "M"

    # get data using your GetData function (assumed defined/imported)
    try:
        latest_date, recent_values = GetData(ticker, freq, n_lags=n_lags)
    except Exception as exc:
        # write error to date cell and skip writing values for this row
        ws[f"{date_col}{row}"].value = f"ERR: {exc}"
        return

    # recent_values expected chronological oldest ... most recent
    values = list(recent_values)

    # Ensure fixed length = n_lags + 1 by padding at front (older side) with None
    expected_len = n_lags + 1
    if len(values) < expected_len:
        pad_len = expected_len - len(values)
        values = [None] * pad_len + values

    # Present is last element, lag1 is second-last, etc.
    present_value = values[-1]
    lag_values = []
    for i in range(1, len(lag_cols) + 1):
        # i=1 -> lag1 => values[-2], general: values[-1 - i]
        idx = -1 - i
        # safe access: if index out of range, return None
        try:
            lag_values.append(values[idx])
        except IndexError:
            lag_values.append(None)

    # write date (convert pandas Timestamp -> python datetime for openpyxl)
    ws[f"{date_col}{row}"].value = _to_python_datetime(latest_date)

    # write present and lag values
    ws[f"{present_col}{row}"].value = present_value
    for col_letter, lag_value in zip(lag_cols, lag_values):
        ws[f"{col_letter}{row}"].value = lag_value


# Output section is the left panel, and other three sections are in the right panel.
# Can automate the identification of columns for each panel
# Left panel: B = ticker, E = freq, C = Latest Date, F = Present, G,H,I,J = Lag1..Lag4
left_panel = {
    "ticker_col": "B",
    "freq_col": "E",
    "date_col": "C",
    "present_col": "F",
    "lag_cols": ["G", "H", "I", "J"],
}

# Right panel: M = ticker, P = freq, N = Latest Date, Q = Present, R,S,T,U = Lag1..Lag4
right_panel = {
    "ticker_col": "M",
    "freq_col": "P",
    "date_col": "N",
    "present_col": "Q",
    "lag_cols": ["R", "S", "T", "U"],
}

# iterate and write
for r in range(start_row, max_row + 1):
    write_panel_for_row(r, **left_panel)
    write_panel_for_row(r, **right_panel)

wb.save(output_xlsx)
