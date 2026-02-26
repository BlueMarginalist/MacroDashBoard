"""
generate_html.py  —  Static HTML generator for the Macro Dashboard.

Ports all logic from streamlit_app.py and writes Website/index.html.
Run via: python generate_html.py
"""

import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime, timezone

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
RAW_DIR = BASE_DIR / "Raw Data"
DOCS_DIR = BASE_DIR / "Website"

# ---------------------------------------------------------------------------
# Data helpers  (identical to streamlit_app.py, minus @st.cache_data)
# ---------------------------------------------------------------------------

def NormalizeFreq(code: str) -> str:
    fm = {
        ("MONTHLY", "M"): "M",
        ("QUARTERLY", "Q"): "Q",
        ("WEEKLY", "W"): "W",
        ("ANNUAL", "A", "Y"): "A",
        ("DAILY", "D"): "D",
    }
    code_up = code.upper()
    for keys, val in fm.items():
        if code_up in keys:
            return val
    raise ValueError(f"Unsupported frequency: {code}")


def _load_csv(path: str):
    return pd.read_csv(path, index_col=0, parse_dates=True)


def GetData(ticker: str, freq: str):
    period_code = NormalizeFreq(freq)
    dates = _load_csv(str(RAW_DIR / f"{period_code}_Dates.csv"))
    values = _load_csv(str(RAW_DIR / f"{period_code}_Values.csv"))
    if ticker not in values.columns:
        raise KeyError(f"Ticker '{ticker}' not found.")
    last_idx = values[ticker].last_valid_index()
    last_pos = values.index.get_loc(last_idx)
    return dates[ticker].iloc[last_pos], values[ticker], last_idx, dates[ticker]


def GetLevel(ticker: str, freq: str, n_lags: int = 4):
    latest_date, value_col, current_period, date_col = GetData(ticker, freq)
    last_idx = value_col.last_valid_index()
    last_pos = value_col.index.get_loc(last_idx)
    period_code = NormalizeFreq(freq)
    latest_dates = date_col.tail(n_lags + 1).tolist()
    if period_code == "D":
        non_null = value_col.iloc[: last_pos + 1].dropna()
        return latest_date, non_null.tail(n_lags + 1).tolist(), current_period, latest_dates
    start_pos = max(0, last_pos - n_lags)
    return latest_date, value_col.iloc[start_pos: last_pos + 1].tolist(), current_period, latest_dates


def GetDelta(ticker: str, freq: str, agg_freq: str, n_lags: int = 4, pct: bool = True):
    freq_code = NormalizeFreq(freq)
    agg_code = NormalizeFreq(agg_freq)
    latest_date, value_col, current_period, date_col = GetData(ticker, freq_code)
    last_idx = value_col.last_valid_index()
    last_pos = value_col.index.get_loc(last_idx)
    start_pos = max(0, last_pos - n_lags)
    latest_dates = date_col.tail(n_lags + 1).tolist()

    if agg_code == freq_code:
        if pct:
            result = value_col.pct_change(1, fill_method=None)
        else:
            result = value_col.diff(1)
        return latest_date, result.iloc[start_pos: last_pos + 1].tolist(), current_period, latest_dates

    idx = value_col.index
    if isinstance(idx, pd.DatetimeIndex) or pd.api.types.is_datetime64_any_dtype(idx):
        offset_map = {
            ("M", "A"): pd.DateOffset(months=12),
            ("M", "Q"): pd.DateOffset(months=3),
            ("Q", "A"): pd.DateOffset(months=12),
            ("W", "A"): pd.DateOffset(weeks=52),
            ("D", "A"): pd.DateOffset(years=1),
        }
        offset = offset_map.get((freq_code, agg_code))
        if offset is None:
            raise ValueError(f"Unsupported conversion: {freq_code} -> {agg_code}")
        prev = value_col.shift(freq=offset).reindex(value_col.index)
        if pct:
            result = (value_col - prev) / prev
        else:
            result = value_col - prev
        return latest_date, result.iloc[start_pos: last_pos + 1].tolist(), current_period, latest_dates

    int_shift = {("M", "A"): 12, ("M", "Q"): 3, ("Q", "A"): 4}.get((freq_code, agg_code))
    if int_shift is None:
        raise ValueError(f"Unsupported conversion: {freq_code} -> {agg_code}")
    if pct:
        result = value_col.pct_change(int_shift, fill_method=None)
    else:
        result = value_col.diff(int_shift)
    return latest_date, result.iloc[start_pos: last_pos + 1].tolist(), current_period, latest_dates


def isNewRelease(ticker: str, freq: str) -> bool:
    freq_code = NormalizeFreq(freq)
    if freq_code == "D":
        return True
    raw_path = RAW_DIR / f"{freq_code}_Raws.csv"
    if not raw_path.exists():
        return True
    raw = pd.read_csv(raw_path, low_memory=False)
    _, _, period, _ = GetLevel(ticker, freq, n_lags=1)
    period = pd.to_datetime(period)
    col = f"Time_{ticker}"
    if col not in raw.columns:
        return True
    raw[col] = pd.to_datetime(raw[col])
    return raw[raw[col] == period].shape[0] <= 1


# ---------------------------------------------------------------------------
# Indicator config  (identical to streamlit_app.py)
# ---------------------------------------------------------------------------

OUTPUT_INDICATORS = [
    ("RGDP", "GDPC1", "Q", "Trln $"),
    ("RGDP", "GDPC1", "Q", "Q/Q % Delta SAAR"),
    ("NGDP", "NGDPSAXDCUSQ", "Q", "Trln $"),
    ("NGDP", "NGDPSAXDCUSQ", "Q", "Q/Q % Delta SAAR"),
    ("GDP Nowcast", "GDPNOW", "Q", "Q/Q % Delta SAAR"),
    ("Smoothed Recession Prob", "RECPROUSM156N", "M", "% (5=5%)"),
    ("RPCE", "PCECC96", "Q", "Trln $"),
    ("RPCE", "PCECC96", "Q", "Q/Q % Delta SAAR"),
    ("RPCE Dur", "PCEDGC96", "M", "M/M % Delta"),
    ("RPCE Dur", "PCEDGC96", "M", "Y/Y % Delta"),
    ("RPCE Ndur", "PCENDC96", "M", "M/M % Delta"),
    ("RPCE Ndur", "PCENDC96", "M", "Y/Y % Delta"),
    ("RPCE Serv.", "PCESC96", "M", "M/M % Delta"),
    ("RPCE Serv.", "PCESC96", "M", "Y/Y % Delta"),
    ("Retail Sales", "RSAFS", "M", "M/M % Delta"),
    ("Retail Sales", "RSAFS", "M", "Y/Y % Delta"),
    ("Real Disp. Personal Inc.", "DSPIC96", "M", "M/M % Delta"),
    ("Real Disp. Personal Inc.", "DSPIC96", "M", "Y/Y % Delta"),
    ("Personal Saving Rate", "PSAVERT", "M", "%, SAAR"),
    ("Vehicle Sales", "TOTALSA", "M", "Mln Units"),
    ("Vehicle Sales", "TOTALSA", "M", "Y/Y % Delta"),
    ("Cons Credit - Revolving", "REVOLSL", "M", "M/M % Delta SAAR"),
    ("Cons Credit - NonRev", "NONREVSL", "M", "M/M % Delta SAAR"),
    ("Gross Priv Fixed Inv NR", "PNFIC1", "Q", "Q/Q % Delta SAAR"),
    ("Gross Priv Fixed Inv Res", "PRFIC1", "Q", "Q/Q % Delta SAAR"),
    ("Dur. Order", "DGORDER", "M", "M/M % Delta"),
    ("Dur. Order", "DGORDER", "M", "Y/Y % Delta"),
    ("Dur Orders Non Def x Air", "ADXDNO", "M", "M/M % Delta"),
    ("Dur Orders Non Def x Air", "ADXDNO", "M", "Y/Y % Delta"),
    ("IP", "INDPRO", "M", "M/M % Delta"),
    ("IP", "INDPRO", "M", "Y/Y % Delta"),
    ("Cap Util", "TCU", "M", "%"),
    ("Productivity", "PRS85006092", "Q", "Q/Q % Delta SAAR"),
    ("Hou. Starts", "HOUST", "M", "SAAR (Thousands)"),
    ("Hou. Starts", "HOUST", "M", "Y/Y % Delta"),
    ("Housing Permits", "PERMIT", "M", "SAAR (Thousands)"),
    ("Housing Permits", "PERMIT", "M", "Y/Y % Delta"),
    ("New Home Sales", "HSN1F", "M", "SAAR (Thousands)"),
    ("New Home Sales", "HSN1F", "M", "Y/Y % Delta"),
    ("Existing Home Sales", "EXHOSLUSM495S", "M", "SAAR (Thousands)"),
    ("Existing Home Sales", "EXHOSLUSM495S", "M", "Y/Y % Delta"),
    ("Gov. Cons", "GCE", "Q", "Trln $"),
    ("Gov. Cons", "GCE", "Q", "Q/Q % Delta SAAR"),
    ("Exports", "BOPTEXP", "M", "Bln $"),
    ("Exports", "BOPTEXP", "M", "M/M % Delta SA"),
    ("Imports", "BOPTIMP", "M", "Bln $"),
    ("Imports", "BOPTIMP", "M", "M/M % Delta SA"),
    ("Trade Balance", "BOPSTB", "M", "Bln $"),
    ("Trade Balance", "BOPSTB", "M", "M/M % Delta SA"),
]

LABOR_INDICATORS = [
    ("NFP, Total NonFarm", "PAYEMS", "M", "M/M Delta (Thousands)"),
    ("NFP, Total NonFarm", "PAYEMS", "M", "Y/Y % Delta"),
    ("ADP, Total NonFarm Priv", "ADPMNUSNERSA", "M", "M/M Delta (Thousands)"),
    ("UR", "UNRATE", "M", "%"),
    ("U-6", "U6RATE", "M", "%"),
    ("LFPR", "CIVPART", "M", "%"),
    ("EPop", "EMRATIO", "M", "%"),
    ("JOLTS Openings Rate", "JTSJOR", "M", "%"),
    ("JOLTS Hires Rate", "JTSHIR", "M", "%"),
    ("JOLTS Separations Rate", "JTSTSR", "M", "%"),
    ("UI Initial Claims", "ICSA", "W", "# People"),
    ("UI Continuing Claims", "CCSA", "W", "# People"),
    ("Avg Weekly Hours", "AWHAETP", "M", "Hrs"),
]

PRICES_INDICATORS = [
    ("CPI", "CPIAUCSL", "M", "M/M % Delta"),
    ("CPI", "CPIAUCSL", "M", "Y/Y % Delta"),
    ("Core CPI", "CPILFESL", "M", "M/M % Delta"),
    ("Core CPI", "CPILFESL", "M", "Y/Y % Delta"),
    ("PPI-FD", "PPIFIS", "M", "M/M % Delta"),
    ("PPI-FD", "PPIFIS", "M", "Y/Y % Delta"),
    ("PCE Deflator", "PCEPI", "M", "M/M % Delta"),
    ("PCE Deflator", "PCEPI", "M", "Y/Y % Delta"),
    ("Core PCE Deflator", "PCEPILFE", "M", "M/M % Delta"),
    ("Core PCE Deflator", "PCEPILFE", "M", "Y/Y % Delta"),
    ("Mich NTM Inflation Exp", "MICH", "M", "%"),
    ("5yr, 5yr Forward", "T5YIFR", "D", "%, NSA"),
    ("10yr TIPS", "T10YIE", "D", "%, NSA"),
    ("ECI", "ECIWAG", "Q", "Q/Q % Delta"),
    ("ECI", "ECIWAG", "Q", "Y/Y % Delta"),
    ("Avg Hrly Earnings", "CES0500000003", "M", "M/M % Delta"),
    ("Avg Hrly Earnings", "CES0500000003", "M", "Y/Y % Delta"),
    ("Real Avg. Hourly Earnings", "RCES0500000003*", "M", "M/M % Delta"),
    ("Real Avg. Hourly Earnings", "RCES0500000003*", "M", "Y/Y % Delta"),
    ("Case Shiller HPI", "CSUSHPINSA", "M", "M/M % Delta"),
    ("Case Shiller HPI", "CSUSHPINSA", "M", "Y/Y % Delta"),
    ("Nominal Broad USD Index", "TWEXBGSMTH", "M", "Index Jan2006=100"),
    ("Nominal Broad USD Index", "TWEXBGSMTH", "M", "Y/Y % Delta"),
    ("Export Prices", "IQ", "M", "M/M % Delta"),
    ("Export Prices", "IQ", "M", "Y/Y % Delta"),
    ("Import Prices", "IR", "M", "M/M % Delta"),
    ("Import Prices", "IR", "M", "Y/Y % Delta"),
]

RATES_INDICATORS = [
    ("FFR", "DFF", "D", "%, NSA"),
    ("2y UST", "DGS2", "D", "%, NSA"),
    ("5y UST", "DGS5", "D", "%, NSA"),
    ("10y UST", "DGS10", "D", "%, NSA"),
    ("30y Mtg.", "MORTGAGE30US", "W", "%, NSA"),
    ("BAA", "DBAA", "D", "%, NSA"),
]

# ---------------------------------------------------------------------------
# Formatting helpers  (identical to streamlit_app.py)
# ---------------------------------------------------------------------------

def _parse_units(units: str, ticker: str):
    u = units.lower()
    if ticker in ("PRS85006092", "GDPNOW"):
        return False, None, False
    if "delta" not in u:
        return False, None, False
    is_pct = "%" in u
    agg_code = None
    if "/" in u:
        parts = u.split("/")
        agg_letter = parts[1].strip()[0].upper()
        agg_map = {"M": "M", "Q": "Q", "Y": "A", "A": "A"}
        agg_code = agg_map.get(agg_letter, "M")
    return True, agg_code, is_pct


def _fmt_value(val, units: str):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return ""
    u = units.lower()
    is_delta = "delta" in u
    is_pct = "%" in u

    if is_delta and is_pct:
        return f"{val * 100:.2f}%"
    if is_delta and not is_pct:
        if abs(val) >= 1:
            return f"{val:,.1f}"
        return f"{val:.4f}"
    if isinstance(val, float):
        if abs(val) >= 1000:
            return f"{val:,.1f}"
        if abs(val) >= 10:
            return f"{val:.2f}"
        return f"{val:.4f}"
    return str(val)


def _fmt_period(period):
    if period is None:
        return ""
    try:
        ts = pd.to_datetime(period)
        return ts.strftime("%Y-%m")
    except Exception:
        return str(period)


def build_section(indicators, n_lags=4):
    rows = []
    for name, ticker, freq, units in indicators:
        is_delta, agg_code, is_pct = _parse_units(units, ticker)
        try:
            if is_delta and agg_code:
                latest_date, values, period, _ = GetDelta(ticker, freq, agg_code, n_lags=n_lags, pct=is_pct)
            else:
                latest_date, values, period, _ = GetLevel(ticker, freq, n_lags=n_lags)
        except Exception as e:
            rows.append({
                "Indicator": name,
                "Units": units,
                "Period": "",
                "Present": f"ERR: {e}",
                **{f"Lag {i}": "" for i in range(1, n_lags + 1)},
                "_fresh": False,
            })
            continue

        expected = n_lags + 1
        vals = list(values)
        if len(vals) < expected:
            vals = [None] * (expected - len(vals)) + vals

        present = vals[-1]
        lags = [vals[-1 - i] if (-1 - i) >= -len(vals) else None for i in range(1, n_lags + 1)]

        fresh = False
        try:
            if pd.notnull(latest_date):
                latest_ts = pd.to_datetime(latest_date)
                if pd.Timestamp.now() - latest_ts <= pd.Timedelta(days=7):
                    fresh = isNewRelease(ticker, freq)
        except Exception:
            pass

        row = {
            "Indicator": name,
            "Units": units,
            "Period": _fmt_period(period),
            "Present": _fmt_value(present, units),
            "_fresh": fresh,
        }
        for i, lag in enumerate(lags, 1):
            row[f"Lag {i}"] = _fmt_value(lag, units)
        rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# HTML rendering
# ---------------------------------------------------------------------------

def render_section_html(title: str, indicators) -> str:
    """Return an HTML string for a titled section table."""
    rows = build_section(indicators)
    if not rows:
        return ""

    display_cols = ["Indicator", "Units", "Period", "Present", "Lag 1", "Lag 2", "Lag 3", "Lag 4"]

    html = f'<h2 class="section-title">{title}</h2>\n'
    html += '<table>\n<thead><tr>'
    for col in display_cols:
        html += f'<th>{col}</th>'
    html += '</tr></thead>\n<tbody>\n'

    for i, row in enumerate(rows):
        row_class = "even" if i % 2 == 0 else "odd"
        fresh = row.get("_fresh", False)
        html += f'<tr class="{row_class}">'
        for col in display_cols:
            val = row.get(col, "")
            if col == "Period" and fresh:
                html += f'<td class="fresh">{val}</td>'
            else:
                html += f'<td>{val}</td>'
        html += '</tr>\n'

    html += '</tbody>\n</table>\n'
    return html


CSS = """
* { box-sizing: border-box; margin: 0; padding: 0; }

body {
    font-family: 'Courier New', Courier, monospace;
    font-size: 13px;
    background: #f4f4f4;
    color: #222;
    padding: 16px;
}

h1 {
    font-size: 22px;
    font-weight: bold;
    color: #1a1a2e;
    margin-bottom: 4px;
}

.subtitle {
    font-size: 11px;
    color: #555;
    margin-bottom: 12px;
}

hr {
    border: none;
    border-top: 1px solid #ccc;
    margin: 10px 0 16px;
}

.layout {
    display: flex;
    gap: 24px;
    align-items: flex-start;
}

.col-left  { flex: 10; min-width: 0; }
.col-right { flex: 10; min-width: 0; }

.section-title {
    font-size: 15px;
    font-weight: bold;
    color: #1a1a2e;
    margin: 20px 0 6px;
}

.col-left .section-title:first-child,
.col-right .section-title:first-child {
    margin-top: 0;
}

table {
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 20px;
}

th {
    background: #1a1a2e;
    color: #eee;
    text-align: left;
    padding: 4px 8px;
    border: 1px solid #555;
    white-space: nowrap;
}

td {
    padding: 4px 8px;
    border: 1px solid #ddd;
    white-space: nowrap;
}

tr.even td { background: #f9f9f9; }
tr.odd  td { background: #ffffff; }

td.fresh { background: #FFFF99 !important; }

.footer {
    margin-top: 16px;
    font-size: 11px;
    color: #555;
    border-top: 1px solid #ccc;
    padding-top: 8px;
}
"""


def generate_html() -> str:
    now_utc = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")

    left_html = render_section_html("Output", OUTPUT_INDICATORS)

    right_html = (
        render_section_html("Labor", LABOR_INDICATORS)
        + render_section_html("Prices", PRICES_INDICATORS)
        + render_section_html("Rates", RATES_INDICATORS)
    )

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Macro Dashboard</title>
<style>
{CSS}
</style>
</head>
<body>

<h1>Macro Dashboard</h1>
<p class="subtitle">Source: FRED / Federal Reserve Bank of St. Louis &nbsp;|&nbsp; Prof. Mike Aguilar</p>
<hr>

<div class="layout">
  <div class="col-left">
{left_html}
  </div>
  <div class="col-right">
{right_html}
  </div>
</div>

<div class="footer">
  <p>All data sourced from FRED. Most series are seasonally adjusted unless noted (NSA).
  RCES0500000003* = CES0500000003 / (PCEPI/100).
  Yellow highlight = new release (within 7 days).</p>
  <p style="margin-top:4px;">Last updated: {now_utc}</p>
</div>

</body>
</html>
"""


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    DOCS_DIR.mkdir(exist_ok=True)
    out_path = DOCS_DIR / "index.html"
    html = generate_html()
    out_path.write_text(html, encoding="utf-8")
    print(f"Generated {out_path}")
