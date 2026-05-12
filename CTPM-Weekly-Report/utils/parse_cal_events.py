"""
Parse Shop Cal Events.xlsx.

- Filter to Status == 'Complete'
- Group by month using Calibration Date
- Return 6-month average as projected monthly in-lab calibration count
"""

import pandas as pd
from datetime import date


def parse_cal_events(filepath: str) -> dict:
    engine = "xlrd" if filepath.lower().endswith(".xls") else "openpyxl"
    df = pd.read_excel(filepath, engine=engine)
    df.columns = [str(c).strip() for c in df.columns]

    status_col = _find_col(df, ["Status", "Cal Status", "Calibration Status", "Event Status"])
    date_col = _find_col(df, ["Calibration Date", "Cal Date", "Date", "Completed Date", "Completion Date"])

    if status_col is None:
        raise ValueError("Shop Cal Events: cannot find a Status column.")
    if date_col is None:
        raise ValueError("Shop Cal Events: cannot find a date column.")

    # Filter to completed records
    df = df[df[status_col].astype(str).str.strip().str.lower() == "complete"].copy()

    df["_date"] = pd.to_datetime(df[date_col], errors="coerce")
    df = df.dropna(subset=["_date"])
    df["_month"] = df["_date"].dt.to_period("M")

    # Group by month and count
    monthly_counts = df.groupby("_month").size()

    if monthly_counts.empty:
        return {"monthly_average": 0, "months_used": 0, "monthly_detail": {}}

    # Use the most recent 6 complete months (exclude the current partial month)
    today = pd.Timestamp(date.today())
    current_period = today.to_period("M")

    past_months = monthly_counts[monthly_counts.index < current_period].sort_index()
    recent_six = past_months.tail(6)

    avg = float(recent_six.mean()) if len(recent_six) > 0 else float(monthly_counts.mean())

    return {
        "monthly_average": round(avg),
        "months_used": len(recent_six),
        "monthly_detail": {str(k): int(v) for k, v in recent_six.items()},
    }


def _find_col(df, candidates):
    cols_upper = {c.upper(): c for c in df.columns}
    for c in candidates:
        if c.upper() in cols_upper:
            return cols_upper[c.upper()]
    return None
