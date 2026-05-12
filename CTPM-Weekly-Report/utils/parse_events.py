"""
Parse Events.xlsx for shop TAT calculation.

Shop stages (in order):
  1. Receiving In-Shop
  2. Shop Calibration
  3. Qc
  4. Shipping

- Group by Work Order + I.D., pivot to get date of each stage
- Only use instruments with all 4 stages present
- Last 90 days only (based on Shipping date)
- Calculate mean (not median) in business days via numpy.busday_count
"""

import pandas as pd
import numpy as np
from datetime import date, timedelta


STAGE_NAMES = ["Receiving In-Shop", "Shop Calibration", "Qc", "Shipping"]

# Aliases to try when exact names aren't present
STAGE_ALIASES = {
    "Receiving In-Shop": ["receiving in-shop", "receive in shop", "receiving", "in-shop receive", "shop receive"],
    "Shop Calibration": [
        "shop calibration", "shop cal", "calibration", "cal",
        "pipette cal in-shop", "pippette cal in-shop",
        "cover letter cal cert",
    ],
    "Qc": ["qc", "quality control", "quality check", "q.c."],
    "Shipping": ["shipping", "ship", "shipped", "shipment"],
}


def parse_events_tat(filepath: str) -> dict:
    engine = "xlrd" if filepath.lower().endswith(".xls") else "openpyxl"
    df = pd.read_excel(filepath, engine=engine)
    df.columns = [str(c).strip() for c in df.columns]

    wo_col = _find_col(df, ["Work Order", "WO", "WO Number", "WorkOrder", "Work Order Number"])
    id_col = _find_col(df, ["I.D.", "ID", "Item ID", "Instrument ID", "ItemID", "Cal ID"])
    stage_col = _find_col(df, ["Event", "Stage", "Status", "Event Type", "Activity", "Event Name", "Action"])
    date_col = _find_col(df, ["Event Date (Universal)", "Date", "Event Date", "Completed Date", "Completion Date", "Activity Date"])

    if not all([wo_col, id_col, stage_col, date_col]):
        missing = [k for k, v in {"Work Order": wo_col, "ID": id_col, "Stage": stage_col, "Date": date_col}.items() if not v]
        raise ValueError(f"Events file: cannot find columns: {', '.join(missing)}")

    df["_date"] = pd.to_datetime(df[date_col], errors="coerce")
    df = df.dropna(subset=["_date"])
    df["_stage_raw"] = df[stage_col].astype(str).str.strip()
    df["_stage"] = df["_stage_raw"].apply(_normalize_stage)

    # Keep only recognized shop stages
    df = df[df["_stage"].notna()].copy()

    # Cutoff: last 90 days
    cutoff = pd.Timestamp(date.today() - timedelta(days=90))

    # For each WO+ID, pivot to get the date of each stage
    df["_key"] = df[wo_col].astype(str) + "||" + df[id_col].astype(str)

    records = []
    for key, grp in df.groupby("_key"):
        stage_dates = {}
        for _, row in grp.iterrows():
            s = row["_stage"]
            if s not in stage_dates or row["_date"] < stage_dates[s]:
                stage_dates[s] = row["_date"]

        # Only use instruments that have all 4 stages
        if not all(s in stage_dates for s in STAGE_NAMES):
            continue

        ship_date = stage_dates["Shipping"]
        if ship_date < cutoff:
            continue

        records.append(
            {
                "receive": stage_dates["Receiving In-Shop"],
                "cal": stage_dates["Shop Calibration"],
                "qc": stage_dates["Qc"],
                "ship": stage_dates["Shipping"],
            }
        )

    if not records:
        return {"receive_to_cal": 0.0, "cal_to_qc": 0.0, "qc_to_ship": 0.0, "total": 0.0, "sample_size": 0}

    def _bdays(start_series, end_series):
        values = []
        for s, e in zip(start_series, end_series):
            try:
                bd = int(np.busday_count(s.date(), e.date()))
                values.append(max(bd, 0))
            except Exception:
                pass
        return float(np.mean(values)) if values else 0.0

    df_r = pd.DataFrame(records)
    r2c = _bdays(df_r["receive"], df_r["cal"])
    c2q = _bdays(df_r["cal"], df_r["qc"])
    q2s = _bdays(df_r["qc"], df_r["ship"])

    return {
        "receive_to_cal": round(r2c, 1),
        "cal_to_qc": round(c2q, 1),
        "qc_to_ship": round(q2s, 1),
        "total": round(r2c + c2q + q2s, 1),
        "sample_size": len(records),
    }


def _normalize_stage(raw: str) -> str | None:
    raw_lower = raw.lower().strip()
    for canonical, aliases in STAGE_ALIASES.items():
        if raw_lower == canonical.lower() or raw_lower in aliases:
            return canonical
    # Substring fallback: anything containing "cal" that isn't shipping counts as Shop Calibration
    if "cal" in raw_lower and "ship" not in raw_lower and "receiv" not in raw_lower:
        return "Shop Calibration"
    return None


def _find_col(df, candidates):
    cols_upper = {c.upper(): c for c in df.columns}
    for c in candidates:
        if c.upper() in cols_upper:
            return cols_upper[c.upper()]
    return None
