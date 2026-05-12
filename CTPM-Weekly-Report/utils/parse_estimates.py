"""
Parse Field Estimates or Shop Estimates .xlsx files.

Structure:
  - Hierarchical rows. Header rows have 'CTPM' in column index 3.
  - Col index 7 = date estimate was converted to a work order (confirmed) or empty (pending).
  - Col index 9 = line item costs that belong to the most-recently-seen header row.
"""

import pandas as pd


def parse_estimates(filepath: str) -> dict:
    engine = "xlrd" if filepath.lower().endswith(".xls") else "openpyxl"
    df = pd.read_excel(filepath, header=None, engine=engine)

    confirmed_total = 0.0
    pending_total = 0.0
    confirmed_count = 0
    pending_count = 0

    # Track whether the current "open" header is confirmed or pending.
    current_confirmed: bool | None = None

    for _, row in df.iterrows():
        col3 = row.iloc[3] if len(row) > 3 else None
        col7 = row.iloc[7] if len(row) > 7 else None
        col9 = row.iloc[9] if len(row) > 9 else None

        # Detect header row: col 3 contains the string 'CTPM'
        if pd.notna(col3) and "CTPM" in str(col3).upper():
            has_wo_date = _has_date(col7)
            current_confirmed = has_wo_date
            if has_wo_date:
                confirmed_count += 1
            else:
                pending_count += 1
            # If the header row itself also has a cost in col 9, count it
            cost = _to_float(col9)
            if cost > 0 and current_confirmed is not None:
                if current_confirmed:
                    confirmed_total += cost
                else:
                    pending_total += cost
            continue

        # Line item row — accumulate cost under the current header
        if current_confirmed is not None:
            cost = _to_float(col9)
            if cost > 0:
                if current_confirmed:
                    confirmed_total += cost
                else:
                    pending_total += cost

    return {
        "confirmed_total": confirmed_total,
        "pending_total": pending_total,
        "confirmed_count": confirmed_count,
        "pending_count": pending_count,
    }


def _has_date(value) -> bool:
    if value is None or not pd.notna(value):
        return False
    s = str(value).strip()
    return s not in ("", "nan", "NaT", "None", "NaN")


def _to_float(value) -> float:
    if value is None or not pd.notna(value):
        return 0.0
    try:
        v = float(str(value).replace(",", "").replace("$", "").strip())
        return v if v > 0 else 0.0
    except (ValueError, TypeError):
        return 0.0
