"""
Parse Current WIP .xls file and estimate costs via tiered lookup against Historic Calibrations.

Key columns: I.D., Company, Type, Description, Serial Number, Work Order, Received On
Cost is not in the file — derived via HistoricLookup tiered match.
"""

import pandas as pd
from utils.parse_historic import HistoricLookup


def parse_wip(filepath: str, lookup: HistoricLookup) -> dict:
    # .xls requires xlrd; .xlsx uses openpyxl
    engine = "xlrd" if filepath.lower().endswith(".xls") else "openpyxl"
    df = pd.read_excel(filepath, engine=engine)
    df.columns = [str(c).strip() for c in df.columns]

    id_col = _find_col(df, ["I.D.", "ID", "Item ID", "Instrument ID", "ItemID", "Cal ID"])
    company_col = _find_col(df, ["Company", "Customer", "Client", "Company Name"])
    type_col = _find_col(df, ["Type", "Instrument Type", "Cal Type"])
    desc_col = _find_col(df, ["Description", "Desc", "Item Description"])
    serial_col = _find_col(df, ["Serial Number", "Serial No", "Serial", "SN", "S/N"])
    wo_col = _find_col(df, ["Work Order", "WO", "WO Number", "WorkOrder"])
    received_col = _find_col(df, ["Received On", "Received Date", "Date Received", "Receive Date"])
    mfr_col = _find_col(df, ["Manufacturer", "Mfr", "Make"])
    model_col = _find_col(df, ["Model Number", "Model No", "Model", "ModelNumber"])

    if id_col is None:
        raise ValueError("Current WIP: cannot find an I.D. column.")

    tier_counts = {"id": 0, "mfr_model": 0, "type": 0, "unmatched": 0}
    total_value = 0.0
    items = []

    for _, row in df.iterrows():
        iid = str(row[id_col]).strip() if pd.notna(row.get(id_col)) else ""
        mfr = str(row[mfr_col]).strip() if mfr_col and pd.notna(row.get(mfr_col)) else ""
        model = str(row[model_col]).strip() if model_col and pd.notna(row.get(model_col)) else ""
        itype = str(row[type_col]).strip() if type_col and pd.notna(row.get(type_col)) else ""
        wo = str(row[wo_col]).strip() if wo_col and pd.notna(row.get(wo_col)) else ""

        if not iid or iid.lower() in ("nan", "none", ""):
            continue

        cost, tier = lookup.lookup(iid, mfr or None, model or None, itype or None)
        tier_counts[tier] += 1
        if cost:
            total_value += cost

        items.append(
            {
                "id": iid,
                "company": str(row[company_col]).strip() if company_col and pd.notna(row.get(company_col)) else "",
                "type": itype,
                "wo": wo,
                "cost": cost or 0.0,
                "tier": tier,
            }
        )

    unique_wos = len({item["wo"] for item in items if item["wo"] and item["wo"].lower() not in ("nan", "none", "")})

    return {
        "item_count": len(items),
        "wo_count": unique_wos,
        "total_value": total_value,
        "unmatched_count": tier_counts["unmatched"],
        "pricing_confidence": tier_counts,
        "items": items,
    }


def _find_col(df, candidates):
    cols_upper = {c.upper(): c for c in df.columns}
    for c in candidates:
        if c.upper() in cols_upper:
            return cols_upper[c.upper()]
    return None
