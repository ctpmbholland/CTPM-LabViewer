"""
Build tiered cost lookup tables from Historic Calibrations.xlsx.

Tiers (in order):
  1. I.D. match
  2. Manufacturer + Model Number match
  3. Type match
"""

import pandas as pd
import numpy as np


class HistoricLookup:
    def __init__(self, id_map: dict, mfr_model_map: dict, type_map: dict):
        self._id = id_map
        self._mfr_model = mfr_model_map
        self._type = type_map

    def lookup(self, instrument_id, manufacturer=None, model_number=None, instrument_type=None):
        """Return (median_cost, tier_label) where tier is 'id', 'mfr_model', 'type', or 'unmatched'."""
        key_id = _norm(instrument_id)
        if key_id and key_id in self._id:
            return self._id[key_id], "id"

        key_mm = (_norm(manufacturer), _norm(model_number))
        if key_mm[0] and key_mm[1] and key_mm in self._mfr_model:
            return self._mfr_model[key_mm], "mfr_model"

        key_type = _norm(instrument_type)
        if key_type and key_type in self._type:
            return self._type[key_type], "type"

        return None, "unmatched"


def build_historic_lookup(filepath: str) -> HistoricLookup:
    engine = "xlrd" if filepath.lower().endswith(".xls") else "openpyxl"
    df = pd.read_excel(filepath, engine=engine)
    df.columns = [str(c).strip() for c in df.columns]

    id_col = _find_col(df, ["I.D.", "ID", "Item ID", "Instrument ID", "ItemID"])
    mfr_col = _find_col(df, ["Manufacturer", "Mfr", "Make"])
    model_col = _find_col(df, ["Model Number", "Model No", "Model", "ModelNumber"])
    type_col = _find_col(df, ["Type", "Instrument Type", "Cal Type"])
    cost_col = _find_col(df, ["Cost", "Price", "Unit Cost", "Per Unit Cost"])

    if cost_col is None:
        raise ValueError("Historic Calibrations: cannot find a Cost column.")

    df["_cost"] = pd.to_numeric(df[cost_col], errors="coerce")
    df = df[df["_cost"] > 0].copy()

    id_map = {}
    if id_col:
        for iid, grp in df.groupby(df[id_col].astype(str).str.strip().str.upper()):
            if iid and iid not in ("NAN", ""):
                id_map[iid] = float(np.median(grp["_cost"]))

    mfr_model_map = {}
    if mfr_col and model_col:
        df["_mm"] = list(
            zip(
                df[mfr_col].astype(str).str.strip().str.upper(),
                df[model_col].astype(str).str.strip().str.upper(),
            )
        )
        for mm, grp in df.groupby("_mm"):
            if mm[0] not in ("NAN", "") and mm[1] not in ("NAN", ""):
                mfr_model_map[mm] = float(np.median(grp["_cost"]))

    type_map = {}
    if type_col:
        for t, grp in df.groupby(df[type_col].astype(str).str.strip().str.upper()):
            if t and t not in ("NAN", ""):
                type_map[t] = float(np.median(grp["_cost"]))

    return HistoricLookup(id_map, mfr_model_map, type_map)


def _find_col(df: pd.DataFrame, candidates: list) -> str | None:
    cols_upper = {c.upper(): c for c in df.columns}
    for c in candidates:
        if c.upper() in cols_upper:
            return cols_upper[c.upper()]
    return None


def _norm(value) -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    return str(value).strip().upper()
