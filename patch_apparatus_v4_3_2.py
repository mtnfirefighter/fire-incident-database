# Monkey‑patch for v4.3.2 — Apparatus roster → report picker (CallSign-first)
# Usage in app.py (after imports):
#   from patch_apparatus_v4_3_2 import apply_patch
#   apply_patch(globals())
#
# Optional: after saving rosters to Excel, refresh in‑memory tables so pickers update:
#   from patch_apparatus_v4_3_2 import refresh_rosters_after_save
#   data = refresh_rosters_after_save(data, file_path, load_workbook, normalize_df)

from typing import Dict, List
import pandas as pd

WANTED_APPARATUS_SCHEMA = [
    "ApparatusID",
    "UnitNumber",
    "UnitType",
    "SeatingCapacity",
    "GPM",
    "TankSize",
    "Active",
    "Name",
]

def _safe_normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    for c in df.columns:
        df[c] = df[c].apply(lambda v: v.strip() if isinstance(v, str) else v)
    return df

def _pick_series_by_names(df: pd.DataFrame, names_in_priority: List[str]):
    # Case-insensitive, space-insensitive column match
    norm_map = {str(c): str(c).strip().lower() for c in df.columns}
    inv = {v:k for k,v in norm_map.items()}
    for name in names_in_priority:
        key = name.strip().lower()
        if key in inv:
            col = inv[key]
            if df[col].notna().any():
                return df[col].astype(str).str.strip()
    return None

def _build_unit_options(df: pd.DataFrame) -> list:
    df = _safe_normalize_df(df)
    if df is None or df.empty:
        return []

    # Prefer Active=Yes if it exists
    try:
        if "Active" in df.columns:
            active_mask = df["Active"].astype(str).str.lower().isin(["yes","true","1"])
            df_use = df[active_mask]
            if df_use.empty:
                df_use = df
        else:
            df_use = df
    except Exception:
        df_use = df

    # Priority order: CallSign (any case) first, then common alternates
    priority = ["callsign", "call sign", "unitnumber", "unit", "unit #", "unit_number", "name", "apparatus", "truck"]

    buckets = []
    # Try primary column first
    s = _pick_series_by_names(df_use, priority)
    if s is not None:
        buckets.append(s)

    # Also gather alternates to be safe (dedupe later)
    alternates = ["unitnumber","unit","unit #","unit_number","call sign","name","apparatus","truck"]
    for alt in alternates:
        ss = _pick_series_by_names(df_use, [alt])
        if ss is not None:
            buckets.append(ss)

    if not buckets:
        return []

    s_all = pd.concat(buckets, ignore_index=True)
    vals = (s_all.dropna()
                 .map(lambda x: x.strip())
                 .replace("", pd.NA)
                 .dropna()
                 .unique()
                 .tolist())
    return sorted(set(vals))

def refresh_rosters_after_save(data: Dict[str, pd.DataFrame], file_path: str, load_workbook, normalize_df):
    try:
        re = load_workbook(file_path)
        if "Personnel" in re:
            data["Personnel"] = normalize_df(re["Personnel"])
        if "Apparatus" in re:
            data["Apparatus"] = normalize_df(re["Apparatus"])
        return data
    except Exception:
        return data

def apply_patch(env: dict):
    # 1) Ensure schema includes desired columns (do not remove existing ones)
    try:
        schema = env.get("APPARATUS_SCHEMA", [])
        for c in WANTED_APPARATUS_SCHEMA:
            if c not in schema:
                schema.append(c)
        env["APPARATUS_SCHEMA"] = schema
    except Exception:
        env["APPARATUS_SCHEMA"] = WANTED_APPARATUS_SCHEMA[:]

    # 2) Install safe normalizer if app didn't define one
    if "normalize_df" not in env:
        env["normalize_df"] = _safe_normalize_df

    # 3) Override unit option builder (CallSign-first, case-insensitive)
    env["build_unit_options"] = _build_unit_options
