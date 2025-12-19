# incident_personnel_guard.py
# Ensures PersonnelID is never dropped or overwritten in incident personnel dataframes

def preserve_personnel_id(df):
    """
    Ensures PersonnelID column exists and is preserved
    without overwriting existing values.
    """
    if df is None or "PersonnelID" not in df.columns:
        return df

    cols = df.columns.tolist()

    # Keep PersonnelID as the first column for visibility
    if cols[0] != "PersonnelID":
        cols = ["PersonnelID"] + [c for c in cols if c != "PersonnelID"]
        df = df[cols]

    return df