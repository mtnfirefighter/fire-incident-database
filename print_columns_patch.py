# print_columns_patch.py
# Print-tab-only column selection (safe, read-only)

def personnel_print_columns(df):
    wanted = [
        "PersonnelID",
        "Name",
        "Role",
        "Hours",
        "RespondedIn",
    ]
    return df[[c for c in wanted if c in df.columns]]


def apparatus_print_columns(df):
    wanted = [
        "ApparatusID",
        "UnitType",
        "Role",
        "Actions",
    ]
    return df[[c for c in wanted if c in df.columns]]