import pandas as pd
import sys
from pathlib import Path
from xlsx_to_df import main as load_excel
from config import PROCESSING_DIR

COLUMNS_TO_KEEP = [
    "Lane Id",
    "Origin City",
    "Origin postal code",
    "Origin country code",
    "Origin Country",
    "Origin region",
    "Destination city",
    "Destination postal code",
    "destination country code",
    "Destination country",
    "Destination region",
    "Payer region",
    "Proposed Origin Port",
    "Proposed Destination Port",
    "Service",
    "Equipment type",
]

RENAME_MAP = {
    "Proposed Origin Port": "Origin port",
    "Proposed Destination Port": "Destination port",
}


def match_column(df_columns: list[str], target: str) -> str | None:
    """Find the actual column name in df that matches target (case-insensitive)."""
    target_lower = target.lower()
    for col in df_columns:
        if str(col).strip().lower() == target_lower:
            return col
    return None


def build_column_map(df: pd.DataFrame) -> dict[str, str]:
    """Map each desired column name to its actual name in the DataFrame (case-insensitive)."""
    col_map = {}
    for desired in COLUMNS_TO_KEEP:
        actual = match_column(list(df.columns), desired)
        if actual is not None:
            col_map[desired] = actual
    return col_map


def get_effective_date(xlsx: pd.ExcelFile) -> str:
    """Read the Revision tab, find the last Effective date, return as DD.MM.YYYY.

    Scans all rows/columns for the header because it may not be on the first row.
    """
    revision_tab = None
    for name in xlsx.sheet_names:
        if name.strip().lower() == "revision":
            revision_tab = name
            break
    if revision_tab is None:
        print("Error: 'Revision' tab not found in the workbook")
        sys.exit(1)

    raw = pd.read_excel(xlsx, sheet_name=revision_tab, header=None)

    header_row = None
    header_col = None
    for row_idx in range(len(raw)):
        for col_idx in range(len(raw.columns)):
            cell = str(raw.iloc[row_idx, col_idx]).strip().lower()
            if "effective" in cell and "date" in cell:
                header_row = row_idx
                header_col = col_idx
                break
        if header_row is not None:
            break

    if header_row is None:
        print("Error: 'Effective date' column not found anywhere in the Revision tab")
        print(f"  Raw content preview:\n{raw.head(10)}")
        sys.exit(1)

    date_values = raw.iloc[header_row + 1:, header_col].dropna()
    if date_values.empty:
        print("Error: No date values found under 'Effective date' in Revision tab")
        sys.exit(1)

    last_date = pd.to_datetime(date_values.iloc[-1])
    
    return last_date.strftime("%d.%m.%Y")


def find_pre_carriage_columns(df: pd.DataFrame) -> list[str]:
    """Find columns from 'pre carriage' / 'pre-carriage' onwards (case-insensitive)."""
    all_cols = list(df.columns)
    start_idx = None
    for i, col in enumerate(all_cols):
        col_lower = str(col).lower().replace("-", " ")
        if "pre carriage" in col_lower:
            start_idx = i
            break

    if start_idx is None:
        print("Warning: No column containing 'pre carriage' / 'pre-carriage' found — no extra columns added")
        return []

    return all_cols[start_idx:]


def process(df: pd.DataFrame, xlsx: pd.ExcelFile) -> pd.DataFrame:
    effective_date = get_effective_date(xlsx)
    print(f"  Effective date: {effective_date}")

    col_map = build_column_map(df)

    missing = [c for c in COLUMNS_TO_KEEP if c not in col_map]
    if missing:
        print(f"Warning: Missing columns in data: {missing}")

    actual_cols = [col_map[c] for c in COLUMNS_TO_KEEP if c in col_map]
    result = df[actual_cols].copy()

    actual_rename = {}
    for old_name, new_name in RENAME_MAP.items():
        if old_name in col_map:
            actual_rename[col_map[old_name]] = new_name
    result.rename(columns=actual_rename, inplace=True)

    desired_to_actual = {desired: col_map[desired] for desired in col_map}
    actual_to_output = {}
    for desired, actual in desired_to_actual.items():
        if actual in actual_rename:
            actual_to_output[desired] = actual_rename[actual]
        else:
            actual_to_output[desired] = actual

    carrier_col_actual = match_column(list(df.columns), "Aptiv Preferred Carrier")
    dest_code_actual = col_map.get("destination country code")
    if carrier_col_actual is not None and dest_code_actual is not None:
        result["Carrier Name"] = df[carrier_col_actual].astype(str) + " " + df[dest_code_actual].astype(str)
    elif carrier_col_actual is not None:
        result["Carrier Name"] = df[carrier_col_actual].astype(str)
        print("Warning: 'destination country code' not found — Carrier Name set without country code")
    else:
        result["Carrier Name"] = ""
        print("Warning: 'Aptiv Preferred Carrier' column not found")

    result["Valid from"] = effective_date
    result["Valid to"] = effective_date

    pre_carriage_cols = find_pre_carriage_columns(df)
    if pre_carriage_cols:
        
        for col in pre_carriage_cols:
            result[col] = df[col].values

    output_col_names = [actual_to_output.get(c, c) for c in COLUMNS_TO_KEEP if c in col_map]
    carrier_idx = output_col_names.index("Service") + 1 if "Service" in output_col_names else len(output_col_names)
    output_col_names.insert(carrier_idx, "Carrier Name")
    equip_idx = output_col_names.index("Carrier Name") + 1
    output_col_names.insert(equip_idx, "Valid from")
    output_col_names.insert(equip_idx + 1, "Valid to")
    equip_col = actual_to_output.get("Equipment type")
    if equip_col and equip_col in output_col_names:
        output_col_names.remove(equip_col)
        output_col_names.insert(output_col_names.index("Valid to") + 1, equip_col)

    col_order = output_col_names + pre_carriage_cols
    result = result[[c for c in col_order if c in result.columns]]

    return result


def save_output(df: pd.DataFrame, original_name: str):
    PROCESSING_DIR.mkdir(exist_ok=True)
    stem = Path(original_name).stem
    output_path = PROCESSING_DIR / f"{stem}_processed.xlsx"
    df.to_excel(output_path, index=False)
    print(f"  Saved → {output_path}")


if __name__ == "__main__":
    df, xlsx, file_path = load_excel()

    result = process(df, xlsx)

    print(f"  Processed — {len(result)} rows x {len(result.columns)} cols")

    save_output(result, file_path.name)
