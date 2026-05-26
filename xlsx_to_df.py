import pandas as pd
import sys
from pathlib import Path
from config import INPUT_DIR


def main():
    if not INPUT_DIR.exists():
        print(f"Error: Input folder not found: {INPUT_DIR}")
        sys.exit(1)

    excel_files = sorted(
        f for f in INPUT_DIR.iterdir() if f.suffix.lower() in (".xlsx", ".xls")
    )
    if not excel_files:
        print(f"No Excel files found in {INPUT_DIR}")
        sys.exit(1)

    print("Excel files in 'input' folder:")
    for i, f in enumerate(excel_files, start=1):
        print(f"  {i}. {f.name}")

    file_choice = input("\nEnter file number or name: ").strip()

    if file_choice.isdigit():
        idx = int(file_choice) - 1
        if idx < 0 or idx >= len(excel_files):
            print(f"Error: Invalid number. Choose between 1 and {len(excel_files)}")
            sys.exit(1)
        path = excel_files[idx]
    else:
        path = INPUT_DIR / file_choice
        if not path.exists():
            print(f"Error: File '{file_choice}' not found in {INPUT_DIR}")
            sys.exit(1)

    xlsx = pd.ExcelFile(path)
    sheet_names = xlsx.sheet_names

    print(f"\nFound {len(sheet_names)} sheet(s):")
    for i, name in enumerate(sheet_names, start=1):
        print(f"  {i}. {name}")

    choice = input("\nEnter sheet number or name: ").strip()

    if choice.isdigit():
        idx = int(choice) - 1
        if idx < 0 or idx >= len(sheet_names):
            print(f"Error: Invalid sheet number. Choose between 1 and {len(sheet_names)}")
            sys.exit(1)
        sheet = sheet_names[idx]
    elif choice in sheet_names:
        sheet = choice
    else:
        print(f"Error: Sheet '{choice}' not found")
        sys.exit(1)

    df = pd.read_excel(xlsx, sheet_name=sheet)
    print(f"  Loaded '{sheet}' — {len(df)} rows x {len(df.columns)} cols")

    df = clean_df(df)
    print(f"  After cleaning — {len(df)} rows")

    return df, xlsx, path


def clean_df(df: pd.DataFrame) -> pd.DataFrame:
    remove_non_emea = input("\nRemove Non-EMEA rows? (y/n): ").strip().lower()
    if remove_non_emea in ("y", "yes"):
        lane_col = None
        for col in df.columns:
            if str(col).strip().lower() == "lane id":
                lane_col = col
                break
        if lane_col is None:
            print("Warning: 'Lane Id' column not found — skipping EMEA filter")
        else:
            before = len(df)
            df = df[df[lane_col].astype(str).str.contains("EMEA", case=False, na=False)]
            print(f"Removed {before - len(df)} Non-EMEA rows")
    return df


if __name__ == "__main__":
    df, xlsx, _ = main()
