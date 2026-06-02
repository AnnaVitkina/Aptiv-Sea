"""Load accessorial charge tabs and append them to the output workbook."""

import math
import re
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

ACCESSORIAL_OUTPUT_COLS = [
    "Source tab",
    "Rate cost name",
    "Rate Card Name",
    "Multiplier",
    "Rate by",
    "Currency",
    "Price",
    "Apply if",
]

FLOW_ACCESSORIAL_CONFIG = {
    "1": {"multiplier": ""},
    "2": {"multiplier": "Quantity/Percentage"},
    "3": {"multiplier": "Quantity/Percentage"},
}

CHARGE_HEADER_ALIASES = ("charge item", "charge heads")
LANE_HEADER_ALIASES = ("lane id",)
MEASUREMENT_HEADER_ALIASES = ("unit price", "unit / currency", "unit currency")

GENERIC_UNIT_VALUES = {
    "per entry",
    "per fcr",
    "per trk move/hour",
    "per trk move/30 mins",
    "per booking",
    "per so / booking per time",
    "per document",
    "contianer per hr",
    "per shipment",
    "per fcr/hbl/cbl",
}

HEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
FONT_HEADER = Font(bold=True, size=10)
FONT_NORMAL = Font(size=10)
ALIGN_WRAP = Alignment(wrap_text=True, vertical="center", horizontal="center")
ALIGN_LEFT_WRAP = Alignment(wrap_text=True, vertical="top", horizontal="left")

ACCESSORIAL_COLUMN_WIDTHS = {
    "Source tab": 28,
    "Rate cost name": 42,
    "Rate Card Name": 42,
    "Multiplier": 18,
    "Rate by": 22,
    "Currency": 10,
    "Price": 24,
    "Apply if": 36,
}


def normalize(s: str) -> str:
    return re.sub(r"[-_]", " ", str(s).strip().lower())


def normalize_quotes(s: str) -> str:
    return str(s).replace("´", "'").replace("`", "'").replace("'", "'").strip()


def cell_matches_aliases(value, aliases: tuple[str, ...]) -> bool:
    if pd.isna(value):
        return False
    n = normalize(value)
    return any(n == alias or alias in n for alias in aliases)


def find_col(columns: list[str], target: str) -> str | None:
    t = normalize(target)
    for c in columns:
        if normalize(c) == t:
            return c
    return None


def find_col_by_aliases(columns: list[str], aliases: tuple[str, ...]) -> str | None:
    for c in columns:
        n = normalize(c)
        for alias in aliases:
            if n == alias or alias in n:
                return c
    return None


def round_up(value, decimals=3):
    if pd.isna(value):
        return value
    try:
        multiplier = 10 ** decimals
        return math.ceil(float(value) * multiplier) / multiplier
    except (TypeError, ValueError):
        return value


def find_accessorial_header_row(raw: pd.DataFrame, max_rows: int = 50) -> int | None:
    for row_idx in range(min(len(raw), max_rows)):
        for col_idx in range(len(raw.columns)):
            if cell_matches_aliases(raw.iloc[row_idx, col_idx], CHARGE_HEADER_ALIASES):
                return row_idx
    return None


def read_accessorial_sheet(xlsx: pd.ExcelFile, sheet_name: str) -> pd.DataFrame:
    raw = pd.read_excel(xlsx, sheet_name=sheet_name, header=None)
    header_row = find_accessorial_header_row(raw)
    if header_row is None:
        print(f"  Warning: header row not found in '{sheet_name}' — skipping tab")
        return pd.DataFrame()

    headers = []
    for col_idx, value in enumerate(raw.iloc[header_row]):
        if pd.isna(value) or str(value).strip() == "":
            headers.append(f"Unnamed: {col_idx}")
        else:
            headers.append(str(value).strip())

    df = raw.iloc[header_row + 1:].copy()
    df.columns = headers
    return df.dropna(how="all").reset_index(drop=True)


def find_amount_currency_columns(columns: list[str]) -> tuple[str | None, str | None]:
    combined = find_col(columns, "Currency and Amount")
    if combined:
        return combined, None

    for col in columns:
        name = str(col).strip().upper()
        if re.fullmatch(r"[A-Z]{3}", name):
            return col, name

    return None, None


def parse_currency_amount(value, default_currency: str | None = None) -> tuple[str, object]:
    if pd.isna(value):
        return default_currency or "", ""

    text = str(value).strip()
    if not text:
        return default_currency or "", ""

    match = re.match(r"^([\d.,]+)\s*([A-Za-z]{3})$", text)
    if match:
        return match.group(2).upper(), round_up(match.group(1).replace(",", ""))

    match = re.match(r"^([A-Za-z]{3})\s*([\d.,]+)$", text)
    if match:
        return match.group(1).upper(), round_up(match.group(2).replace(",", ""))

    try:
        return default_currency or "", round_up(float(text.replace(",", "")))
    except ValueError:
        return default_currency or "", text


def split_lane_ids(value) -> list[str]:
    if pd.isna(value):
        return [""]
    parts = re.split(r"[,;/]+", str(value).strip())
    return [part.strip() for part in parts if part.strip()]


def is_skippable_charge(value) -> bool:
    if pd.isna(value):
        return True
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return True
    return text.upper().startswith("IF APPLICABLE")


def extract_tab_country(sheet_name: str) -> str:
    match = re.search(r"Charges\s+(\w+)\s+Lane", sheet_name, re.IGNORECASE)
    if match:
        return match.group(1).upper()
    return sheet_name


def split_container_measurements(measurement: str) -> list[str]:
    """Split combined container values such as 40' DRY / 40' HC into separate types."""
    text = normalize_quotes(measurement)
    if not text:
        return [""]

    if re.search(r"40\s*'\s*DRY\s*/\s*40\s*'\s*HC", text, re.IGNORECASE):
        return ["40' DRY", "40' HC"]

    return [text]


def is_generic_per_container(measurement: str) -> bool:
    m = normalize(measurement)
    if m in ("per container", "container"):
        return True
    if "per container" in m:
        return True
    return False


def is_container_type(measurement: str) -> bool:
    m = normalize(normalize_quotes(measurement))
    if not m or m in GENERIC_UNIT_VALUES:
        return False
    if is_generic_per_container(measurement):
        return False
    return bool(re.search(r"20|40|'|dry|hc|\bft\b", m))


def collect_tab_container_types(measurements: list[str]) -> list[str]:
    types: list[str] = []
    seen: set[str] = set()

    for meas in measurements:
        for part in split_container_measurements(str(meas).strip()):
            if not part or not is_container_type(part):
                continue
            key = normalize(part)
            if key not in seen:
                seen.add(key)
                types.append(part)

    return types


def is_tiered_container_price(value) -> bool:
    if pd.isna(value):
        return False
    text = normalize_quotes(str(value))
    return bool(re.search(r"per\s+20\s*['']", text, re.I)) and (
        "days" in text.lower() or "onwards" in text.lower()
    )


def parse_tiered_container_prices(price_text: str) -> list[dict]:
    """Parse tiered per-container prices such as 'Days 4 to 7 $92 per 20' / $114 per 40'."""
    tiers: list[dict] = []
    text = normalize_quotes(str(price_text))

    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue

        range_match = re.search(
            r"Days?\s+(\d+)\s+to\s+(\d+)\s+.*?"
            r"([\d.]+)\s+per\s+20\s*[''].*?"
            r"([\d.]+)\s+per\s+40",
            line,
            re.IGNORECASE,
        )
        if range_match:
            tiers.append({
                "day_from": int(range_match.group(1)),
                "day_to": int(range_match.group(2)),
                "price_20": round_up(range_match.group(3)),
                "price_40": round_up(range_match.group(4)),
            })
            continue

        onwards_match = re.search(
            r"Day\s+(\d+)\s+onwards.*?"
            r"([\d.]+)\s+per\s+20\s*[''].*?"
            r"([\d.]+)\s+per\s+40",
            line,
            re.IGNORECASE,
        )
        if onwards_match:
            tiers.append({
                "day_from": int(onwards_match.group(1)),
                "day_to": None,
                "price_20": round_up(onwards_match.group(2)),
                "price_40": round_up(onwards_match.group(3)),
            })

    return tiers


def tiered_rate_by(charge: str) -> str:
    n = normalize(charge)
    if "demurrage" in n:
        return "Quantity/Demurrage Day"
    if "detention" in n:
        return "Quantity/Detention Day"
    return "Quantity/Day"


def format_tiered_price(tiers: list[dict], container: str) -> str:
    """Format tiered prices for one container size (currency is in Currency column)."""
    price_key = "price_20" if container == "20'" else "price_40"
    lines: list[str] = []

    if tiers:
        lines.append(f"<{tiers[0]['day_from']} - 0")

    for tier in tiers:
        price = tier[price_key]
        price_str = int(price) if float(price) == int(float(price)) else price

        if tier["day_to"] is not None:
            lines.append(f"<={tier['day_to']} {price_str}")
        else:
            lines.append(f">={tier['day_from']} - {price_str}")

    return "\n".join(lines)


def build_tiered_container_rows(
    charge: str,
    sheet_name: str,
    entry: dict,
    flow_multiplier: str,
) -> list[dict]:
    tiers = parse_tiered_container_prices(entry["price_raw"])
    if not tiers:
        return []

    apply_if = build_apply_if(entry["lane_ids"])
    currency = entry["currency"] or "USD"
    rate_by = tiered_rate_by(charge)
    rows: list[dict] = []

    for container_label in ("20'", "40'"):
        rows.append({
            "Source tab": sheet_name,
            "Rate cost name": charge,
            "Rate Card Name": f"{charge.strip()} ({container_label})",
            "Multiplier": flow_multiplier,
            "Rate by": rate_by,
            "Currency": currency,
            "Price": format_tiered_price(tiers, container_label),
            "Apply if": apply_if,
        })

    return rows


def map_rate_card_and_rate_by(
    charge: str,
    measurement: str,
    sheet_name: str,
) -> tuple[str, str]:
    """Return (rate_card_name, rate_by) for a mapped accessorial charge."""
    n = normalize(charge)
    meas = normalize_quotes(measurement) if pd.notna(measurement) else ""
    country = extract_tab_country(sheet_name)

    def card(template: str) -> str:
        return template.format(measurement=meas, country=country)

    if "laden container" in n and "lift" in n:
        return card("Loading/Unloading/Re-entry Fee ({measurement})"), meas

    if "early gate in" in n or "nominate empty container" in n:
        return card("Early Gate In Fee ({measurement})"), meas

    if "milk run" in n:
        return card("Milk-Run Fee({measurement})"), meas

    if "cds liquidity" in n and "green" in n:
        return card("CDS liquidity at port (Green lane) ({measurement})"), meas

    if "cds liquidity" in n and "red" in n:
        return card("CDS liquidity at port ( Red lane) ({measurement})"), meas

    if n == "chb" or n.startswith("chb "):
        return "CHB Fee", "ACC/CHB Fee"

    if "depot storage" in n:
        return card("Depot Storage Fee({measurement})"), meas

    if "additional ccam" in n:
        return "Additional CCAM Fee", "ACC/Additional CCAM Fee"

    if "manual booking" in n:
        return "Booking", meas or "ACC/Booking"

    if "bill amendment" in n:
        return "Bill Amendment Fee", "Condition/Change of billing"

    if "waiting time in china" in n or ("waiting time" in n and "china" in n):
        return "Waiting Time", "Quantity/Hour"

    if "waiting time in destination" in n:
        return f"Waiting Time in Destination ({country})", "Quantity/Hour"

    if "t1 in cz" in n or n.startswith("t1 in cz"):
        return card("T1 (in CZ({measurement}))"), meas

    # Unmapped charges: no Rate Card Name, measurement goes to Rate by
    return "", meas


def build_apply_if(lane_ids: list[str]) -> str:
    lanes: list[str] = []
    seen: set[str] = set()
    for lane_id in lane_ids:
        if lane_id and lane_id not in seen:
            seen.add(lane_id)
            lanes.append(lane_id)
    if not lanes:
        return ""
    return f"LANE ID - Applies if Lane Id equals {', '.join(lanes)}"


def parse_lanes_from_apply_if(apply_if: str) -> list[str]:
    prefix = "LANE ID - Applies if Lane Id equals "
    if not apply_if or not apply_if.startswith(prefix):
        return []
    tail = apply_if[len(prefix):].strip()
    return [part.strip() for part in tail.split(",") if part.strip()]


def consolidate_accessorial_rows(rows: list[dict]) -> list[dict]:
    """Merge rows that differ only by Apply if lane id."""
    groups: dict[tuple, dict] = {}

    for row in rows:
        key = (
            row["Source tab"],
            row["Rate cost name"],
            row["Rate Card Name"],
            row["Multiplier"],
            row["Rate by"],
            row["Currency"],
            row["Price"],
        )
        lanes = parse_lanes_from_apply_if(row["Apply if"])

        if key not in groups:
            grouped = row.copy()
            grouped["_lanes"] = lanes
            groups[key] = grouped
            continue

        existing = groups[key]["_lanes"]
        seen = set(existing)
        for lane in lanes:
            if lane not in seen:
                existing.append(lane)
                seen.add(lane)

    result: list[dict] = []
    for grouped in groups.values():
        lanes = grouped.pop("_lanes")
        grouped["Apply if"] = build_apply_if(lanes)
        result.append(grouped)

    return result


def expand_entry_measurements(
    entry: dict,
    tab_container_types: list[str],
) -> list[dict]:
    """Expand one raw entry across container types and split combined measurements."""
    meas = entry["measurement"]

    if is_generic_per_container(meas):
        if tab_container_types:
            return [{**entry, "measurement": container} for container in tab_container_types]
        return [entry]

    parts = split_container_measurements(meas)
    if len(parts) > 1:
        return [{**entry, "measurement": part} for part in parts]

    return [entry]


def expand_charge_rows(
    row_entries: list[dict],
    tab_container_types: list[str],
) -> list[dict]:
    expanded: list[dict] = []
    for entry in row_entries:
        expanded.extend(expand_entry_measurements(entry, tab_container_types))
    return expanded


def duplicate_with_maersk(rows: list[dict]) -> list[dict]:
    result: list[dict] = []
    for row in rows:
        result.append(row)
        if row["Rate Card Name"]:
            maersk = row.copy()
            maersk["Rate Card Name"] = f"{row['Rate Card Name']} (Maersk)"
            result.append(maersk)
    return result


def transform_accessorial_df(
    df: pd.DataFrame,
    flow: str,
    sheet_name: str,
) -> pd.DataFrame:
    cols = list(df.columns)
    charge_col = find_col_by_aliases(cols, CHARGE_HEADER_ALIASES)
    lane_col = find_col_by_aliases(cols, LANE_HEADER_ALIASES)
    measurement_col = find_col_by_aliases(cols, MEASUREMENT_HEADER_ALIASES)
    amount_col, header_currency = find_amount_currency_columns(cols)

    if not charge_col:
        print("  Warning: charge column not found — skipping tab")
        return pd.DataFrame(columns=ACCESSORIAL_OUTPUT_COLS)

    work = df.copy()
    work[charge_col] = work[charge_col].ffill()

    all_measurements = (
        work[measurement_col].dropna().astype(str).tolist()
        if measurement_col
        else []
    )
    tab_container_types = collect_tab_container_types(all_measurements)

    cfg = FLOW_ACCESSORIAL_CONFIG.get(flow, FLOW_ACCESSORIAL_CONFIG["3"])
    multiplier = cfg["multiplier"]

    charge_groups: dict[str, list[dict]] = {}
    for _, row in work.iterrows():
        charge = row.get(charge_col)
        if is_skippable_charge(charge):
            continue

        measurement = row.get(measurement_col, "") if measurement_col else ""
        if pd.isna(measurement):
            measurement = ""

        currency, price = "", ""
        price_raw = row.get(amount_col) if amount_col else None
        if amount_col:
            currency, price = parse_currency_amount(price_raw, header_currency)

        charge_groups.setdefault(str(charge).strip(), []).append({
            "measurement": str(measurement).strip(),
            "currency": currency,
            "price": price,
            "price_raw": price_raw,
            "lane_ids": split_lane_ids(row.get(lane_col) if lane_col else ""),
        })

    rows: list[dict] = []

    for charge, entries in charge_groups.items():
        tiered_handled = False
        for entry in entries:
            if is_tiered_container_price(entry.get("price_raw")):
                tiered_handled = True
                rows.extend(
                    build_tiered_container_rows(
                        charge, sheet_name, entry, multiplier
                    )
                )
        if tiered_handled:
            continue

        expanded_entries = expand_charge_rows(entries, tab_container_types)

        for entry in expanded_entries:
            meas = normalize_quotes(entry["measurement"])
            rate_card_name, rate_by = map_rate_card_and_rate_by(charge, meas, sheet_name)
            apply_if = build_apply_if(entry["lane_ids"])

            rows.append({
                "Source tab": sheet_name,
                "Rate cost name": charge,
                "Rate Card Name": rate_card_name,
                "Multiplier": multiplier,
                "Rate by": rate_by,
                "Currency": entry["currency"],
                "Price": entry["price"],
                "Apply if": apply_if,
            })

    rows = consolidate_accessorial_rows(rows)
    rows = duplicate_with_maersk(rows)
    return pd.DataFrame(rows, columns=ACCESSORIAL_OUTPUT_COLS)


def prompt_accessorial_sheets(xlsx: pd.ExcelFile) -> list[str] | None:
    answer = input("\nAdd accessorial costs? (y/n): ").strip().lower()
    if answer not in ("y", "yes"):
        return None

    print("\nAvailable sheets for accessorial costs:")
    for i, name in enumerate(xlsx.sheet_names, start=1):
        print(f"  {i}. {name}")

    choice = input("\nEnter sheet numbers or names (comma-separated): ").strip()
    if not choice:
        return None

    selected: list[str] = []
    for part in choice.split(","):
        part = part.strip()
        if not part:
            continue
        if part.isdigit():
            idx = int(part) - 1
            if 0 <= idx < len(xlsx.sheet_names):
                name = xlsx.sheet_names[idx]
                if name not in selected:
                    selected.append(name)
            continue

        if part in xlsx.sheet_names:
            if part not in selected:
                selected.append(part)
            continue

        match = next((n for n in xlsx.sheet_names if n.lower() == part.lower()), None)
        if match and match not in selected:
            selected.append(match)

    return selected or None


def load_accessorial_data(xlsx: pd.ExcelFile, sheet_names: list[str], flow: str) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    for sheet in sheet_names:
        df = read_accessorial_sheet(xlsx, sheet)
        if df.empty:
            print(f"  No accessorial rows from '{sheet}'")
            continue

        transformed = transform_accessorial_df(df, flow, sheet)
        if not transformed.empty:
            frames.append(transformed)
            print(f"  Loaded accessorial tab '{sheet}' — {len(transformed)} rows")
        else:
            print(f"  No accessorial rows from '{sheet}'")

    if not frames:
        return pd.DataFrame(columns=ACCESSORIAL_OUTPUT_COLS)

    return pd.concat(frames, ignore_index=True)


def format_accessorial_sheet(ws, row_count: int) -> None:
    for col_idx, header in enumerate(ACCESSORIAL_OUTPUT_COLS, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = FONT_HEADER
        cell.fill = HEADER_FILL
        cell.alignment = ALIGN_WRAP
        letter = get_column_letter(col_idx)
        ws.column_dimensions[letter].width = ACCESSORIAL_COLUMN_WIDTHS.get(header, 14)

    ws.row_dimensions[1].height = 24
    ws.freeze_panes = "A2"

    for row_idx in range(2, row_count + 2):
        price_cell = ws.cell(row=row_idx, column=ACCESSORIAL_OUTPUT_COLS.index("Price") + 1)
        line_count = str(price_cell.value or "").count("\n") + 1
        ws.row_dimensions[row_idx].height = max(18, 16 * line_count)

        for col_idx in range(1, len(ACCESSORIAL_OUTPUT_COLS) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = FONT_NORMAL
            header = ACCESSORIAL_OUTPUT_COLS[col_idx - 1]
            cell.alignment = (
                ALIGN_LEFT_WRAP
                if header in (
                    "Rate cost name", "Rate Card Name", "Apply if", "Source tab", "Price"
                )
                else ALIGN_WRAP
            )


def append_accessorial_sheet(output_path: Path, df: pd.DataFrame) -> None:
    wb = load_workbook(output_path)
    if "Accessorial Costs" in wb.sheetnames:
        del wb["Accessorial Costs"]

    ws = wb.create_sheet("Accessorial Costs")
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    format_accessorial_sheet(ws, len(df))
    wb.save(output_path)


def maybe_add_accessorial_costs(
    xlsx: pd.ExcelFile,
    output_path: Path | None,
    flow: str,
) -> None:
    if output_path is None or not output_path.exists():
        return

    sheet_names = prompt_accessorial_sheets(xlsx)
    if not sheet_names:
        return

    print(f"  Processing {len(sheet_names)} accessorial tab(s)...")
    accessorial_df = load_accessorial_data(xlsx, sheet_names, flow)
    if accessorial_df.empty:
        print("  No accessorial costs to add")
        return

    append_accessorial_sheet(output_path, accessorial_df)
    print(f"  Added 'Accessorial Costs' tab ({len(accessorial_df)} rows) → {output_path}")
