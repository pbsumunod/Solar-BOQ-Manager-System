"""Materials catalog: a reusable, editable list of default materials/prices
that BOQ line items can be pulled from, independent of any single project.

Kept free of any Streamlit imports, same as boq_data.py, so it can be tested
and reused (e.g. by the standalone import script) on its own.
"""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

PROJECT_ROOT = Path(__file__).resolve().parents[2]
CATALOG_PATH = PROJECT_ROOT / "data" / "materials_catalog.xlsx"

CATALOG_COLUMNS = ["Category", "Material", "Brand", "Model", "Unit of Measurement", "Default Unit Cost (₱)"]

# Collapse known spelling/casing variants of the same physical unit down to
# one canonical string, so "pc"/"Pc"/"pcs" (etc) never coexist in the
# catalog as if they were different units.
UNIT_ALIASES = {
    "pc": "pcs",
    "pcs": "pcs",
    "piece": "pcs",
    "pieces": "pcs",
    "mtr": "m",
    "mtrs": "m",
    "meter": "m",
    "meters": "m",
    "metre": "m",
    "metres": "m",
}

# Rows below a "Description"/category cell containing any of these phrases
# are cost-summary/labor lines, not materials -- stop importing a sheet once
# one of these is hit.
STOP_PHRASES = (
    "total material cost",
    "mark up on materials",
    "sign and seal",
    "installation",
    "engineering",
    "transportation",
    "total service cost",
    "total project cost",
)


def _clean(value) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def normalize_unit(value) -> str:
    """Map a unit-of-measurement string to its single canonical spelling."""
    cleaned = _clean(value)
    if not cleaned:
        return ""
    return UNIT_ALIASES.get(cleaned.lower(), cleaned.lower())


def blank_catalog_df() -> pd.DataFrame:
    return pd.DataFrame(columns=CATALOG_COLUMNS)


def normalize_catalog_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in CATALOG_COLUMNS:
        if col not in df.columns:
            df[col] = "" if col != "Default Unit Cost (₱)" else 0.0
    df["Default Unit Cost (₱)"] = pd.to_numeric(df["Default Unit Cost (₱)"], errors="coerce").fillna(0.0).astype(float)
    for col in ("Category", "Material", "Brand", "Model"):
        df[col] = df[col].fillna("").astype(str)
    df["Unit of Measurement"] = df["Unit of Measurement"].apply(normalize_unit)
    return df[CATALOG_COLUMNS].reset_index(drop=True)


def load_catalog() -> pd.DataFrame:
    if not CATALOG_PATH.exists():
        return blank_catalog_df()
    df = pd.read_excel(CATALOG_PATH, sheet_name="Catalog", engine="openpyxl")
    return normalize_catalog_df(df)


def save_catalog(df: pd.DataFrame) -> Path:
    CATALOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    df = normalize_catalog_df(df)
    df.to_excel(CATALOG_PATH, sheet_name="Catalog", index=False, engine="openpyxl")
    return CATALOG_PATH


def parse_package_workbook(path: str | Path) -> pd.DataFrame:
    """Extract catalog rows from a hierarchical package BOQ workbook.

    Expects the layout used by files like PACKAGES.xlsx: a header row
    ("No., Description, Technical Specifications, Brand/Model, Quantity,
    UoM, Price, Unit Cost, Total Cost (PHP)"), followed by material rows
    where the Description column is only filled on the first row of a new
    category and blank on continuation rows for the same category. Stops
    at the first cost-summary row (Total Material Cost, Installation, etc).

    The source "Brand/Model" column is a mix of real brand names, model
    numbers, and "X or equivalent" notes -- this generic parser can't
    reliably tell those apart, so it's dropped into Brand as-is and Model
    is left blank. Review/split Brand vs Model by hand afterward (in the
    Materials Catalog tab) for anything imported this way.
    """
    wb = load_workbook(path, data_only=True)
    records = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        header_row_idx = None
        for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if row and _clean(row[0]).lower() == "no.":
                header_row_idx = i
                break
        if header_row_idx is None:
            continue

        category = ""
        for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
            description = _clean(row[1]) if len(row) > 1 else ""
            if description:
                if any(phrase in description.lower() for phrase in STOP_PHRASES):
                    break
                category = description

            spec = _clean(row[2]) if len(row) > 2 else ""
            brand = _clean(row[3]) if len(row) > 3 else ""
            uom = normalize_unit(row[5]) if len(row) > 5 else ""
            unit_cost = row[7] if len(row) > 7 else None

            material = spec or description
            if not material or not category:
                continue
            if not isinstance(unit_cost, (int, float)):
                continue

            records.append(
                {
                    "Category": category,
                    "Material": material,
                    "Brand": brand,
                    "Model": "",
                    "Unit of Measurement": uom,
                    "Default Unit Cost (₱)": float(unit_cost),
                }
            )

    if not records:
        return blank_catalog_df()

    df = pd.DataFrame(records, columns=CATALOG_COLUMNS)
    df = df.drop_duplicates(subset=["Category", "Material", "Brand", "Unit of Measurement"]).reset_index(drop=True)
    return df


def merge_into_catalog(new_items_df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Merge new catalog rows into the existing catalog, skipping exact
    (Category, Material, Brand, Unit of Measurement) duplicates. Returns
    (merged_df, number_of_rows_added)."""
    existing = load_catalog()
    key_cols = ["Category", "Material", "Brand", "Unit of Measurement"]
    existing_keys = set(map(tuple, existing[key_cols].values.tolist()))

    rows_to_add = [
        row for row in new_items_df.to_dict("records")
        if tuple(row[c] for c in key_cols) not in existing_keys
    ]

    if not rows_to_add:
        return existing, 0

    new_df = pd.DataFrame(rows_to_add, columns=CATALOG_COLUMNS)
    merged = new_df if existing.empty else pd.concat([existing, new_df], ignore_index=True)
    return merged, len(rows_to_add)
