"""One-time, idempotent migration to the multi-store catalog schema.

What this does:
  1. Renames the catalog's "Default Unit Cost (₱)" column to "Unit Cost (₱)"
     and backfills every existing catalog row's "Store Name" to "Plug and
     Go" (the store the original 50-item catalog came from).
  2. Seeds the Stores / Categories reference data from the catalog's
     current distinct values (Stores gets "Plug and Go" with a blank
     Address -- fill it in later via "Manage stores"). The Category<->
     Material mapping itself needs no separate seeding or backfill: it's
     derived directly from the catalog's own Category/Material columns,
     not stored anywhere else.
  3. Backfills "Store Name" = "Plug and Go" on every existing Materials row
     of the local sample-project (and, with --sheets, the live "6.6 kWp PV
     rooftop project").
  4. With --sheets: does the equivalent against the live Google Sheets
     catalog spreadsheet and project, via the same token.json OAuth flow
     tools/setup_google_auth.py already uses.

Safe to re-run any number of times: every write checks current state first
(column presence, blank-ness) rather than assuming "ran before" from any
single signal, so a partial or repeated run is a no-op for whatever's
already migrated. Verify this yourself: run once, then run again
immediately and confirm it reports "already up to date" everywhere.

CRITICAL ORDERING NOTE: run this BEFORE deploying the corresponding code
changes (the new CATALOG_COLUMNS/CORE_MATERIAL_COLUMNS). This script's
migrate_catalog_df_raw()/migrate_materials_df_raw() always read raw
(pd.read_excel / worksheet.get_all_values, literal old-header-name checks)
rather than through catalog.load_catalog()/gsheets_storage.load_catalog()
-- once the code is updated, those functions already expect
"Unit Cost (₱)" and would read pre-migration data's
"Default Unit Cost (₱)" column as entirely missing, silently defaulting
every price to 0.0. Migrate the data first, deploy the code second.

Usage:
    .venv/bin/python tools/migrate_multi_store_catalog.py            # local files only
    .venv/bin/python tools/migrate_multi_store_catalog.py --sheets   # also migrate live Google Sheets

Back up data/materials_catalog.xlsx and data/projects/*.xlsx yourself
before running this the first time -- this script does not do it for you.
"""

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import pandas as pd

from lib import boq_data, catalog, gsheets_storage

PROJECT_ROOT = Path(__file__).resolve().parent.parent
TOKEN_PATH = PROJECT_ROOT / "token.json"
LIVE_PROJECT_NAME = "6.6 kWp PV rooftop project"

DEFAULT_STORE = "Plug and Go"
OLD_COST_COLUMN = "Default Unit Cost (₱)"
NEW_COST_COLUMN = "Unit Cost (₱)"


def _backfill_store_name(df: pd.DataFrame) -> tuple[pd.DataFrame, bool]:
    """Shared by both catalog and materials migration: add Store Name if
    missing, or fill in any genuinely blank cells if the column already
    exists. Never touches a cell that already has a real value -- re-running
    after someone hand-edited a Store Name shouldn't clobber it."""
    changed = False
    if "Store Name" not in df.columns:
        df["Store Name"] = DEFAULT_STORE
        changed = True
    else:
        blank = df["Store Name"].isna() | (df["Store Name"].astype(str).str.strip() == "")
        if blank.any():
            df.loc[blank, "Store Name"] = DEFAULT_STORE
            changed = True
    return df, changed


def migrate_catalog_df_raw(df: pd.DataFrame) -> tuple[pd.DataFrame, bool]:
    changed = False
    if OLD_COST_COLUMN in df.columns and NEW_COST_COLUMN not in df.columns:
        df = df.rename(columns={OLD_COST_COLUMN: NEW_COST_COLUMN})
        changed = True
    df, store_changed = _backfill_store_name(df)
    return df, (changed or store_changed)


def migrate_materials_df_raw(df: pd.DataFrame) -> tuple[pd.DataFrame, bool]:
    return _backfill_store_name(df)


def seed_reference_lists(catalog_df: pd.DataFrame, store_module) -> None:
    """store_module is whichever backend (catalog or gsheets_storage) has
    load_stores/save_stores/load_category_list/etc -- both expose the same
    function names, so this works unmodified against either."""
    stores_df = store_module.load_stores()
    if DEFAULT_STORE.lower() not in {s.lower() for s in stores_df["Store Name"]}:
        stores_df = pd.concat(
            [stores_df, pd.DataFrame([{"Store Name": DEFAULT_STORE, "Address": ""}])], ignore_index=True
        )
        store_module.save_stores(stores_df)
        print(f'  Seeded Stores with "{DEFAULT_STORE}" (blank Address -- fill in via Manage stores).')

    # Add whatever categories are genuinely new (never touches/removes
    # anything already there). The Category<->Material mapping itself
    # needs no seeding -- it's derived directly from the catalog.
    cat_list = store_module.load_category_list()
    new_cat_list = catalog.sync_category_list_from_catalog(catalog_df, cat_list)
    if len(new_cat_list) != len(cat_list):
        store_module.save_category_list(new_cat_list)
        print(f"  Seeded {len(new_cat_list) - len(cat_list)} new categor{'y' if len(new_cat_list) - len(cat_list) == 1 else 'ies'}.")


def migrate_local() -> None:
    print("=== Local files ===")

    if catalog.CATALOG_PATH.exists():
        raw = pd.read_excel(catalog.CATALOG_PATH, sheet_name="Catalog", engine="openpyxl")
    else:
        raw = catalog.blank_catalog_df()
    migrated, changed = migrate_catalog_df_raw(raw)
    catalog.save_catalog(migrated)  # normalizes to the (already code-updated) new schema regardless
    reloaded = catalog.load_catalog()
    print(f"Local catalog: {'migrated' if changed else 'already up to date'} ({len(reloaded)} rows).")
    seed_reference_lists(reloaded, catalog)

    project_path = boq_data.PROJECTS_DIR / "sample-project.xlsx"
    if project_path.exists():
        raw_materials = pd.read_excel(project_path, sheet_name="Materials", engine="openpyxl")
        migrated_materials, changed = migrate_materials_df_raw(raw_materials)
        if changed:
            # validate_and_normalize_materials (not just recompute_material_totals)
            # so the saved file gets proper column order/dtypes, not just the
            # backfilled column tacked onto wherever pandas happened to put it.
            migrated_materials, _ = boq_data.validate_and_normalize_materials(migrated_materials)
            _, expenses_df, meta = boq_data.load_project("sample-project")
            boq_data.save_project("sample-project", migrated_materials, expenses_df, meta)
            print(f"Local sample-project: backfilled Store Name on {len(migrated_materials)} row(s).")
        else:
            print("Local sample-project: already up to date.")
    else:
        print("Local sample-project.xlsx not found -- skipped.")


def migrate_sheets() -> None:
    print("\n=== Live Google Sheets ===")
    if not TOKEN_PATH.exists():
        print("No token.json found -- skipping live Google Sheets migration.")
        return

    import gspread
    from google.oauth2.credentials import Credentials

    from setup_google_auth import find_or_create_catalog_spreadsheet, find_or_create_projects_folder
    from googleapiclient.discovery import build

    token_info = json.loads(TOKEN_PATH.read_text())
    credentials = Credentials(
        token=None,
        refresh_token=token_info["refresh_token"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=token_info["client_id"],
        client_secret=token_info["client_secret"],
        scopes=gsheets_storage.SCOPES,
    )
    client = gspread.authorize(credentials)
    drive = build("drive", "v3", credentials=credentials)

    # Reuse the existing idempotent find-or-create helpers instead of
    # duplicating tab-creation/recovery logic -- they already handle "the
    # spreadsheet exists but a tab is missing" and ensure all four catalog
    # tabs exist before we touch anything.
    folder_id = find_or_create_projects_folder(drive)
    catalog_id = find_or_create_catalog_spreadsheet(client)
    gsheets_storage.configure(token_info, folder_id, catalog_id)

    sh = client.open_by_key(catalog_id)
    ws = sh.worksheet(gsheets_storage.CATALOG_SHEET)
    raw_values = ws.get_all_values()
    raw_df = pd.DataFrame(raw_values[1:], columns=raw_values[0]) if len(raw_values) > 1 else catalog.blank_catalog_df()
    migrated, changed = migrate_catalog_df_raw(raw_df)
    if changed:
        gsheets_storage.save_catalog(migrated)
        reloaded = gsheets_storage.load_catalog()
        print(f"Live catalog spreadsheet: migrated ({len(reloaded)} rows).")
    else:
        reloaded = gsheets_storage.load_catalog()
        print("Live catalog spreadsheet: already up to date.")
    seed_reference_lists(reloaded, gsheets_storage)

    files = client.list_spreadsheet_files(folder_id=folder_id)
    match = next((f for f in files if f["name"] == LIVE_PROJECT_NAME), None)
    if match is None:
        print(f'Live project "{LIVE_PROJECT_NAME}" not found -- skipped.')
        return

    proj_sh = client.open_by_key(match["id"])
    proj_ws = proj_sh.worksheet(gsheets_storage.MATERIALS_SHEET)
    raw_proj_values = proj_ws.get_all_values()
    raw_materials = (
        pd.DataFrame(raw_proj_values[1:], columns=raw_proj_values[0])
        if len(raw_proj_values) > 1
        else boq_data.blank_materials_df()
    )
    migrated_materials, changed = migrate_materials_df_raw(raw_materials)
    if changed:
        migrated_materials, _ = boq_data.validate_and_normalize_materials(migrated_materials)
        gsheets_storage._write_df(proj_ws, migrated_materials)
        print(f"Live project: backfilled Store Name on {len(migrated_materials)} row(s).")
    else:
        print("Live project: already up to date.")


def main() -> None:
    migrate_local()
    if "--sheets" in sys.argv[1:]:
        migrate_sheets()
    else:
        print("\nSkipped live Google Sheets migration (pass --sheets to include it).")


if __name__ == "__main__":
    main()
