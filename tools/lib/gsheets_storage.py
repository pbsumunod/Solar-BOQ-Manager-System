"""Google Sheets-backed storage: an alternative to local_storage's Excel
files, for deployments (e.g. Streamlit Community Cloud) that don't have a
persistent local filesystem.

Same public function names/shapes as the local-file functions in
boq_data.py (list_projects, load_project, save_project, create_project) and
catalog.py (load_catalog, save_catalog) so boq_app.py can pick whichever
backend is configured and call it the same way either way. All the actual
schema/validation/derived-value logic (column list, totals, dtype
normalization) is reused from boq_data.py/catalog.py rather than
duplicated -- this module only handles reading/writing rows to Sheets.

Uses OAuth (not a service account) so project files are owned by a real
Google account's normal Drive storage -- see tools/setup_google_auth.py.
Kept free of Streamlit imports: call configure() once with plain data
before using anything else here; boq_app.py is responsible for sourcing
that data from st.secrets or a local token.json.
"""

from __future__ import annotations

from datetime import datetime

import gspread
import pandas as pd
from google.oauth2.credentials import Credentials

from . import boq_data, catalog

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

MATERIALS_SHEET = "Materials"
EXPENSES_SHEET = "Expenses"
META_SHEET = "Meta"
CATALOG_SHEET = "Catalog"

_state: dict = {"client": None, "projects_folder_id": None, "catalog_spreadsheet_id": None}


def configure(token_info: dict, projects_folder_id: str, catalog_spreadsheet_id: str) -> None:
    """One-time setup. token_info needs client_id, client_secret, and
    refresh_token -- the same fields tools/setup_google_auth.py writes to
    token.json (and that get copied into deployment secrets)."""
    credentials = Credentials(
        token=None,
        refresh_token=token_info["refresh_token"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=token_info["client_id"],
        client_secret=token_info["client_secret"],
        scopes=SCOPES,
    )
    _state["client"] = gspread.authorize(credentials)
    _state["projects_folder_id"] = projects_folder_id
    _state["catalog_spreadsheet_id"] = catalog_spreadsheet_id


def is_configured() -> bool:
    return _state["client"] is not None


def _client() -> gspread.Client:
    if _state["client"] is None:
        raise RuntimeError("gsheets_storage.configure() must be called before use.")
    return _state["client"]


def _df_to_rows(df: pd.DataFrame) -> list:
    body = df.astype(object).where(pd.notna(df), "").values.tolist()
    return [list(df.columns)] + body


def _worksheet_to_df(ws, expected_columns: list) -> pd.DataFrame:
    values = ws.get_all_values()
    if len(values) < 2:
        return pd.DataFrame(columns=expected_columns)
    header, *rows = values
    return pd.DataFrame(rows, columns=header)


def _write_df(ws, df: pd.DataFrame) -> None:
    ws.clear()
    ws.update(_df_to_rows(df), value_input_option="USER_ENTERED")


def _meta_from_worksheet(ws) -> dict:
    values = ws.get_all_values()
    meta = {}
    for row in values[1:]:
        if row and row[0]:
            meta[row[0]] = row[1] if len(row) > 1 else ""
    return meta


def _write_meta(ws, meta: dict) -> None:
    ws.clear()
    rows = [["key", "value"]] + [[k, v] for k, v in meta.items()]
    ws.update(rows, value_input_option="USER_ENTERED")


# ---------------------------------------------------------------------------
# Projects -- one Google Spreadsheet per project, three tabs each, all
# living in one Drive folder. The Drive file ID doubles as the project
# "slug" everywhere (Drive already guarantees it's unique, so there's no
# need for our own slugify/collision-avoidance logic here).
# ---------------------------------------------------------------------------

def list_projects() -> list[dict]:
    files = _client().list_spreadsheet_files(folder_id=_state["projects_folder_id"])
    projects = []
    for f in files:
        try:
            meta = _meta_from_worksheet(_client().open_by_key(f["id"]).worksheet(META_SHEET))
        except Exception:
            meta = {}
        projects.append(
            {
                "slug": f["id"],
                "name": meta.get("project_name") or f["name"],
                "path": None,
                "last_modified_at": meta.get("last_modified_at", ""),
            }
        )
    projects.sort(key=lambda p: str(p["name"]).lower())
    return projects


def load_project(slug: str) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    sh = _client().open_by_key(slug)
    materials_df = _worksheet_to_df(sh.worksheet(MATERIALS_SHEET), boq_data.CORE_MATERIAL_COLUMNS)
    expenses_df = _worksheet_to_df(sh.worksheet(EXPENSES_SHEET), boq_data.CORE_EXPENSE_COLUMNS)
    meta = _meta_from_worksheet(sh.worksheet(META_SHEET))
    materials_df, _ = boq_data.validate_and_normalize_materials(materials_df)
    expenses_df = boq_data.validate_and_normalize_expenses(expenses_df)
    return materials_df, expenses_df, meta


def save_project(slug: str, materials_df: pd.DataFrame, expenses_df: pd.DataFrame, meta: dict) -> str:
    sh = _client().open_by_key(slug)
    materials_df = boq_data.recompute_material_totals(materials_df)
    meta = dict(meta)
    meta["last_modified_at"] = datetime.now().isoformat(timespec="seconds")
    meta.setdefault("schema_version", boq_data.SCHEMA_VERSION)
    _write_df(sh.worksheet(MATERIALS_SHEET), materials_df)
    _write_df(sh.worksheet(EXPENSES_SHEET), expenses_df)
    _write_meta(sh.worksheet(META_SHEET), meta)
    return slug


def create_project(
    name: str,
    materials_df: pd.DataFrame | None = None,
    expenses_df: pd.DataFrame | None = None,
) -> str:
    materials_df = (
        boq_data.recompute_material_totals(materials_df) if materials_df is not None else boq_data.blank_materials_df()
    )
    expenses_df = expenses_df if expenses_df is not None else boq_data.blank_expenses_df()

    sh = _client().create(name.strip(), folder_id=_state["projects_folder_id"])
    default_ws = sh.sheet1

    ws_materials = sh.add_worksheet(MATERIALS_SHEET, rows=200, cols=max(len(materials_df.columns) + 2, 10))
    ws_expenses = sh.add_worksheet(EXPENSES_SHEET, rows=100, cols=5)
    ws_meta = sh.add_worksheet(META_SHEET, rows=20, cols=2)
    sh.del_worksheet(default_ws)

    now = datetime.now().isoformat(timespec="seconds")
    meta = {
        "project_name": name.strip(),
        "slug": sh.id,
        "created_at": now,
        "last_modified_at": now,
        "schema_version": boq_data.SCHEMA_VERSION,
    }
    _write_df(ws_materials, materials_df)
    _write_df(ws_expenses, expenses_df)
    _write_meta(ws_meta, meta)
    return sh.id


# ---------------------------------------------------------------------------
# Catalog -- one dedicated Spreadsheet, single "Catalog" tab.
# ---------------------------------------------------------------------------

def load_catalog() -> pd.DataFrame:
    ws = _client().open_by_key(_state["catalog_spreadsheet_id"]).worksheet(CATALOG_SHEET)
    df = _worksheet_to_df(ws, catalog.CATALOG_COLUMNS)
    return catalog.normalize_catalog_df(df)


def save_catalog(df: pd.DataFrame) -> str:
    df = catalog.normalize_catalog_df(df)
    ws = _client().open_by_key(_state["catalog_spreadsheet_id"]).worksheet(CATALOG_SHEET)
    _write_df(ws, df)
    return _state["catalog_spreadsheet_id"]
