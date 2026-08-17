"""One-time, interactive local setup for the Google Sheets storage backend.

What this does:
  1. Runs an OAuth consent flow in your browser (using credentials.json,
     downloaded from Google Cloud Console -- see workflows/manage_boq_inventory.md
     for how to get one) and saves the resulting token.json locally.
  2. Creates a "BOQ Projects" Drive folder and a "BOQ Materials Catalog"
     spreadsheet if they don't already exist yet (idempotent -- safe to
     re-run; it reuses what it finds instead of duplicating).
  3. If data/materials_catalog.xlsx exists locally, seeds the new Google
     Sheets catalog from it (one-time migration of your current catalog).
  4. Offers to copy any existing local projects (data/projects/*.xlsx) into
     new Google Sheets, so nothing already saved gets stranded.
  5. Prints the exact secrets.toml block to paste into Streamlit Cloud's
     app secrets (or your local .streamlit/secrets.toml for testing).

Usage:
    .venv/bin/python tools/setup_google_auth.py
    .venv/bin/python tools/setup_google_auth.py --yes   # auto-confirm copying local projects, no prompt

Everything this script writes locally (token.json) is already gitignored.
Safe to re-run: reuses an existing token.json (delete it to force a fresh
sign-in) and won't duplicate the Drive folder/catalog spreadsheet if they
already exist.
"""

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from google_auth_oauthlib.flow import InstalledAppFlow
from gspread.exceptions import WorksheetNotFound

from lib import boq_data, catalog, gsheets_storage

PROJECT_ROOT = Path(__file__).resolve().parent.parent
CREDENTIALS_PATH = PROJECT_ROOT / "credentials.json"
TOKEN_PATH = PROJECT_ROOT / "token.json"

PROJECTS_FOLDER_NAME = "BOQ Projects"
CATALOG_SPREADSHEET_NAME = "BOQ Materials Catalog"


def run_oauth_flow() -> dict:
    if TOKEN_PATH.exists():
        token_info = json.loads(TOKEN_PATH.read_text())
        print(f"Reusing existing {TOKEN_PATH} (delete it to force a fresh sign-in).")
        return token_info

    if not CREDENTIALS_PATH.exists():
        print(f"Missing {CREDENTIALS_PATH}.")
        print("Download an OAuth Client ID (type: Desktop app) from Google Cloud")
        print("Console and save it as credentials.json in the project root, then re-run this.")
        sys.exit(1)

    flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_PATH), scopes=gsheets_storage.SCOPES)
    credentials = flow.run_local_server(port=0)

    token_info = {
        "refresh_token": credentials.refresh_token,
        "client_id": credentials.client_id,
        "client_secret": credentials.client_secret,
    }
    TOKEN_PATH.write_text(json.dumps(token_info, indent=2))
    print(f"Saved {TOKEN_PATH}")
    return token_info


def find_or_create_projects_folder(drive) -> str:
    query = (
        f"name = '{PROJECTS_FOLDER_NAME}' and "
        "mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    )
    existing = drive.files().list(q=query, fields="files(id,name)").execute().get("files", [])
    if existing:
        folder_id = existing[0]["id"]
        print(f'Found existing "{PROJECTS_FOLDER_NAME}" folder: {folder_id}')
        return folder_id

    created = drive.files().create(
        body={"name": PROJECTS_FOLDER_NAME, "mimeType": "application/vnd.google-apps.folder"},
        fields="id",
    ).execute()
    folder_id = created["id"]
    print(f'Created "{PROJECTS_FOLDER_NAME}" folder: {folder_id}')
    return folder_id


def _seed_catalog_worksheet(ws) -> int:
    local_catalog = catalog.load_catalog()
    rows = [catalog.CATALOG_COLUMNS] + local_catalog.astype(object).where(local_catalog.notna(), "").values.tolist()
    ws.update(rows, value_input_option="USER_ENTERED")
    return len(local_catalog)


def find_or_create_catalog_spreadsheet(client) -> str:
    existing = client.list_spreadsheet_files(title=CATALOG_SPREADSHEET_NAME)
    if existing:
        spreadsheet_id = existing[0]["id"]
        sh = client.open_by_key(spreadsheet_id)
        try:
            sh.worksheet(gsheets_storage.CATALOG_SHEET)
            print(f'Found existing "{CATALOG_SPREADSHEET_NAME}" (already set up): {spreadsheet_id}')
            return spreadsheet_id
        except WorksheetNotFound:
            # A spreadsheet with the right name exists but is missing its
            # Catalog tab -- almost certainly a leftover half-created file
            # from an earlier run that failed partway through (e.g. the
            # Sheets API wasn't enabled yet). Finish setting it up rather
            # than treating "name matches" as "fully configured".
            print(f'Found existing "{CATALOG_SPREADSHEET_NAME}" but its Catalog tab is missing (likely from an interrupted earlier run) -- fixing it now.')
            ws = sh.add_worksheet(gsheets_storage.CATALOG_SHEET, rows=200, cols=10)
            count = _seed_catalog_worksheet(ws)
            print(f'Added and seeded the Catalog tab ({count} items): {spreadsheet_id}')
            return spreadsheet_id

    sh = client.create(CATALOG_SPREADSHEET_NAME)
    default_ws = sh.sheet1
    ws = sh.add_worksheet(gsheets_storage.CATALOG_SHEET, rows=200, cols=10)
    sh.del_worksheet(default_ws)

    count = _seed_catalog_worksheet(ws)
    print(f'Created "{CATALOG_SPREADSHEET_NAME}" ({count} items seeded from data/materials_catalog.xlsx): {sh.id}')
    return sh.id


def migrate_local_projects(auto_yes: bool = False) -> None:
    """Offer to copy existing local projects (data/projects/*.xlsx) into new
    Google Sheets, so switching backends doesn't strand anything already
    saved. Uses gsheets_storage directly (must already be configure()'d),
    reusing the same create_project logic real app usage goes through."""
    local_projects = boq_data.list_projects()
    if not local_projects:
        return

    names = ", ".join(f'"{p["name"]}"' for p in local_projects)
    if auto_yes:
        print(f"\nFound {len(local_projects)} local project(s) ({names}). Copying into Google Sheets (--yes passed)...")
    else:
        try:
            answer = input(f"\nFound {len(local_projects)} local project(s) ({names}). Copy them into Google Sheets too? [y/N] ")
        except EOFError:
            # No interactive stdin available (e.g. run from a non-interactive
            # shell) -- default to skipping rather than crashing. Re-run with
            # --yes to migrate without the prompt.
            print("\nNo interactive input available, so skipping the local-project-copy prompt.")
            print(f"Found {len(local_projects)} local project(s) ({names}) not yet migrated.")
            print("Re-run with --yes to copy them into Google Sheets non-interactively.")
            return
        if answer.strip().lower() != "y":
            print("Skipped local project migration.")
            return

    already_in_sheets = {p["name"] for p in gsheets_storage.list_projects()}
    for p in local_projects:
        if p["name"] in already_in_sheets:
            print(f'  Skipped "{p["name"]}" -- a project with that name already exists in Google Sheets.')
            continue
        materials_df, expenses_df, meta = boq_data.load_project(p["slug"])
        new_slug = gsheets_storage.create_project(meta.get("project_name") or p["name"], materials_df, expenses_df)
        print(f'  Copied "{p["name"]}" -> new Google Sheets project {new_slug}')


def main() -> None:
    auto_yes = "--yes" in sys.argv[1:]
    token_info = run_oauth_flow()

    import gspread
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build

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

    folder_id = find_or_create_projects_folder(drive)
    catalog_id = find_or_create_catalog_spreadsheet(client)

    gsheets_storage.configure(token_info, folder_id, catalog_id)
    migrate_local_projects(auto_yes=auto_yes)

    print("\n" + "=" * 70)
    print("Setup complete. Paste this into your Streamlit Cloud app's Secrets")
    print("(or .streamlit/secrets.toml for local testing):")
    print("=" * 70)
    print(f"""
app_password = "choose-a-shared-password"

[google_oauth]
refresh_token = "{token_info['refresh_token']}"
client_id = "{token_info['client_id']}"
client_secret = "{token_info['client_secret']}"

[boq]
projects_folder_id = "{folder_id}"
catalog_spreadsheet_id = "{catalog_id}"
""")


if __name__ == "__main__":
    main()
