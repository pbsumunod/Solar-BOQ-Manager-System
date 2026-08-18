# Workflow: Deploy the BOQ App for Shared Access

## Objective
Get the BOQ Manager (`tools/boq_app.py`) running on a free, always-on URL
that people other than you can open, gated by a shared password, with data
that survives restarts/redeploys (unlike local Excel files on a hosted
container's ephemeral disk).

## Why Google Sheets instead of local files, once deployed
Locally, the app stores each project and the materials catalog as `.xlsx`
files under `data/`. That's fine on your own machine, but free hosting
platforms (including Streamlit Community Cloud) run the app in a container
that can restart, redeploy, or sleep at any time -- local files written
during a session are **not guaranteed to survive** that. So for a deployed,
shared version, project/catalog data instead lives in Google Sheets, which
is genuinely persistent and lets everyone see the same live data.

The app supports **both** storage backends and picks automatically:
local `.xlsx` files if no Google secrets are configured (e.g. running on
your own machine the normal way), or Google Sheets if they are (see
`tools/boq_app.py`'s `_select_storage_backend()`). You don't need Google
Sheets set up just to keep using the app locally.

## Why OAuth, not a service account
A Google **service account** (the usual "headless API access" pattern) was
considered and rejected here: service accounts have no personal Drive
storage quota of their own, and commonly hit a "Service Accounts do not
have storage quota" error when creating new files -- a real, easy-to-hit
dead end for exactly this use case. Using OAuth against a real Google
account instead means every project/catalog spreadsheet is owned by that
account's normal Drive storage, with no quota surprises, and you can open
any of them directly in Google Sheets to double check.

## One-time setup (do this once, from your own machine)

### 1. Create a Google Cloud OAuth Client
1. Go to [console.cloud.google.com](https://console.cloud.google.com/) and create a project (or reuse one).
2. Enable the **Google Sheets API** and **Google Drive API** for that project (APIs & Services → Library).
3. APIs & Services → Credentials → Create Credentials → OAuth client ID.
   - If prompted, configure the OAuth consent screen first (External is fine; you can leave it in "Testing" mode and just add your own Google account as a test user -- no need to publish it).
   - Application type: **Desktop app**.
4. Download the JSON and save it as `credentials.json` in the project root (`/Users/pamelasumunod/Desktop/Projects/Solar System/credentials.json`). It's already gitignored.

### 2. Run the setup script
```
cd "Solar System"
.venv/bin/python tools/setup_google_auth.py
```
This opens your browser for a one-time Google sign-in/consent, then:
- saves `token.json` locally (gitignored),
- creates a "BOQ Projects" Drive folder and a "BOQ Materials Catalog"
  spreadsheet (with its `Catalog`/`Stores`/`Categories`/`Material List` tabs,
  each seeded from your current local data) if they don't already exist --
  safe to re-run, and also fills in any of those four tabs that are missing
  from an existing spreadsheet (e.g. an older catalog spreadsheet from
  before multi-store pricing existed),
- offers to copy any existing local projects (e.g. `sample-project.xlsx`)
  into new Google Sheets, so nothing gets stranded,
- prints a ready-to-paste secrets block (refresh token, client id/secret,
  folder/spreadsheet IDs, and a placeholder for your shared password).

**Save that printed block somewhere** -- you'll paste it into Streamlit
Cloud's Secrets in step 4. Pick a real value for `app_password` before
pasting (that's the password everyone will use to open the app).

### 3. Push the code to GitHub
Streamlit Community Cloud deploys from a GitHub repo.
1. Create a new repo on [github.com](https://github.com/new) (public or
   private both work). Don't initialize it with a README.
2. Tell me the repo's URL and I'll add it as a remote and push what's here
   (or run these yourself):
   ```
   git remote add origin <your-repo-url>
   git branch -M main
   git push -u origin main
   ```
   Nothing sensitive gets pushed -- `credentials.json`, `token.json`,
   `PACKAGES.xlsx`, `data/`, and `.streamlit/secrets.toml` are all
   gitignored already.

### 4. Deploy on Streamlit Community Cloud
1. Go to [share.streamlit.io](https://share.streamlit.io/) and sign in with
   GitHub.
2. "Create app" → pick your repo/branch → set **Main file path** to
   `tools/boq_app.py`.
3. Before/after deploying, open the app's **Settings → Secrets** and paste
   the block from step 2 (with your real `app_password`).
4. Deploy. First build takes a couple of minutes (installing
   `requirements.txt`).
5. In the app's Settings, you can also set it to **private** if you want an
   extra layer beyond the password gate -- with a shared password already in
   place, this is optional, not required.

## Ongoing: updating the deployed app
Push new commits to the connected branch; Streamlit Cloud redeploys
automatically. No redeploy is needed for data changes (those go straight to
Google Sheets) -- only for code changes.

## Verification
1. Locally, with no `.streamlit/secrets.toml`: app boots straight into the
   BOQ Manager (no password prompt), uses local `data/` files -- unchanged
   from before this workflow existed.
2. Locally, with a `.streamlit/secrets.toml` containing just `app_password`:
   app shows the password prompt before anything else; wrong password shows
   an error and doesn't proceed; correct password unlocks the rest of the
   app for that browser session.
3. With the full secrets block (Google + password) configured, either
   locally or on Streamlit Cloud: "New project" creates a real Google
   Sheets spreadsheet in the "BOQ Projects" folder (check it directly in
   Google Drive), editing/saving updates that spreadsheet, and the
   Materials Catalog tab reads/writes the shared "BOQ Materials Catalog"
   spreadsheet.
4. Two browser sessions (e.g. your laptop + your phone, or two people)
   both editing: last write wins on Save/Save catalog -- there's no
   conflict merging, so if two people save the same project within
   moments of each other, the later save overwrites the earlier one's
   changes. Fine for the expected small-team, one-editor-at-a-time usage;
   worth knowing if usage grows.

## Things learned / to revisit
- Free-tier Streamlit Community Cloud apps can go to sleep after a period
  of inactivity and take a few seconds to wake back up on the next visit --
  normal, not a bug.
- `gspread.Client` has no built-in folder-creation helper (it's
  Sheets-focused); `find_or_create_projects_folder()` in
  `tools/setup_google_auth.py` uses `google-api-python-client`'s Drive v3
  API directly for that one operation instead of reaching into gspread's
  undocumented internals.
- The Drive file ID doubles as the project "slug" in the Google Sheets
  backend (`gsheets_storage.py`) -- Drive already guarantees uniqueness, so
  there's no need to reimplement the local backend's slugify/collision
  logic for this backend.
- Schema changes that touch already-deployed, already-populated live data
  (e.g. the multi-store catalog rework) get their own one-time migration
  script (`tools/migrate_multi_store_catalog.py`) rather than being folded
  into `setup_google_auth.py` -- see `workflows/manage_boq_inventory.md`'s
  "Materials catalog" section for what it does and why it must run against
  live data *before* the corresponding code gets deployed, not after.
