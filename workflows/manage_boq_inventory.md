# Workflow: Manage BOQ / Inventory

## Objective
Give the business a local webpage for creating, editing, and tracking Bills of
Quantities (BOQ) for solar installation jobs — materials, quantities, costs,
suppliers, and other project expenses (logistics, labor, permits) — with
running totals and quick links to manually cross-check material prices on
Shopee, Lazada, and Facebook Marketplace.

## Tool
`tools/boq_app.py` — a Streamlit app (a new kind of "tool": a long-running
local web UI rather than a one-shot script). Its logic is split out into
`tools/lib/boq_data.py` (schemas, validation, template/export generation,
and the local Excel-file storage backend), `tools/lib/catalog.py` (the
reusable materials catalog's schema/validation and local storage),
`tools/lib/gsheets_storage.py` (the Google Sheets storage backend, used when
deployed — see `workflows/deploy_boq_app.md`), and `tools/lib/price_links.py`
(Shopee/Lazada/Facebook Marketplace search URL builders) so the non-UI logic
can be reused or tested independently of Streamlit.
`tools/import_catalog_from_workbook.py` is a standalone script for
seeding/growing the catalog from a package BOQ file.

## Layout
The 📋 BOQ tab keeps the Materials Total / Other Expenses Total / Grand Total
metrics pinned near the top (right below the title, above everything else)
so the bottom line is visible without scrolling past the tables — this uses
a `st.container()` declared early and filled in later in the script, after
both editors below it have captured the run's edits, so the numbers are
never stale. Materials, Other Expenses, and Quick price check are each in
their own bordered card (`st.container(border=True)`) for clear visual
separation instead of relying only on `st.divider()`. The sidebar groups
actions by frequency of use: the active project selector plus Save/Export
sit at the top, then Rename, then the less-frequent New Project / Download
Template actions below.

Success confirmations that are immediately followed by `st.rerun()` (e.g.
"Created project", "Renamed") use `st.toast()`, not `st.success()` —
`st.success()` right before a rerun tends to get cut off before the user
ever sees it, while `st.toast()` is designed to persist across a rerun.

Search-box-plus-button rows (materials search + Columns popover, catalog
search + Save catalog) align the two widgets by giving the search box
`label_visibility="collapsed"` (with the search intent conveyed by a
placeholder instead) rather than padding the button's column with a blank
`st.write("")` spacer — the blank-spacer approach doesn't reliably match a
text input's label-row height, so it drifted out of alignment. Collapsing
the label removes the height mismatch at the source. Use this pattern for
any future search-box-next-to-button row instead of a spacer.

## How to run
```
cd "Solar System"
.venv/bin/streamlit run tools/boq_app.py
```
Opens at `http://localhost:8501` by default. Local-only — no external hosting.

## Inputs
- A project name.
- One of three starting points: a **blank** project, an **uploaded** `.xlsx`
  BOQ file, or **copied from an existing project** (duplicates that
  project's Materials and Expenses as a starting point — pick the source
  from a dropdown of your existing projects). Use the "⬇️ Download blank BOQ
  template" button in the sidebar to get a correctly formatted starting file
  if you don't already have one and aren't copying/uploading. Uploaded files
  don't need to match the template exactly — common header variants (e.g.
  "Qty", "Unit Price") are auto-recognized.
- The downloadable template's Category and Material columns include Excel
  dropdown suggestions built from the current Materials Catalog (via
  `boq_data._add_catalog_dropdowns`, using the in-session catalog if the
  Materials Catalog tab has unsaved edits, else the catalog on disk). These
  are suggestions, not a hard restriction — typing a value that isn't in the
  catalog still works, since a new project may need materials that aren't
  catalogued yet. The dropdown list lives on a hidden `Lists` sheet in the
  workbook (referenced by cell range, not a literal comma list, so category/
  material names containing commas or quotes don't break anything).

## Outputs
- One durable project file per project at `data/projects/<slug>.xlsx`
  (created on "Create project", updated on "💾 Save project"). Each file has
  three sheets: `Materials`, `Expenses`, `Meta`.
- Ad hoc exports via "⬇️ Export current BOQ" (current in-app state, whether
  saved or not) and the blank template.

## Renaming a project
The sidebar's "✏️ Rename project" expander (shown whenever a project is
loaded) lets you change the display name at any time — it updates
`project_name` in the file's `Meta` sheet and saves immediately. The on-disk
filename (slug, e.g. `data/projects/cruz-residence-aug-2026.xlsx`) is not
renamed, so any existing exports/links to that file path keep working.

## Data model
**Materials** (core columns, in order — custom columns the user adds appear
after these): `Category, Material, Brand, Quantity, Unit of Measurement,
Unit Cost (₱), Total Cost (₱), Notes`. `Total Cost (₱)` is always derived as
`Quantity × Unit Cost (₱)` and is not directly editable — this is
intentional, so displayed totals can never drift out of sync with the
line-item math. (There is no Supplier/Store column — it was removed. It's
also in `boq_data.LEGACY_DROPPED_MATERIAL_COLUMNS`, so it's actively
stripped on load from any older project file that still has it, rather than
reappearing as a custom column. Add other retired column names there if a
future schema change needs the same treatment.)

All currency columns (`Unit Cost (₱)`, `Total Cost (₱)`, `Amount (₱)`,
`Default Unit Cost (₱)` in the catalog) display with the `"accounting"`
Streamlit number format, i.e. comma thousands separators (e.g. `108,000.00`)
rather than a plain unformatted number.

**Expenses**: `Expense Name, Amount (₱), Notes` — for logistics, labor,
permits, contingency, etc. Tracked separately from materials and added into
the Grand Total shown at the bottom of the page.

## Materials catalog
The 🗂️ Materials Catalog tab holds a reusable, project-independent list of
default materials — `Category, Material, Brand, Model, Unit of Measurement,
Default Unit Cost (₱)` — stored at `data/materials_catalog.xlsx`. `Brand` is
the manufacturer/supplier name; `Model` is a specific SKU/model code when
one is known (most hardware/consumable items don't have one and leave it
blank — only major equipment like panels/inverters typically do). In the
📋 BOQ tab, "➕ Add items from catalog" lets you multi-select several catalog
entries at once (or use "Select all" / "Clear all" to pick the whole catalog
in one click), set a quantity for each in a preview table, and append them
all to the current project's Materials table in one action (Model, if
present, is carried into each row's Notes as `Model: ...` since the BOQ
table itself has no separate Model column). Each added row's `Unit Cost (₱)`
is then a normal, independently editable cell — changing it only affects
that project, not the catalog default.

`Unit of Measurement` values are normalized to one canonical spelling per
physical unit (e.g. `pc`/`Pc`/`pcs` → `pcs`, `Mtrs`/`mtrs` → `m`) via
`catalog.normalize_unit()` / `catalog.UNIT_ALIASES`, applied on every load
and save, so the same unit never ends up spelled multiple ways in the
catalog. Add more aliases there if a new variant shows up.

The catalog itself is directly editable in its tab (add/edit/remove rows,
including default prices) and persisted with "💾 Save catalog" — useful since
supplier prices change over time. It was seeded from `PACKAGES.xlsx` (a
6.6kWp/6kWac/314AH-battery package quote) via:
```
.venv/bin/python tools/import_catalog_from_workbook.py PACKAGES.xlsx
```
That script is reusable for any other package-quote file with the same
layout (a "No./Description/Technical Specifications/Brand/Model/Quantity/
UoM/Price/Unit Cost/Total Cost" header row, category labels only on the
first row of each group) — run it again against a new file to add more
default materials without duplicating existing catalog entries (matched on
Category + Material + Brand + Unit of Measurement). It deliberately stops
importing a sheet at the first cost-summary row (Total Material Cost,
Installation, Engineering, etc. — see `catalog.STOP_PHRASES`), since those
are labor/markup lines, not materials.

## Price cross-checking
Shopee and Lazada block automated scraping and none of the three platforms
(Shopee, Lazada, Facebook Marketplace) offer a public product-search API, so
this deliberately does **not** attempt live scraping. Instead, the app
generates pre-filled search links (per material row, and via a standalone
"Quick price check" box) that open in a new browser tab for the user to
compare prices manually and update `Unit Cost (₱)` by hand. URL templates
live in `tools/lib/price_links.py` — if a platform's search URL pattern
changes, that's the one file to fix.

Per-row links query on `"{Brand} {Material}"` (not just the material name)
whenever a Brand is set on that row, since brand-qualified searches return
much more relevant results (e.g. "Aiko Solar Panel 660W" instead of just
"Solar Panel 660W"). The standalone "Quick price check" box has no
associated row, so it searches on whatever text is typed into it as-is.

**Known caveat**: the Facebook Marketplace link uses a location-less search
URL that can behave inconsistently depending on the viewer's login/location
state in their browser. If it stops returning useful results, swap in a
location-scoped URL in `price_links.py` (commented inline with an example).

## Data location and durability
Locally (the default, no extra setup required), project files live in
`data/projects/`, and the catalog lives at `data/materials_catalog.xlsx` —
both **durable, not disposable**, unlike `.tmp/`: nothing in `data/` is
regenerated automatically, and it should not be cleaned up casually. Both
are gitignored (Excel binaries don't diff usefully in git, and this may hold
client pricing), so back them up manually (e.g. periodic copy to Drive) if
that matters.

When deployed for shared access, the app switches to a Google Sheets
storage backend instead (local files don't survive a hosted container
restarting) — see `workflows/deploy_boq_app.md` for the full setup. Both
backends are always available in the code; which one is active is decided
automatically based on whether Google credentials are configured, via
`tools/boq_app.py`'s `_select_storage_backend()`.

## Edge cases handled
- **Uploaded file missing required columns** (Material, Brand, Quantity,
  Unit of Measurement, Unit Cost): rejected with a clear message listing
  exactly which columns are missing, rather than silently proceeding with
  bad data.
- **Uploaded file with non-standard headers**: matched via a small alias
  table in `boq_data.COLUMN_ALIASES` (case-insensitive, whitespace-trimmed).
  Add more aliases there if a common variant isn't recognized.
- **Duplicate project names**: the on-disk filename (slug) gets a numeric
  suffix (`-2`, `-3`, …) on collision; the display name can still be reused.
- **Editing while searching**: row add/delete is disabled while a search
  filter is active (edits to existing visible rows still work) — this avoids
  the complexity/risk of reconciling inserted or deleted rows against a
  filtered subset. Clear the search box to add or remove rows.
- **Custom columns**: no separate list is persisted — any column present in
  a project's `Materials` sheet that isn't one of the core columns is
  automatically treated as custom (offered for removal, protected from
  accidental deletion of core columns).

## Things learned / to revisit
- The project's `.venv` runs Python 3.9.6 (no Homebrew/pyenv on this
  machine). Contrary to an initial assumption during planning, current
  Streamlit (1.50.0) installs and runs fine on 3.9 — `Requires-Python:
  >=3.9,!=3.9.7`. No Python upgrade was needed. Don't assume Streamlit
  requires 3.10+ without checking the actual wheel metadata for whatever
  version pip resolves.
- The generic importer (`parse_package_workbook`) can't reliably separate
  Brand from Model from freeform spec text, so the initial `PACKAGES.xlsx`
  import was manually reviewed and re-curated row by row rather than trusted
  as-is. Two items are worth a second look against actual supplier docs: the
  Solar Panel row lists brand "Aiko" with model "TSM-600DE21" (a
  Trina-Solar-style code), and the Inverter row lists brand "Deye" with
  model "SUN2000-60KTL-M1" (a Huawei-style commercial-inverter code) — both
  look like they may be leftover model numbers from a template this package
  quote was built from, rather than the actual Aiko/Deye part numbers.
- Currency columns silently lost their decimal places (e.g. showed `52800`
  instead of `52,800.00`) whenever every value in the column happened to be
  a whole number. Cause: `pd.to_numeric(...).fillna(0.0)` doesn't force
  float64 -- if the source had no NaNs and only whole numbers, pandas keeps
  it int64, and Streamlit's `format="accounting"` displays int64 columns
  without decimals unless a `step` is also set. Fixed by explicitly
  `.astype(float)` on every numeric column in `boq_data.py`/`catalog.py`,
  plus `step=0.01` on every currency `NumberColumn` as a second line of
  defense. Keep both if adding new currency columns.
