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

The Materials table (both tabs: BOQ Materials, Materials Catalog) is
wrapped in `st.form(..., border=False)` with an "Apply changes" submit
button, rather than a bare `st.data_editor`. Without the form, Streamlit
reruns the *entire script* every time a cell edit is committed (tabbing or
clicking to the next cell) -- filling in one new row across 5-6 columns
meant up to 5-6 full-page reruns before the user finished, which reads as
the page "refreshing" mid-entry. Inside a form, `st.data_editor`'s edits
(including add/delete rows) are batched client-side and only reach Python
when the submit button is clicked -- one rerun for an entire row (or
batch of edits) instead of one per cell. The tradeoff, explicitly accepted
for this app: things that depend on the edited data (Materials
Total/Grand Total metrics, the price-check list, Total Cost recomputation)
only update once "Apply changes" is clicked, not live as you type. The
"➕ Add items from catalog" preview table and the Other Expenses table were
deliberately left un-batched (out of scope when this was raised, and lower
friction anyway -- typically only 1-2 cells being edited at a time).

**Apply changes vs. Save**: "Apply changes" only commits edits into this
browser session (`st.session_state`) -- fast, local, no network call.
"💾 Save" (sidebar, for the active project) and "💾 Save catalog" (Materials
Catalog tab) are the only actions that actually write to storage (Google
Sheets when deployed, local `.xlsx` otherwise), which is comparatively slow
and, when deployed, immediately visible to every other viewer of the app.
Deliberately kept as two separate steps rather than merged, so a user can
apply several rounds of edits and review the result before publishing them
to shared storage. To make the distinction legible rather than just
"two buttons that sound similar," both Save buttons track a dirty flag: a
snapshot of what's actually on disk/Sheets is taken right after every
load and every successful save (`_snapshot_project()` /
`_snapshot_catalog()`), compared against the live session state on every
render (`_project_dirty()` / `_catalog_dirty()` via `DataFrame.equals()`),
and the button switches label/color accordingly -- "🟠 Save*" (`type="primary"`)
when there's anything applied but not yet persisted, "✅ Save"
(`type="secondary"`) when the session matches storage exactly. Comparison
failures (rare; e.g. an unexpected dtype mismatch) intentionally default to
"dirty" rather than falsely claiming everything's saved.

**Sorting vs. in-grid row add/delete**: Streamlit's `st.data_editor` can't
have both native column-header click-to-sort *and* `num_rows="dynamic"`
(the mode that lets you type into a blank trailing row to add one, or
click a row's trash icon to delete it) -- `num_rows="dynamic"` explicitly
disables sorting. Since both Materials tables (BOQ Materials, Materials
Catalog) need to be sortable, both now always run `num_rows="fixed"`, and
row add/delete moved elsewhere:
- **Adding**: BOQ Materials rows come from "➕ Add items from catalog"
  (pick-and-set-quantity) or "➕ Add all materials (lowest price)" (see
  below) -- there's no "type a new row" option, but there never was a
  useful one anyway, since `Material` there has always been a closed
  dropdown sourced from the catalog. Materials Catalog rows come from the
  "➕ Add a new material" expander (a small form: Category/Material/Brand/
  Model/UoM/Store/Unit Cost, "Add material" button) above the table.
- **Deleting**: both tables have a transient `🗑️ Remove?` checkbox column
  (`REMOVE_COL` in `boq_app.py`) prepended to the display copy right before
  rendering, never part of the real schema -- check a row, click "Apply
  changes", and it's dropped from the underlying DataFrame along with
  whatever other edits were made in the same submission. Stripped back out
  immediately after the editor returns, before anything touches
  `st.session_state` or storage, so it can never leak into a saved file.

**Default sort order**: both tables load pre-sorted --
`boq_data.sort_materials_df()` (Category, then Material, ascending) for
BOQ Materials, `catalog.sort_catalog_df()` (Category, then Material, then
Unit Cost, ascending) for the Materials Catalog and everywhere the catalog
feeds a picker (e.g. "Add items from catalog"'s multiselect, so browsing
and picking always agree on order). Both sort case-insensitively via a
throwaway lowercase key column, never mutating the real Category/Material
text. This is a *display* default, re-applied on every load -- a user's
in-browser column-header sort (click to reorder) is a client-side view on
top of it and doesn't change what gets saved; the next reload reverts to
the Category/Material(/Unit Cost) default.

**"Add all materials (lowest price)"** (`catalog.cheapest_catalog_rows()`):
a one-click bulk-add in the BOQ Materials tab's "➕ Add items from catalog"
expander. Adds one row per distinct Material across the *entire* catalog,
each priced at whichever store carries it cheapest (ties keep whichever
row appears first) -- not just the currently store-filtered subset the
multiselect above it is scoped to. Materials whose name already appears
anywhere in the current BOQ (case-insensitive) are skipped, so clicking it
a second time is a no-op ("Nothing to add...") rather than creating
duplicate rows.

**CSV export**: no dedicated button -- both tables' built-in toolbar
download icon (hover over the table, top-right) already exports exactly
what's currently rendered (all columns, and whatever rows the search/store
filter currently shows) as CSV. This is a Streamlit frontend feature, not
something configured from `boq_app.py`.

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
Store Name, Unit Cost (₱), Total Cost (₱), Notes`. `Total Cost (₱)` is
always derived as `Quantity × Unit Cost (₱)` and is not directly editable —
this is intentional, so displayed totals can never drift out of sync with
the line-item math. (There is no Supplier/Store column — it was removed
early on. It's still in `boq_data.LEGACY_DROPPED_MATERIAL_COLUMNS`, so it's
actively stripped on load from any older project file that still has it,
rather than reappearing as a custom column. The newer `Store Name` column,
added for multi-store pricing, is a deliberately distinct concept — tied to
a real Stores master list with addresses and per-store catalog pricing, not
a resurrection of the old free-text column. Add other retired column names
to `LEGACY_DROPPED_MATERIAL_COLUMNS` if a future schema change needs the
same treatment.)

`Category` and `Material` are strict dropdown-only fields
(`st.column_config.SelectboxColumn`), picked from the editable Category/
Material master lists (see below) rather than free text — this matters
because the Store → Unit Cost price lookup (below) needs exact Material-name
matches across different stores' catalog rows to work. `Store Name` is also
a dropdown, sourced from the Stores list plus one extra sentinel option,
`"— Custom / Manual —"` (`catalog.CUSTOM_STORE_LABEL`) — picking that opts a
row out of the auto price-lookup so you can type an arbitrary price by hand.
Legacy/uploaded rows with no Store Name default to this sentinel too, not to
a real store name, since defaulting to e.g. "Plug and Go" would falsely
assert provenance and wrongly make the row eligible for the lookup against a
store it was never actually priced from. `Brand` and `Unit of Measurement`
stay free text (`Unit of Measurement` already has its own consistency
mechanism via `normalize_unit`/`UNIT_ALIASES`).

All currency columns (`Unit Cost (₱)`, `Total Cost (₱)`, `Amount (₱)`, and
the catalog's own `Unit Cost (₱)`) display with the `"accounting"` Streamlit
number format, i.e. comma thousands separators (e.g. `108,000.00`) rather
than a plain unformatted number.

**Expenses**: `Expense Name, Amount (₱), Notes` — for logistics, labor,
permits, contingency, etc. Tracked separately from materials and added into
the Grand Total shown at the bottom of the page.

## Materials catalog (multi-store)
The 🗂️ Materials Catalog tab holds a reusable, project-independent,
**multi-store** list of materials and prices — `Category, Material, Brand,
Model, Unit of Measurement, Store Name, Unit Cost (₱)` — stored at
`data/materials_catalog.xlsx`. It's one flat table, not a separate table per
store: the same Material can appear multiple times, once per store that
carries it, each with its own `Unit Cost (₱)`. A store filter at the top of
the tab (and inside "➕ Add items from catalog") gives it the visual/
functional feel of separate per-store catalogs without the storage
complexity of literally separating them. `Brand` is the manufacturer name;
`Model` is a specific SKU/model code when one is known (most hardware/
consumable items don't have one and leave it blank — only major equipment
like panels/inverters typically do).

**Stores and Categories — two more editable master lists**, under the tab's
"⚙️ Manage stores & categories" expander:
- **Stores** (`catalog.STORES_COLUMNS = ["Store Name", "Address"]`,
  `load_stores`/`save_stores`/`ensure_store_exists`): unique on Store Name
  (case-insensitive). Each store with a non-blank Address gets a
  "📍 Open in Google Maps" link below the editor
  (`price_links.build_maps_link`), opening in a new tab. **"Add a new
  store"** creates a Store Name + Address entry and optionally copies an
  existing store's entire catalog (all rows, including prices) as a
  starting point to edit from — mirrors the existing sidebar "Copy from
  existing project" pattern.
- **Categories** (`catalog.CATEGORY_LIST_COLUMNS = ["Category"]`): a flat
  list, edited via its own small `st.data_editor`, and the source of the
  required `Category` dropdown on the main catalog editor.

**There is no separate Materials master list.** Earlier this existed as a
`Category → Material` mapping edited independently of the catalog, but that
risked drifting out of sync with the catalog's own Category/Material data —
the catalog already *is* the authoritative Category-Material mapping, one
pair per row. So:
- In the **Materials Catalog** tab, `Category` is a required
  `SelectboxColumn` (sourced from the Categories list) and `Material` is
  free text — this is where new materials get defined, alongside the
  Category they belong to.
- In the **BOQ Materials** table, `Material` is a `SelectboxColumn` sourced
  from `catalog.distinct_materials(catalog_df)` (the catalog's own distinct
  Material names, deduped case-insensitively), and `Category` is a
  **disabled/derived column** (`st.column_config.TextColumn(disabled=True)`)
  — the same "derived fields aren't independently editable" principle
  `Total Cost (₱)` already uses.
- **`boq_data.apply_material_categories(df, catalog_df)`** is what makes
  that derivation happen: on every "Apply changes" in the BOQ Materials
  table (same insertion point as `apply_store_pricing`, right before
  `recompute_material_totals`), each row's `Category` gets overwritten from
  `catalog.material_category_map(catalog_df)` (first catalog row for that
  Material wins). A Material with no catalog row keeps its existing
  Category unchanged rather than getting blanked out (same non-destructive
  philosophy as the pricing lookup).
- **`catalog.sync_category_list_from_catalog(catalog_df, category_list_df)`**:
  given a catalog (or newly parsed/imported rows), adds any Category not
  already in the Category list — never removes or overwrites existing
  entries. Shared by `tools/import_catalog_from_workbook.py` (so a freshly
  imported category is immediately selectable) and the migration script's
  initial seeding.
- Stores/Categories dedupe case-insensitively (keeping the first
  occurrence's casing) via `catalog.normalize_stores_df` /
  `catalog.normalize_simple_list_df` — "Panels" and "panels" never both
  survive as separate entries.

**Storage layout differs by backend, deliberately**: locally, Stores/
Categories each live in their own small file (`data/stores.xlsx`,
`data/category_list.xlsx`), *not* extra sheets inside `materials_catalog.xlsx`
— `catalog.save_catalog()`'s `df.to_excel(path, sheet_name="Catalog")` call
replaces the **entire workbook file**, not just that one sheet, so
colocating them would mean every catalog save silently wipes any sheets
sharing that file. On Google Sheets this risk doesn't exist (`gspread` only
touches the one tab it's told to), so there they're just two more tabs
(`Stores`, `Categories`) on the same catalog Spreadsheet — no new secret/ID
needed.

In the 📋 BOQ tab, "➕ Add items from catalog" lets you filter by store
(defaults to "Plug and Go" if present, to preserve the original
single-store experience), multi-select several catalog entries at once (or
"Select all" / "Clear all"), set a quantity for each in a preview table, and
append them all to the current project's Materials table in one action —
each row's `Store Name` and `Unit Cost (₱)` come straight from the picked
catalog row (Model, if present, is carried into the row's Notes as
`Model: ...`, since the BOQ table itself has no separate Model column).
Changing the store filter resets the picker (`catalog_pick` session key),
since previously-picked indices may no longer be valid options in the new
filtered set.

**`boq_data.apply_store_pricing(materials_df, catalog_df)`** is what makes
picking a Store on a BOQ row actually update the price: on every "Apply
changes" click (both the search-filtered and unfiltered Materials editor
branches, right before `recompute_material_totals`), every row whose Store
Name is a real store (not the Custom/Manual sentinel) gets its
`Unit Cost (₱)` looked up by `(Material, Store Name)` against the catalog
and overwritten. **No match found → the existing Unit Cost is left
untouched and a warning is shown** — never reset to 0, since that would
silently destroy a real, previously-entered price the moment a store name
doesn't have a matching catalog row.

`Unit of Measurement` values are normalized to one canonical spelling per
physical unit (e.g. `pc`/`Pc`/`pcs` → `pcs`, `Mtrs`/`mtrs` → `m`) via
`catalog.normalize_unit()` / `catalog.UNIT_ALIASES`, applied on every load
and save, so the same unit never ends up spelled multiple ways in the
catalog. Add more aliases there if a new variant shows up.

The catalog itself (plus Stores/Categories) is directly editable in its tab
and persisted together with one "💾 Save catalog" button (uses a bundled
dirty-state snapshot covering all three pieces of session state —
`_snapshot_catalog_bundle`/`_catalog_bundle_dirty` — same 🟠/✅ pattern as
the project Save button). It was originally seeded from `PACKAGES.xlsx` (a
6.6kWp/6kWac/314AH-battery package quote from "Plug and Go") via:
```
.venv/bin/python tools/import_catalog_from_workbook.py PACKAGES.xlsx
```
or, for a different store's price list:
```
.venv/bin/python tools/import_catalog_from_workbook.py OtherQuote.xlsx "ABC Hardware"
```
That script is reusable for any package-quote file with the same layout (a
"No./Description/Technical Specifications/Brand/Model/Quantity/UoM/Price/
Unit Cost/Total Cost" header row, category labels only on the first row of
each group), attributing every parsed row to the given store name (defaults
to "Plug and Go") — run it again to add more materials/stores without
duplicating existing catalog entries (matched on Category + Material +
Brand + Unit of Measurement + **Store Name**). It deliberately stops
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
`data/projects/`, and the catalog + its master lists live at
`data/materials_catalog.xlsx`, `data/stores.xlsx`, `data/category_list.xlsx`
— all **durable, not disposable**, unlike `.tmp/`:
nothing in `data/` is regenerated automatically, and it should not be
cleaned up casually. All are gitignored (Excel binaries don't diff usefully
in git, and this may hold client pricing), so back them up manually (e.g.
periodic copy to Drive) if that matters.

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
- The sidebar's "Download blank BOQ template" button re-evaluates its
  `data=` argument on *every* script rerun (Streamlit needs the bytes ready
  in case the button gets clicked), not only when actually clicked. That
  means anything it reads from `st.session_state` -- here, the live
  Materials Catalog -- can be in a transient, partially-edited state (e.g.
  a just-added, still-blank catalog row) far more often than you'd expect
  from "only runs on download." Code driven by widgets like this needs to
  tolerate mid-edit data, not just fully-saved data. This crashed
  `_add_catalog_dropdowns()` (fixed by dropping nulls via `.dropna()`
  before stringifying, instead of `.astype(str)` on the raw column, plus
  wrapping the whole thing in a try/except so malformed catalog data
  degrades to "no dropdown suggestions" rather than crashing the app).

  That first fix only patched the one call site that had already crashed.
  A second, still-blank catalog row later crashed a *different* reader
  (`_format_catalog_choice`'s `f"{cost:,.2f}"`, where `cost` was `None`) --
  same root cause, different symptom, because `st.session_state.catalog_df`
  itself was never guaranteed clean, only whatever individual code happened
  to defend against it. Patched that call site too, but the real fix was
  upstream: `catalog_df` is now run through `catalog.normalize_catalog_df()`
  immediately after every `st.data_editor()` call, before it's ever stored
  back to `st.session_state` -- the same pattern `materials_df` already used
  (`recompute_material_totals()` right after its own editor). One
  still-blank row now becomes `""`/`0.0` at the source instead of `None`/
  `NaN` propagating to every future reader to individually defend against.
  Lesson: when a widget's raw output feeds `st.session_state`, normalize
  once at that single write point rather than patching each read site as
  it's discovered crashing -- especially for anything editable via
  `num_rows="dynamic"`, since a freshly added, not-yet-filled-in row is a
  completely normal, frequent state to land in `st.session_state` in.
- Verified before building the multi-store catalog: `df.to_excel(path,
  sheet_name="Catalog")` -- the exact call `catalog.save_catalog()` already
  used -- replaces the **entire local `.xlsx` file**, not just the named
  sheet. This ruled out colocating the new Stores/Categories/Material-List
  reference data as extra sheets inside `materials_catalog.xlsx` (every
  catalog save would have silently wiped them) in favor of three separate
  small files. Google Sheets has no equivalent risk (`gspread` only touches
  the one worksheet tab it's told to), so that backend just gets three more
  tabs on the same catalog Spreadsheet -- a deliberate asymmetry between the
  two backends, not an oversight.
- `st.testing.v1.AppTest` (Streamlit's own headless test framework) can
  drive real widget interactions -- button clicks, selectbox/radio/
  multiselect changes, text inputs, and reading/writing `st.session_state`
  directly -- across full script reruns, without a browser. Used it to
  verify the "Add a new store" + "copy from existing store" flow, dirty-state
  button styling, Google Maps link rendering, and case-insensitive duplicate
  validation, all end-to-end against the real migrated data. **Limitation**:
  it has no accessor for `st.data_editor` -- individual cell edits inside a
  data table can't be driven this way, so anything that specifically
  requires changing a data_editor cell's value (e.g. actually picking a
  different Store on a Materials row and confirming the price updates via
  `apply_store_pricing`) still needs a real click-through, same as this
  project's established pattern for other data_editor-driven features.
- The multi-store migration script (`tools/migrate_multi_store_catalog.py`)
  reads raw (`pd.read_excel` / `worksheet.get_all_values()`, checking for the
  literal old `"Default Unit Cost (₱)"` header) rather than through
  `catalog.load_catalog()`/`gsheets_storage.load_catalog()` -- once the code
  expects the new `"Unit Cost (₱)"` name, those functions would read
  pre-migration data's old-named column as entirely missing and silently
  default every price to 0.0. Migrate data first, deploy code second, for
  any future rename of a column real data already exists under.
- **Google Sheets API rate limit crashed the deployed app right after the
  multi-store rollout** (`AttributeError` inside `_init_state()` calling
  `gsheets_storage.load_stores()`) -- most likely a 429 (the API caps reads
  at 60/minute *per underlying account*, and this app has exactly one
  shared account backing *every* viewer, per the OAuth-not-service-account
  design in `workflows/deploy_boq_app.md`) surfacing as an unhandled
  `AttributeError` somewhere in gspread's response parsing rather than a
  clean `APIError`, worsened by heavy migration/verification API traffic
  run moments before. `_init_state()` was doing 4 separate Sheets reads
  (catalog, stores, categories, material list) on every fresh session, and
  `list_projects()` -- a Drive listing plus one read per project -- was
  running *unconditionally on every single rerun*, not just once per
  session, multiplying the problem further. Fixed by wrapping both in
  `st.cache_data(ttl=60)`: `_cached_load_catalog_bundle()` and
  `_cached_list_projects()`. Since `st.cache_data`'s cache is shared across
  *all* sessions (unlike `st.session_state`, which is per-session), this is
  exactly the right tool here -- the catalog/project list genuinely are
  global, shared data, so N concurrent viewers now cost a handful of API
  calls total instead of 4-per-session and 1-per-rerun-per-project. Cleared
  explicitly (`.clear()`) right after any action that changes what these
  return (save catalog, create project, rename project) so that session
  still sees its own change immediately rather than waiting out the TTL --
  only *other, new* sessions see up to ~60s of staleness, which is an
  acceptable tradeoff for a small-team internal tool.
- Some buttons persist to real storage **immediately**, not just into
  session state -- "Rename" in particular (a deliberate design choice, see
  the "Renaming a project" section above). Driving one of those via
  `AppTest.button(...).click()` during testing actually renamed the real
  local project file to the test's throwaway value; caught it by re-reading
  the file afterward and reverted by hand. Lesson: before clicking a button
  in an automated test, check whether its handler calls a `*_store.save_*`/
  `*_store.create_*` function directly (persists for real) versus only
  mutating `st.session_state` (safe, session-local) -- this project's own
  Apply-vs-Save split (see above) is precisely the distinction to check for.
- The Category<->Material mapping was first added as a standalone
  `Category → Material` master list, backfilled from the catalog's
  existing Category-Material pairs (the catalog was already the ground
  truth for all 50 items, so no manual re-categorization was needed). But
  maintaining it as a *separate* list turned out to be the wrong call --
  it could drift out of sync with the catalog itself, which already
  encodes the exact same mapping, one pair per row. Simplified shortly
  after to derive the mapping directly from the catalog
  (`catalog.material_category_map`/`catalog.distinct_materials`) and
  removed the separate list entirely. Worth remembering generally: when a
  new "master list" would just restate data another table already has,
  prefer deriving it on read over maintaining a second copy that can go
  stale -- and don't hesitate to walk back a recently-added structure once
  that becomes clear, rather than layering more sync logic on top of it.
- **Root cause of the deployed app's repeated crashes, finally confirmed**:
  widening `_diagnostics` to wrap the Apply-changes blocks (see above)
  surfaced the real, previously-redacted exception -- `gspread.exceptions.
  APIError: [500]: Internal error encountered` from a plain
  `load_catalog()` call, i.e. Google's Sheets API itself returning a
  transient server-side error on an otherwise valid, well-formed request.
  This is common and expected under normal load per Google's own API
  guidance (retry with exponential backoff), not a sign of a bug in this
  app's code -- and it's the likely explanation for the earlier, differently
  -shaped `AttributeError` crashes too: gspread/its HTTP layer can surface
  a transient server error as an oddly-typed exception depending on exactly
  where in a chained call (`open_by_key(...).worksheet(...)`) it lands,
  rather than a clean `APIError` every time. Fixed by adding a `_retry()`
  helper in `gsheets_storage.py` (exponential backoff, up to 4 attempts,
  only for `{429, 500, 502, 503, 504}`; anything else -- e.g. a real 404 --
  raises immediately) and routing every Sheets API call in that module
  through it, including the ones inside `_worksheet_to_df`/`_write_df`/
  `_meta_from_worksheet`/`_write_meta` that take an already-opened
  worksheet object. Verified with a mocked `APIError` (transient → retries
  then succeeds; non-transient → raises on first attempt; persistent
  transient → exhausts retries and raises) since this can't be triggered
  on demand against the real API. Lesson: when a Streamlit Cloud crash
  can't be reproduced locally and looks structurally weird (wrong
  exception type for the apparent cause), suspect the underlying API
  itself being flaky before assuming an app-logic bug -- surfacing the
  real traceback (not guessing) is what made this diagnosable at all.
