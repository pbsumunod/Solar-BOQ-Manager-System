"""BOQ / inventory manager for solar installation projects.

Run with: streamlit run tools/boq_app.py
"""

import contextlib
import sys
import traceback
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import pandas as pd
import streamlit as st

from lib import boq_data, catalog, gsheets_storage, price_links

st.set_page_config(page_title="BOQ Manager", layout="wide", page_icon="☀️")

EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _safe_secret(*keys):
    """Dig into st.secrets, returning None instead of raising if secrets
    aren't configured at all (e.g. running locally with no secrets.toml) or
    the requested path isn't present."""
    try:
        value = st.secrets
        for key in keys:
            value = value[key]
        return value
    except Exception:
        return None


def _check_password() -> bool:
    """Gate the whole app behind a shared password when one is configured
    (via the "app_password" secret). No secret configured -- e.g. local
    dev -- means no gate at all, so this never locks you out of your own
    machine."""
    required_password = _safe_secret("app_password")
    if not required_password:
        return True
    if st.session_state.get("_authenticated"):
        return True

    st.title("☀️ BOQ Manager")
    st.text_input("Password", type="password", key="_password_attempt")
    if st.button("Enter", type="primary"):
        if st.session_state.get("_password_attempt") == required_password:
            st.session_state["_authenticated"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


def _select_storage_backend():
    """Pick Google Sheets (if configured via secrets) or local Excel files
    otherwise. Both expose the same list_projects/load_project/save_project/
    create_project/load_stores/save_stores/load_category_list/... functions,
    so the rest of the app doesn't need to know which one is active."""
    google_oauth = _safe_secret("google_oauth")
    boq_config = _safe_secret("boq")
    if google_oauth and boq_config and not gsheets_storage.is_configured():
        gsheets_storage.configure(
            dict(google_oauth),
            boq_config["projects_folder_id"],
            boq_config["catalog_spreadsheet_id"],
        )
    if gsheets_storage.is_configured():
        return gsheets_storage, gsheets_storage
    return boq_data, catalog


def _snapshot_project() -> None:
    """Record what's currently on disk/Sheets, so _project_dirty() can tell
    whether there are edits sitting only in this session that Save hasn't
    persisted yet."""
    st.session_state.project_saved_snapshot = (
        st.session_state.materials_df.copy(),
        st.session_state.expenses_df.copy(),
    )


def _project_dirty() -> bool:
    snapshot = st.session_state.get("project_saved_snapshot")
    if snapshot is None:
        return False
    try:
        return not (
            st.session_state.materials_df.equals(snapshot[0])
            and st.session_state.expenses_df.equals(snapshot[1])
        )
    except Exception:
        # If comparison itself is ever unclear, err toward "unsaved" rather
        # than falsely reassuring the user everything's persisted.
        return True


def _snapshot_catalog_bundle() -> None:
    """The Materials Catalog tab now covers three pieces of session state
    (the catalog itself, plus Stores/Categories master data) under one
    "Save catalog" button -- snapshot all three together so the dirty flag
    reflects any of them changing."""
    st.session_state.catalog_bundle_snapshot = (
        st.session_state.catalog_df.copy(),
        st.session_state.stores_df.copy(),
        st.session_state.category_list_df.copy(),
    )


def _catalog_bundle_dirty() -> bool:
    snapshot = st.session_state.get("catalog_bundle_snapshot")
    if snapshot is None:
        return False
    try:
        return not (
            st.session_state.catalog_df.equals(snapshot[0])
            and st.session_state.stores_df.equals(snapshot[1])
            and st.session_state.category_list_df.equals(snapshot[2])
        )
    except Exception:
        return True


@contextlib.contextmanager
def _diagnostics(label: str):
    """Streamlit Cloud redacts the real exception message/traceback from
    any *uncaught* error before showing it to the user ("This app has
    encountered an error... redacted to prevent data leaks"), which makes
    genuinely diagnosing a production-only failure (one that doesn't
    reproduce locally, e.g. because of a Python-version or dependency
    difference between here and the deploy target) close to impossible
    from the outside. Catching it ourselves and displaying it via
    st.error()/st.code() bypasses that redaction -- those are our own
    explicit output, not an uncaught crash, so Streamlit's redaction
    doesn't touch them -- then st.stop() halts cleanly instead of
    crashing further up into Streamlit's generic error screen.

    Use as `with _diagnostics("doing X"): ...` around any block, not just
    a single call -- e.g. the multi-step "apply this table's edits" logic,
    where the exact failing line matters for diagnosis."""
    try:
        yield
    except Exception:
        st.error(f"Something went wrong while {label}. Full details below -- please copy/paste this if reporting it:")
        st.code(traceback.format_exc())
        st.stop()


def _load_with_diagnostics(label: str, fn):
    with _diagnostics(f"loading {label}"):
        return fn()


@st.cache_data(ttl=60, show_spinner=False)
def _cached_load_catalog_bundle():
    """The Materials Catalog (plus Stores/Categories) is genuinely shared,
    global data -- cache it across sessions, not just within one, so N
    concurrent viewers loading a fresh session cost 3 Sheets API calls
    total, not 3*N. Google's Sheets API caps reads at 60 per minute *per
    underlying account*, and every viewer of this app shares the same one
    (see workflows/manage_boq_inventory.md) -- without this, a handful of
    people opening the app around the same time can exhaust that quota and
    crash everyone's session. A short 60s TTL means "Save catalog"
    elsewhere in the app is still reflected for new sessions within about
    a minute, without needing explicit cache invalidation."""
    return (
        _load_with_diagnostics("the Materials Catalog", catalog_store.load_catalog),
        _load_with_diagnostics("the Stores list", catalog_store.load_stores),
        _load_with_diagnostics("the Category list", catalog_store.load_category_list),
    )


@st.cache_data(ttl=60, show_spinner=False)
def _cached_list_projects():
    """list_projects() runs unconditionally at the top of the sidebar on
    *every* rerun (not just once per session, unlike the catalog bundle
    above) so newly created projects show up without a full page reload.
    For the Sheets backend that's a Drive folder listing plus one Sheets
    read per project just to populate the sidebar dropdown -- caching it
    means rapid-fire reruns (e.g. clicking "Apply changes" repeatedly while
    editing) don't each cost a fresh round of API calls. Cleared explicitly
    right after create_project() so a just-created project appears
    immediately rather than waiting out the TTL."""
    return _load_with_diagnostics("the project list", project_store.list_projects)


def _init_state() -> None:
    if "current_project_slug" not in st.session_state:
        st.session_state.current_project_slug = None
        st.session_state.current_project_name = None
        st.session_state.materials_df = boq_data.blank_materials_df()
        st.session_state.expenses_df = boq_data.blank_expenses_df()
        st.session_state.project_meta = {}
        _snapshot_project()
    if "catalog_df" not in st.session_state:
        (
            st.session_state.catalog_df,
            st.session_state.stores_df,
            st.session_state.category_list_df,
        ) = _cached_load_catalog_bundle()
        _snapshot_catalog_bundle()


def _load_project(slug: str) -> None:
    materials_df, expenses_df, meta = project_store.load_project(slug)
    st.session_state.current_project_slug = slug
    st.session_state.current_project_name = meta.get("project_name") or slug
    st.session_state.materials_df = materials_df
    st.session_state.expenses_df = expenses_df
    st.session_state.project_meta = meta
    _snapshot_project()


def _text_search_mask(df: pd.DataFrame, query: str, columns: list) -> pd.Series:
    if not query.strip():
        return pd.Series(True, index=df.index)
    q = query.strip().lower()
    haystack = pd.Series("", index=df.index)
    for col in columns:
        haystack = haystack + " " + df[col].astype(str).str.lower()
    return haystack.str.contains(q, regex=False, na=False)


if not _check_password():
    st.stop()

project_store, catalog_store = _select_storage_backend()
_init_state()

# ---------------------------------------------------------------------------
# Sidebar: project management, grouped by how often each action is used --
# switching/saving the active project first, setup/reference actions last.
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("☀️ BOQ Manager")

    projects = _cached_list_projects()
    project_options = {p["slug"]: p["name"] for p in projects}

    st.subheader("📁 Active project")
    if project_options:
        slugs = list(project_options.keys())
        current = st.session_state.current_project_slug
        index = slugs.index(current) if current in slugs else 0
        selected_slug = st.selectbox(
            "Project",
            options=slugs,
            format_func=lambda s: project_options[s],
            index=index,
            key="project_selector",
            label_visibility="collapsed",
        )
        if selected_slug != st.session_state.current_project_slug:
            _load_project(selected_slug)
            st.rerun()

        project_dirty = _project_dirty()
        col_save, col_export = st.columns(2)
        with col_save:
            save_label = "🟠 Save*" if project_dirty else "✅ Save"
            if st.button(
                save_label,
                type="primary" if project_dirty else "secondary",
                use_container_width=True,
                help="Unsaved changes in this session" if project_dirty else "Everything is saved",
            ):
                project_store.save_project(
                    st.session_state.current_project_slug,
                    st.session_state.materials_df,
                    st.session_state.expenses_df,
                    st.session_state.project_meta,
                )
                _snapshot_project()
                st.toast("Project saved.", icon="💾")
                st.rerun()
        with col_export:
            st.download_button(
                "⬇️ Export",
                data=boq_data.export_bytes(
                    st.session_state.materials_df,
                    st.session_state.expenses_df,
                    st.session_state.project_meta,
                ),
                file_name=f"{st.session_state.current_project_slug}.xlsx",
                mime=EXCEL_MIME,
                use_container_width=True,
            )

        with st.expander("✏️ Rename project"):
            renamed = st.text_input(
                "Project name", value=st.session_state.current_project_name, key="rename_project_input"
            )
            if st.button("Rename"):
                if not renamed.strip():
                    st.error("Enter a project name.")
                else:
                    st.session_state.current_project_name = renamed.strip()
                    st.session_state.project_meta["project_name"] = renamed.strip()
                    project_store.save_project(
                        st.session_state.current_project_slug,
                        st.session_state.materials_df,
                        st.session_state.expenses_df,
                        st.session_state.project_meta,
                    )
                    _cached_list_projects.clear()
                    _snapshot_project()
                    st.toast("Renamed.", icon="✏️")
                    st.rerun()
    else:
        st.caption("No projects yet — create one below.")

    st.divider()
    st.subheader("➕ New project")
    with st.expander("Create a project", expanded=not project_options):
        new_name = st.text_input("Project name", key="new_project_name")
        source_choices = ["Blank", "Upload existing BOQ"]
        if project_options:
            source_choices.append("Copy from existing project")
        source = st.radio("Start from", source_choices, key="new_project_source")

        uploaded_file = None
        copy_source_slug = None
        if source == "Upload existing BOQ":
            uploaded_file = st.file_uploader("BOQ Excel file (.xlsx)", type=["xlsx"], key="new_project_upload")
        elif source == "Copy from existing project":
            copy_source_slug = st.selectbox(
                "Copy materials and expenses from",
                options=list(project_options.keys()),
                format_func=lambda s: project_options[s],
                key="copy_source_project",
            )

        if st.button("Create project", type="primary"):
            if not new_name.strip():
                st.error("Enter a project name.")
            elif source == "Upload existing BOQ" and uploaded_file is None:
                st.error("Upload a .xlsx file, or switch to 'Blank'.")
            else:
                try:
                    materials_df = expenses_df = None
                    if uploaded_file is not None:
                        materials_df, expenses_df, warnings = boq_data.read_uploaded_boq(uploaded_file)
                        for w in warnings:
                            st.warning(w)
                    elif source == "Copy from existing project" and copy_source_slug:
                        materials_df, expenses_df, _ = project_store.load_project(copy_source_slug)
                    slug = project_store.create_project(new_name, materials_df, expenses_df)
                    _cached_list_projects.clear()
                    _load_project(slug)
                    st.toast(f'Created project "{new_name}".', icon="✅")
                    st.rerun()
                except ValueError as e:
                    st.error(str(e))

    st.divider()
    st.download_button(
        "⬇️ Download blank BOQ template",
        data=boq_data.generate_template_workbook(
            st.session_state.category_list_df, st.session_state.catalog_df
        ),
        file_name="boq_template.xlsx",
        mime=EXCEL_MIME,
        use_container_width=True,
        help="Category and Material columns include dropdown suggestions from the Materials Catalog.",
    )

# ---------------------------------------------------------------------------
# Main area
# ---------------------------------------------------------------------------
tab_boq, tab_catalog = st.tabs(["📋 BOQ", "🗂️ Materials Catalog"])

with tab_boq:
    if not st.session_state.current_project_slug:
        st.title("☀️ BOQ Manager")
        st.info("👈 Create or select a project from the sidebar to get started.")
    else:
        st.title(f"BOQ — {st.session_state.current_project_name}")

        # Reserved near the top so the bottom line is always visible without
        # scrolling past the tables below -- filled in further down, once
        # this run's edits to Materials/Expenses have been captured.
        summary_slot = st.container()
        st.divider()

        # ---- Materials ----
        with st.container(border=True):
            st.subheader("📦 Materials")

            materials_df = st.session_state.materials_df
            catalog_df = st.session_state.catalog_df
            stores_df = st.session_state.stores_df
            store_names = stores_df["Store Name"].tolist()

            if catalog_df.empty:
                st.caption("Materials Catalog is empty — add default materials in the 🗂️ Materials Catalog tab.")
            else:
                with st.expander("➕ Add items from catalog"):
                    add_store_options = ["All stores"] + store_names
                    default_store_index = (
                        add_store_options.index("Plug and Go") if "Plug and Go" in add_store_options else 0
                    )
                    add_store_filter = st.selectbox(
                        "Filter by store", add_store_options, index=default_store_index, key="add_from_catalog_store_filter"
                    )

                    # Reset the picker when the store filter changes, since
                    # its previously-selected indices may no longer be valid
                    # options in the new filtered set.
                    if st.session_state.get("_last_add_store_filter") != add_store_filter:
                        st.session_state["catalog_pick"] = []
                        st.session_state["_last_add_store_filter"] = add_store_filter

                    filtered_catalog_for_add = (
                        catalog_df if add_store_filter == "All stores" else catalog_df[catalog_df["Store Name"] == add_store_filter]
                    )

                    def _format_catalog_choice(idx: int) -> str:
                        row = catalog_df.loc[idx]
                        brand_model = " (" + " / ".join(filter(None, [row["Brand"], row["Model"]])) + ")" if row["Brand"] or row["Model"] else ""
                        unit = row["Unit of Measurement"] or "unit"
                        store = row["Store Name"] or "—"
                        return f"[{store}] {row['Category']} — {row['Material']}{brand_model} · ₱{row['Unit Cost (₱)']:,.2f}/{unit}"

                    col_select_all, col_clear_all = st.columns(2)
                    with col_select_all:
                        if st.button("Select all", key="catalog_select_all", use_container_width=True):
                            st.session_state["catalog_pick"] = list(filtered_catalog_for_add.index)
                    with col_clear_all:
                        if st.button("Clear all", key="catalog_clear_all", use_container_width=True):
                            st.session_state["catalog_pick"] = []

                    picked_indices = st.multiselect(
                        "Catalog items",
                        options=list(filtered_catalog_for_add.index),
                        format_func=_format_catalog_choice,
                        key="catalog_pick",
                    )

                    if picked_indices:
                        preview_cols = ["Category", "Material", "Brand", "Model", "Unit of Measurement", "Store Name", "Unit Cost (₱)"]
                        preview_df = catalog_df.loc[picked_indices, preview_cols].copy()
                        preview_df["Quantity"] = 1.0

                        st.caption("Set quantities, then add all of them to the BOQ at once.")
                        edited_preview = st.data_editor(
                            preview_df,
                            column_config={
                                "Category": st.column_config.TextColumn(disabled=True),
                                "Material": st.column_config.TextColumn(disabled=True),
                                "Brand": st.column_config.TextColumn(disabled=True),
                                "Model": st.column_config.TextColumn(disabled=True),
                                "Unit of Measurement": st.column_config.TextColumn(disabled=True),
                                "Store Name": st.column_config.TextColumn(disabled=True),
                                "Unit Cost (₱)": st.column_config.NumberColumn(disabled=True, format="accounting", step=0.01),
                                "Quantity": st.column_config.NumberColumn(min_value=0.0, step=1.0),
                            },
                            num_rows="fixed",
                            hide_index=True,
                            key="catalog_multi_preview",
                        )

                        if st.button(f"Add {len(picked_indices)} item(s) to BOQ", type="primary"):
                            new_rows = []
                            for idx, preview_row in edited_preview.iterrows():
                                picked = catalog_df.loc[idx]
                                new_row = {col: "" for col in materials_df.columns}
                                new_row.update(
                                    {
                                        "Category": picked["Category"],
                                        "Material": picked["Material"],
                                        "Brand": picked["Brand"],
                                        "Quantity": preview_row["Quantity"],
                                        "Unit of Measurement": picked["Unit of Measurement"],
                                        "Store Name": picked["Store Name"],
                                        "Unit Cost (₱)": picked["Unit Cost (₱)"],
                                        "Notes": f"Model: {picked['Model']}" if picked["Model"] else "",
                                    }
                                )
                                new_rows.append(new_row)
                            new_rows_df = pd.DataFrame(new_rows)
                            materials_df = new_rows_df if materials_df.empty else pd.concat(
                                [materials_df, new_rows_df], ignore_index=True
                            )
                            st.session_state.materials_df = boq_data.recompute_material_totals(materials_df)
                            st.rerun()

            col_search, col_columns = st.columns([3, 1])
            with col_search:
                search_query = st.text_input(
                    "Search materials",
                    key="search_query",
                    placeholder="🔍 Search materials by name or brand…",
                    label_visibility="collapsed",
                )

            materials_df = st.session_state.materials_df

            with col_columns:
                with st.popover("⚙️ Columns", use_container_width=True):
                    new_col_name = st.text_input("New column name", key="new_col_name")
                    if st.button("Add column"):
                        name = new_col_name.strip()
                        if not name:
                            st.error("Enter a column name.")
                        elif name in materials_df.columns:
                            st.error("A column with that name already exists.")
                        else:
                            materials_df[name] = ""
                            st.session_state.materials_df = materials_df
                            st.rerun()

                    custom_cols = [c for c in materials_df.columns if c not in boq_data.CORE_MATERIAL_COLUMNS]
                    if custom_cols:
                        col_to_remove = st.selectbox(
                            "Remove a custom column", options=["— select —"] + custom_cols, key="col_to_remove"
                        )
                        if st.button("Remove column") and col_to_remove != "— select —":
                            st.session_state.materials_df = materials_df.drop(columns=[col_to_remove])
                            st.rerun()
                    else:
                        st.caption("No custom columns yet.")

            column_config = {
                # Category is derived from Material (apply_material_categories,
                # below) via the Category<->Material mapping, the same
                # "pick X, Y follows" pattern as Store -> Unit Cost -- so it's
                # disabled here rather than an independently pickable field,
                # which would let it drift out of sync with what the
                # material actually is.
                "Category": st.column_config.TextColumn("Category", disabled=True),
                "Material": st.column_config.SelectboxColumn(
                    "Material", options=catalog.distinct_materials(catalog_df), required=True
                ),
                "Quantity": st.column_config.NumberColumn("Quantity", min_value=0.0),
                "Store Name": st.column_config.SelectboxColumn(
                    "Store Name", options=store_names + [catalog.CUSTOM_STORE_LABEL], required=True
                ),
                "Unit Cost (₱)": st.column_config.NumberColumn("Unit Cost (₱)", min_value=0.0, format="accounting", step=0.01),
                "Total Cost (₱)": st.column_config.NumberColumn("Total Cost (₱)", disabled=True, format="accounting", step=0.01),
            }

            materials_search_cols = ["Material", "Brand"]
            if search_query.strip():
                filtered = materials_df[_text_search_mask(materials_df, search_query, materials_search_cols)]
                st.caption(f"Showing {len(filtered)} of {len(materials_df)} items. Clear search to add or delete rows.")
                with st.form("materials_form_filtered", border=False):
                    edited = st.data_editor(
                        filtered,
                        num_rows="fixed",
                        column_config=column_config,
                        key="materials_editor_filtered",
                    )
                    st.form_submit_button("Apply changes", type="primary")
                with _diagnostics("applying your Materials changes"):
                    materials_df.loc[edited.index] = edited
                    materials_df = boq_data.apply_material_categories(materials_df, catalog_df)
                    materials_df, pricing_warnings = boq_data.apply_store_pricing(materials_df, catalog_df)
                    materials_df = boq_data.recompute_material_totals(materials_df)
            else:
                with st.form("materials_form", border=False):
                    edited = st.data_editor(
                        materials_df,
                        num_rows="dynamic",
                        column_config=column_config,
                        key="materials_editor",
                    )
                    st.form_submit_button("Apply changes", type="primary")
                with _diagnostics("applying your Materials changes"):
                    edited = boq_data.apply_material_categories(edited, catalog_df)
                    materials_df, pricing_warnings = boq_data.apply_store_pricing(edited, catalog_df)
                    materials_df = boq_data.recompute_material_totals(materials_df)

            for w in pricing_warnings:
                st.warning(w)

            st.session_state.materials_df = materials_df
            st.caption("Edit cells or add/remove rows above, then click \"Apply changes\" — totals and the price-check list below update after you apply. The sidebar's \"Save\" button turns 🟠 whenever there's anything applied but not yet written to storage.")

            with st.expander("🔎 Price-check links (Shopee / Lazada / Facebook Marketplace)"):
                visible_materials = materials_df[_text_search_mask(materials_df, search_query, materials_search_cols)]
                if visible_materials.empty:
                    st.caption("No materials to check yet.")
                else:
                    for _, row in visible_materials.iterrows():
                        material_name = str(row.get("Material", "")).strip()
                        if not material_name:
                            continue
                        brand_name = str(row.get("Brand", "")).strip()
                        query_text = f"{brand_name} {material_name}".strip() if brand_name else material_name
                        links = price_links.build_price_check_links(query_text)
                        link_md = " &nbsp;|&nbsp; ".join(f"[{site}]({url})" for site, url in links.items())
                        st.markdown(f"**{query_text}** — {link_md}")

        # ---- Other expenses ----
        with st.container(border=True):
            st.subheader("🧾 Other Expenses")
            st.caption("Logistics, labor, permits, and anything else outside the materials list.")
            expenses_df = st.data_editor(
                st.session_state.expenses_df,
                num_rows="dynamic",
                column_config={
                    "Amount (₱)": st.column_config.NumberColumn("Amount (₱)", min_value=0.0, format="accounting", step=0.01),
                },
                key="expenses_editor",
            )
            st.session_state.expenses_df = boq_data.validate_and_normalize_expenses(expenses_df)

        # ---- Quick, ad hoc price check ----
        with st.container(border=True):
            st.subheader("🔎 Quick price check")
            quick_query = st.text_input("Search term", key="quick_price_query", placeholder="e.g. 6mm² DC cable")
            if quick_query.strip():
                links = price_links.build_price_check_links(quick_query)
                cols = st.columns(len(links))
                for col, (site, url) in zip(cols, links.items()):
                    col.link_button(f"Search {site}", url, use_container_width=True)

        # ---- Fill in the totals summary reserved near the top ----
        materials_total = boq_data.materials_total(st.session_state.materials_df)
        expenses_total = boq_data.expenses_total(st.session_state.expenses_df)
        grand_total = materials_total + expenses_total

        with summary_slot:
            c1, c2, c3 = st.columns(3)
            c1.metric("Materials Total", f"₱{materials_total:,.2f}")
            c2.metric("Other Expenses Total", f"₱{expenses_total:,.2f}")
            c3.metric("Grand Total", f"₱{grand_total:,.2f}")

with tab_catalog:
    st.title("🗂️ Materials Catalog")
    st.caption(
        "Multi-store default materials and prices you can pull into any BOQ. Edits here update the "
        "defaults for future use — they don't change unit costs already saved on existing projects."
    )

    stores_df = st.session_state.stores_df
    store_names = stores_df["Store Name"].tolist()

    col_store_filter, col_search, col_save = st.columns([1, 2, 1])
    with col_store_filter:
        catalog_store_filter = st.selectbox(
            "Filter by store",
            ["All stores"] + store_names,
            key="catalog_store_filter",
            label_visibility="collapsed",
        )
    with col_search:
        catalog_search = st.text_input(
            "Search catalog",
            key="catalog_search",
            placeholder="🔍 Search by category, material, brand, model…",
            label_visibility="collapsed",
        )
    with col_save:
        catalog_dirty = _catalog_bundle_dirty()
        save_catalog_label = "🟠 Save catalog*" if catalog_dirty else "✅ Save catalog"
        if st.button(
            save_catalog_label,
            type="primary" if catalog_dirty else "secondary",
            use_container_width=True,
            help="Unsaved changes in this session" if catalog_dirty else "Everything is saved",
        ):
            catalog_store.save_catalog(st.session_state.catalog_df)
            catalog_store.save_stores(st.session_state.stores_df)
            catalog_store.save_category_list(st.session_state.category_list_df)
            _cached_load_catalog_bundle.clear()  # so new sessions see this save immediately, not after the TTL
            _snapshot_catalog_bundle()
            st.toast(f"Catalog saved ({len(st.session_state.catalog_df)} items).", icon="💾")
            st.rerun()

    with st.expander("⚙️ Manage stores & categories"):
        st.markdown("**Stores**")
        edited_stores = st.data_editor(
            stores_df,
            num_rows="dynamic",
            column_config={"Address": st.column_config.TextColumn("Address", width="large")},
            key="stores_editor",
        )
        st.session_state.stores_df = catalog.normalize_stores_df(edited_stores)

        stores_with_address = st.session_state.stores_df[st.session_state.stores_df["Address"].str.strip() != ""]
        if not stores_with_address.empty:
            for _, row in stores_with_address.iterrows():
                maps_url = price_links.build_maps_link(row["Address"])
                st.markdown(f"📍 **{row['Store Name']}** — {row['Address']} · [Open in Google Maps]({maps_url})")

        st.markdown("---")
        st.markdown("**Add a new store**")
        current_store_names = st.session_state.stores_df["Store Name"].tolist()
        new_store_name = st.text_input("Store name", key="new_store_name")
        new_store_address = st.text_input("Address", key="new_store_address")
        copy_choice_options = ["Blank"] + (["Copy from existing store"] if current_store_names else [])
        copy_choice = st.radio("Starting catalog", copy_choice_options, key="new_store_copy_choice")
        copy_source = None
        if copy_choice == "Copy from existing store":
            copy_source = st.selectbox("Copy catalog from", current_store_names, key="new_store_copy_source")

        if st.button("Create store", type="primary"):
            name = new_store_name.strip()
            if not name:
                st.error("Enter a store name.")
            elif name.lower() in {s.lower() for s in current_store_names}:
                st.error(f'A store named "{name}" already exists.')
            else:
                updated_stores = pd.concat(
                    [st.session_state.stores_df, pd.DataFrame([{"Store Name": name, "Address": new_store_address.strip()}])],
                    ignore_index=True,
                )
                st.session_state.stores_df = catalog.normalize_stores_df(updated_stores)

                if copy_source:
                    copied_rows = st.session_state.catalog_df[st.session_state.catalog_df["Store Name"] == copy_source].copy()
                    copied_rows["Store Name"] = name
                    updated_catalog = pd.concat([st.session_state.catalog_df, copied_rows], ignore_index=True)
                    st.session_state.catalog_df = catalog.normalize_catalog_df(updated_catalog)

                st.toast(f'Created store "{name}".', icon="✅")
                st.rerun()

        st.markdown("---")
        st.markdown("**Categories**")
        edited_cat_list = st.data_editor(
            st.session_state.category_list_df, num_rows="dynamic", hide_index=True, key="category_list_editor"
        )
        st.session_state.category_list_df = catalog.normalize_simple_list_df(edited_cat_list, "Category")

    catalog_df = st.session_state.catalog_df
    catalog_search_cols = ["Category", "Material", "Brand", "Model", "Store Name"]

    catalog_column_config = {
        # This IS where a Material's Category gets defined -- the BOQ
        # Materials table derives its own Category from whatever's set
        # here (apply_material_categories, keyed off Material), so this
        # is a required dropdown rather than a derived/disabled field.
        "Category": st.column_config.SelectboxColumn(
            "Category", options=st.session_state.category_list_df["Category"].tolist(), required=True
        ),
        "Store Name": st.column_config.SelectboxColumn("Store Name", options=store_names, required=True),
        "Unit Cost (₱)": st.column_config.NumberColumn(
            "Unit Cost (₱)", min_value=0.0, format="accounting", step=0.01
        ),
    }

    text_mask = _text_search_mask(catalog_df, catalog_search, catalog_search_cols)
    store_mask = (
        pd.Series(True, index=catalog_df.index)
        if catalog_store_filter == "All stores"
        else (catalog_df["Store Name"] == catalog_store_filter)
    )
    is_filtered = bool(catalog_search.strip()) or catalog_store_filter != "All stores"

    if is_filtered:
        visible_catalog = catalog_df[text_mask & store_mask]
        st.caption(f"Showing {len(visible_catalog)} of {len(catalog_df)} items. Clear filters to add or delete rows.")
        with st.form("catalog_form_filtered", border=False):
            edited_catalog = st.data_editor(
                visible_catalog,
                num_rows="fixed",
                column_config=catalog_column_config,
                key="catalog_editor_filtered",
            )
            st.form_submit_button("Apply changes", type="primary")
        with _diagnostics("applying your catalog changes"):
            catalog_df.loc[edited_catalog.index] = edited_catalog
    else:
        with st.form("catalog_form", border=False):
            catalog_df = st.data_editor(
                catalog_df,
                num_rows="dynamic",
                column_config=catalog_column_config,
                key="catalog_editor",
            )
            st.form_submit_button("Apply changes", type="primary")
    st.caption("Edit cells or add/remove rows above, then click \"Apply changes\" to update this session. \"Save catalog\" (top right) turns 🟠 whenever there's anything applied but not yet written to storage.")

    # Normalize right after editing (same pattern as materials_df going
    # through recompute_material_totals after its own editor) so a
    # still-blank row the user just added -- None/NaN price, empty text --
    # never leaks into session_state as raw, untyped data for every other
    # reader of the catalog (template dropdowns, the "Add items from
    # catalog" picker, etc.) to potentially trip over.
    st.session_state.catalog_df = catalog.normalize_catalog_df(catalog_df)
