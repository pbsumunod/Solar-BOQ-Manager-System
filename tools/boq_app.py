"""BOQ / inventory manager for solar installation projects.

Run with: streamlit run tools/boq_app.py
"""

import sys
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
    create_project (and load_catalog/save_catalog) functions, so the rest of
    the app doesn't need to know which one is active."""
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


def _init_state() -> None:
    if "current_project_slug" not in st.session_state:
        st.session_state.current_project_slug = None
        st.session_state.current_project_name = None
        st.session_state.materials_df = boq_data.blank_materials_df()
        st.session_state.expenses_df = boq_data.blank_expenses_df()
        st.session_state.project_meta = {}
    if "catalog_df" not in st.session_state:
        st.session_state.catalog_df = catalog_store.load_catalog()


def _load_project(slug: str) -> None:
    materials_df, expenses_df, meta = project_store.load_project(slug)
    st.session_state.current_project_slug = slug
    st.session_state.current_project_name = meta.get("project_name") or slug
    st.session_state.materials_df = materials_df
    st.session_state.expenses_df = expenses_df
    st.session_state.project_meta = meta


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

    projects = project_store.list_projects()
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

        col_save, col_export = st.columns(2)
        with col_save:
            if st.button("💾 Save", type="primary", use_container_width=True):
                project_store.save_project(
                    st.session_state.current_project_slug,
                    st.session_state.materials_df,
                    st.session_state.expenses_df,
                    st.session_state.project_meta,
                )
                st.toast("Project saved.", icon="💾")
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
                    _load_project(slug)
                    st.toast(f'Created project "{new_name}".', icon="✅")
                    st.rerun()
                except ValueError as e:
                    st.error(str(e))

    st.divider()
    st.download_button(
        "⬇️ Download blank BOQ template",
        data=boq_data.generate_template_workbook(st.session_state.catalog_df),
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

            if catalog_df.empty:
                st.caption("Materials Catalog is empty — add default materials in the 🗂️ Materials Catalog tab.")
            else:
                with st.expander("➕ Add items from catalog"):
                    def _format_catalog_choice(idx: int) -> str:
                        row = catalog_df.loc[idx]
                        brand_model = " (" + " / ".join(filter(None, [row["Brand"], row["Model"]])) + ")" if row["Brand"] or row["Model"] else ""
                        unit = row["Unit of Measurement"] or "unit"
                        return f"{row['Category']} — {row['Material']}{brand_model} · ₱{row['Default Unit Cost (₱)']:,.2f}/{unit}"

                    col_select_all, col_clear_all = st.columns(2)
                    with col_select_all:
                        if st.button("Select all", key="catalog_select_all", use_container_width=True):
                            st.session_state["catalog_pick"] = list(catalog_df.index)
                    with col_clear_all:
                        if st.button("Clear all", key="catalog_clear_all", use_container_width=True):
                            st.session_state["catalog_pick"] = []

                    picked_indices = st.multiselect(
                        "Catalog items",
                        options=list(catalog_df.index),
                        format_func=_format_catalog_choice,
                        key="catalog_pick",
                    )

                    if picked_indices:
                        preview_cols = ["Category", "Material", "Brand", "Model", "Unit of Measurement", "Default Unit Cost (₱)"]
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
                                "Default Unit Cost (₱)": st.column_config.NumberColumn(disabled=True, format="accounting", step=0.01),
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
                                        "Unit Cost (₱)": picked["Default Unit Cost (₱)"],
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
                "Quantity": st.column_config.NumberColumn("Quantity", min_value=0.0),
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
                materials_df.loc[edited.index] = edited
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
                materials_df = boq_data.recompute_material_totals(edited)

            st.session_state.materials_df = materials_df
            st.caption("Edit cells or add/remove rows above, then click \"Apply changes\" — totals and the price-check list below update after you apply.")

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
        "Default materials and prices you can pull into any BOQ. Edits here update the "
        "defaults for future use — they don't change unit costs already saved on existing projects."
    )

    col_search, col_save = st.columns([3, 1])
    with col_search:
        catalog_search = st.text_input(
            "Search catalog",
            key="catalog_search",
            placeholder="🔍 Search by category, material, brand, model…",
            label_visibility="collapsed",
        )
    with col_save:
        if st.button("💾 Save catalog", type="primary", use_container_width=True):
            catalog_store.save_catalog(st.session_state.catalog_df)
            st.toast(f"Catalog saved ({len(st.session_state.catalog_df)} items).", icon="💾")

    catalog_df = st.session_state.catalog_df
    catalog_search_cols = ["Category", "Material", "Brand", "Model"]

    catalog_column_config = {
        "Default Unit Cost (₱)": st.column_config.NumberColumn(
            "Default Unit Cost (₱)", min_value=0.0, format="accounting", step=0.01
        ),
    }

    if catalog_search.strip():
        visible_catalog = catalog_df[_text_search_mask(catalog_df, catalog_search, catalog_search_cols)]
        st.caption(f"Showing {len(visible_catalog)} of {len(catalog_df)} items. Clear search to add or delete rows.")
        with st.form("catalog_form_filtered", border=False):
            edited_catalog = st.data_editor(
                visible_catalog,
                num_rows="fixed",
                column_config=catalog_column_config,
                key="catalog_editor_filtered",
            )
            st.form_submit_button("Apply changes", type="primary")
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
    st.caption("Edit cells or add/remove rows above, then click \"Apply changes\" to save them into this session (still separate from \"💾 Save catalog\", which writes to storage).")

    # Normalize right after editing (same pattern as materials_df going
    # through recompute_material_totals after its own editor) so a
    # still-blank row the user just added -- None/NaN price, empty text --
    # never leaks into session_state as raw, untyped data for every other
    # reader of the catalog (template dropdowns, the "Add items from
    # catalog" picker, etc.) to potentially trip over.
    st.session_state.catalog_df = catalog.normalize_catalog_df(catalog_df)
