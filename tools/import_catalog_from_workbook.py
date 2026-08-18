"""Import materials + prices from a package BOQ workbook into the
persistent materials catalog (data/materials_catalog.xlsx), attributed to a
specific store.

Reusable for any future package file that follows the same layout as
PACKAGES.xlsx (e.g. a different system size/config quoted the same way, or
a quote from a different store entirely). Existing catalog entries are left
untouched; only new (Category, Material, Brand, Unit of Measurement, Store
Name) combinations are added.

Usage:
    .venv/bin/python tools/import_catalog_from_workbook.py PACKAGES.xlsx
    .venv/bin/python tools/import_catalog_from_workbook.py OtherStoreQuote.xlsx "ABC Hardware"
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from lib import catalog


def main() -> None:
    if len(sys.argv) not in (2, 3):
        print('Usage: python tools/import_catalog_from_workbook.py <path-to-workbook.xlsx> ["Store Name"]')
        sys.exit(1)

    source_path = Path(sys.argv[1])
    store_name = sys.argv[2] if len(sys.argv) > 2 else "Plug and Go"
    if not source_path.exists():
        print(f"File not found: {source_path}")
        sys.exit(1)

    new_items = catalog.parse_package_workbook(source_path, store_name=store_name)
    if new_items.empty:
        print(f"No material rows recognized in {source_path}. Nothing imported.")
        return

    catalog.ensure_store_exists(store_name)
    merged, added_count = catalog.merge_into_catalog(new_items)
    catalog.save_catalog(merged)

    # So every newly imported Category/Material is immediately selectable
    # in the strict dropdowns elsewhere in the app, not just present in
    # the flat catalog -- existing list entries are never touched, only
    # additions.
    category_list = catalog.load_category_list()
    material_list = catalog.load_material_list()
    category_list, material_list = catalog.sync_reference_lists_from_catalog(merged, category_list, material_list)
    catalog.save_category_list(category_list)
    catalog.save_material_list(material_list)

    print(f'Parsed {len(new_items)} material row(s) from {source_path.name} for store "{store_name}".')
    print(f"Added {added_count} new catalog entr{'y' if added_count == 1 else 'ies'}.")
    print(f"Skipped {len(new_items) - added_count} already in the catalog.")
    print(f"Catalog now has {len(merged)} total entries -> {catalog.CATALOG_PATH}")


if __name__ == "__main__":
    main()
