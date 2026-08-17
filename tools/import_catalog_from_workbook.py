"""Import materials + default prices from a package BOQ workbook into the
persistent materials catalog (data/materials_catalog.xlsx).

Reusable for any future package file that follows the same layout as
PACKAGES.xlsx (e.g. a different system size/config quoted the same way).
Existing catalog entries are left untouched; only new (Category, Material,
Brand, Unit of Measurement) combinations are added.

Usage:
    .venv/bin/python tools/import_catalog_from_workbook.py PACKAGES.xlsx
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from lib import catalog


def main() -> None:
    if len(sys.argv) != 2:
        print("Usage: python tools/import_catalog_from_workbook.py <path-to-workbook.xlsx>")
        sys.exit(1)

    source_path = Path(sys.argv[1])
    if not source_path.exists():
        print(f"File not found: {source_path}")
        sys.exit(1)

    new_items = catalog.parse_package_workbook(source_path)
    if new_items.empty:
        print(f"No material rows recognized in {source_path}. Nothing imported.")
        return

    merged, added_count = catalog.merge_into_catalog(new_items)
    catalog.save_catalog(merged)

    print(f"Parsed {len(new_items)} material row(s) from {source_path.name}.")
    print(f"Added {added_count} new catalog entr{'y' if added_count == 1 else 'ies'}.")
    print(f"Skipped {len(new_items) - added_count} already in the catalog.")
    print(f"Catalog now has {len(merged)} total entries -> {catalog.CATALOG_PATH}")


if __name__ == "__main__":
    main()
