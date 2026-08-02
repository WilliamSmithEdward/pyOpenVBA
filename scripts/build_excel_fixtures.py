"""Build Excel fixture workbooks from fixture source directories.

For each sub-directory under tests/fixtures/excel/:
  - *.bas files → standard VBA modules
      (Module1.bas updates the default Module1 via set_module;
       all others are added as new standard modules)
  - *.cls files → class VBA modules (added via add_module)

Built workbooks are written to build/fixtures/excel/<fixture_name>.xlsm.

Run from the repo root:

    python scripts/build_excel_fixtures.py
"""

from __future__ import annotations

import sys
from pathlib import Path

from pyopenvba import ExcelFile  # noqa: E402
from pyopenvba.vba import VBAModuleKind  # noqa: E402

FIXTURES_DIR = Path(__file__).resolve().parents[1] / "tests" / "fixtures" / "excel"
BUILD_DIR = Path(__file__).resolve().parents[1] / "build" / "fixtures" / "excel"

# Modules that exist by default in a new workbook (no add_module needed).
_DEFAULT_MODULES = {"ThisWorkbook", "Sheet1", "Module1"}


def _build_fixture(fixture_dir: Path) -> Path:
    name = fixture_dir.name
    output = BUILD_DIR / f"{name}.xlsm"
    BUILD_DIR.mkdir(parents=True, exist_ok=True)

    with ExcelFile.create_new(output) as wb:
        proj = wb.vba_project()

        for bas in sorted(fixture_dir.glob("*.bas")):
            module_name = bas.stem
            source = bas.read_text(encoding="utf-8")
            if module_name in _DEFAULT_MODULES:
                wb.set_module(module_name, source)
            else:
                proj.add_module(module_name, source, kind=VBAModuleKind.standard)

        for cls_file in sorted(fixture_dir.glob("*.cls")):
            module_name = cls_file.stem
            source = cls_file.read_text(encoding="utf-8")
            proj.add_module(module_name, source, kind=VBAModuleKind.other)

        wb.save()

    return output


def main() -> None:
    fixture_dirs = sorted(d for d in FIXTURES_DIR.iterdir() if d.is_dir())
    if not fixture_dirs:
        print(f"No fixture directories found under {FIXTURES_DIR}")
        sys.exit(1)

    for fixture_dir in fixture_dirs:
        output = _build_fixture(fixture_dir)
        with ExcelFile(output) as wb:
            modules = wb.module_names()
        print(f"Built: {output.relative_to(BUILD_DIR.parents[2])}  modules={modules}")


if __name__ == "__main__":
    main()
