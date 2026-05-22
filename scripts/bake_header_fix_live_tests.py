"""
Generate xlsm files for live Excel testing of the attribute-header preservation fix.

Outputs are written to ``tests/live_excel_testing/_header_fix_live/``.
Open each one in Excel and verify the described expected behavior.

Run from the repo root:

    python scripts/bake_header_fix_live_tests.py
"""
from __future__ import annotations

import shutil
from pathlib import Path

from pyopenvba.excel import ExcelFile
from pyopenvba.vba import VBAModuleKind


REPO_ROOT = Path(__file__).resolve().parents[1]
LIVE = REPO_ROOT / "tests" / "live_excel_testing"
OUT_DIR = LIVE / "_header_fix_live"

WORKBOOK_ONLY = LIVE / "workbook_only_module_test.xlsm"
SHEET_ONLY = LIVE / "sheet_only_module_test.xlsm"
TEMPLATE = LIVE / "test_macro_workbook.xlsm"


def _reset() -> None:
    if OUT_DIR.exists():
        shutil.rmtree(OUT_DIR)
    OUT_DIR.mkdir(parents=True, exist_ok=True)


def _copy(src: Path, name: str) -> Path:
    if not src.exists():
        raise SystemExit(f"missing source workbook: {src}")
    dst = OUT_DIR / name
    shutil.copy(src, dst)
    return dst


def case_01_thisworkbook_body_only() -> Path:
    """set_module('ThisWorkbook', body) — header must be auto-preserved.

    Expected in Excel: VBE shows ThisWorkbook with a Hello() Sub that
    pops a MsgBox, and no repair prompt on open.
    """
    out = _copy(WORKBOOK_ONLY, "01_thisworkbook_body_only.xlsm")
    with ExcelFile(out) as wb:
        wb.set_module(
            "ThisWorkbook",
            'Sub Hello()\r\n    MsgBox "ThisWorkbook body-only edit"\r\nEnd Sub\r\n',
        )
        wb.save()
    return out


def case_02_sheet_body_only() -> Path:
    """set_module('Sheet1', body) — header must be auto-preserved.

    Expected in Excel: VBE shows Sheet1 with a SayHi() Sub.
    """
    out = _copy(SHEET_ONLY, "02_sheet_body_only.xlsm")
    with ExcelFile(out) as wb:
        wb.set_module(
            "Sheet1",
            'Sub SayHi()\r\n    MsgBox "Sheet1 body-only edit"\r\nEnd Sub\r\n',
        )
        wb.save()
    return out


def case_03_add_standard_module_body_only() -> Path:
    """add_module('Module1', body, standard) on workbook-only fixture.

    Expected in Excel: VBE shows ThisWorkbook (with original code preserved)
    plus a new standard Module1 containing a Foo() Sub.
    """
    out = _copy(WORKBOOK_ONLY, "03_add_standard_module_body_only.xlsm")
    with ExcelFile(out) as wb:
        wb.vba_project().add_module(
            "Module1",
            'Sub Foo()\r\n    MsgBox "added Module1.Foo"\r\nEnd Sub\r\n',
            kind=VBAModuleKind.standard,
        )
        wb.save()
    return out


def case_04_thisworkbook_plus_added_module() -> Path:
    """Combined: edit ThisWorkbook body AND add a new Module1.

    This is the exact failing workflow the user reported.

    Expected in Excel: both modules show their new code.
    """
    out = _copy(WORKBOOK_ONLY, "04_thisworkbook_edit_plus_added_module.xlsm")
    with ExcelFile(out) as wb:
        wb.set_module(
            "ThisWorkbook",
            'Sub Hello()\r\n    MsgBox "ThisWorkbook updated"\r\nEnd Sub\r\n',
        )
        wb.vba_project().add_module(
            "Module1",
            'Sub Foo()\r\n    MsgBox "Module1 added"\r\nEnd Sub\r\n',
            kind=VBAModuleKind.standard,
        )
        wb.save()
    return out


def case_05_full_source_replacement() -> Path:
    """set_module with explicit full source (Attribute header included).

    Expected in Excel: ThisWorkbook shows a Custom() Sub (no repair prompt).
    """
    out = _copy(WORKBOOK_ONLY, "05_full_source_replacement.xlsm")
    full = (
        'Attribute VB_Name = "ThisWorkbook"\r\n'
        'Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"\r\n'
        'Attribute VB_GlobalNameSpace = False\r\n'
        'Attribute VB_Creatable = False\r\n'
        'Attribute VB_PredeclaredId = True\r\n'
        'Attribute VB_Exposed = True\r\n'
        'Attribute VB_TemplateDerived = False\r\n'
        'Attribute VB_Customizable = True\r\n'
        '\r\n'
        'Sub Custom()\r\n    MsgBox "full source replacement"\r\nEnd Sub\r\n'
    )
    with ExcelFile(out) as wb:
        wb.set_module("ThisWorkbook", full)
        wb.save()
    return out


def case_06_fresh_workbook_full_workflow() -> Path:
    """create_new() then write ThisWorkbook body AND a new Module1.

    Expected in Excel: brand-new workbook opens with no repair prompt,
    ThisWorkbook has Hello() Sub, Module1 has Foo() Sub.
    """
    out = OUT_DIR / "06_create_new_full_workflow.xlsm"
    wb = ExcelFile.create_new(out)
    try:
        wb.set_module(
            "ThisWorkbook",
            'Sub Hello()\r\n    MsgBox "fresh workbook ThisWorkbook"\r\nEnd Sub\r\n',
        )
        # create_new() already ships with an empty Module1 -- edit its body.
        wb.set_module(
            "Module1",
            'Sub Foo()\r\n    MsgBox "fresh workbook Module1"\r\nEnd Sub\r\n',
        )
        wb.save()
    finally:
        wb.close()
    return out


CASES = [
    case_01_thisworkbook_body_only,
    case_02_sheet_body_only,
    case_03_add_standard_module_body_only,
    case_04_thisworkbook_plus_added_module,
    case_05_full_source_replacement,
    case_06_fresh_workbook_full_workflow,
]


def main() -> None:
    _reset()
    print(f"writing live-test workbooks to {OUT_DIR.relative_to(REPO_ROOT)}\n")
    for fn in CASES:
        path = fn()
        # Validate the saved file parses cleanly before handing it off.
        with ExcelFile(path) as wb:
            problems = wb.validate()
        status = "OK" if not problems else f"PROBLEMS: {problems}"
        print(f"  [{status:8}] {path.name}")
        if problems:
            raise SystemExit(f"validation failed for {path.name}: {problems}")
    print("\nopen each .xlsm in Excel and confirm the VBE shows the expected code.")


if __name__ == "__main__":
    main()
