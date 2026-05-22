"""Generate a battery of mutated workbooks for manual Excel verification.

Run from the repo root:

    python scripts/build_excel_verification_set.py

Output goes to ``tests/live_excel_testing/_excel_verify/`` (gitignored
by convention; safe to delete).  Each file is a copy of an in-tree
fixture with a specific mutation applied via :func:`ExcelFile.save`,
covering the highest-risk save paths:

    01_noop_xlsm.xlsm                  -- pure no-op (cache verbatim)
    02_source_edit_xlsm.xlsm           -- Module1 edited; cache invalidated
    03_add_module_xlsm.xlsm            -- new standard module added
    04_rename_module_xlsm.xlsm         -- Module1 renamed to Module1Renamed
    05_delete_module_xlsm.xlsm         -- Module1 deleted
    06_noop_xlsb.xlsb                  -- xlsb pure no-op
    07_source_edit_xlsb.xlsb           -- xlsb Module1 edited
    08_protected_noop.xlsm             -- protected: no-op (no opt-in needed)
    09_protected_source_edit.xlsm      -- protected: edit w/ allow_protected=True

For each file, open in Excel and verify:
    * Excel does NOT show a "repair" dialog.
    * Alt+F11 opens the VBE.
    * Each listed module is present with the expected source.
    * For the protected workbook, the VBA password ("test") still unlocks
      the project.
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO / "src"))

from pyopenvba.excel import ExcelFile  # noqa: E402

FIX = REPO / "tests" / "live_excel_testing"
OUT = FIX / "_excel_verify"

XLSM = FIX / "test_macro_workbook.xlsm"
XLSB = FIX / "test_macro_workbook.xlsb"
PROT = FIX / "workbook_with_password_protected_vba_modules.xlsm"

EDIT_MARKER = "\r\n' pyOpenVBA round-trip marker -- safe to delete\r\n"


def _copy(src: Path, dest: Path) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dest)
    return dest


def _noop(src: Path, dest_name: str) -> Path:
    out = _copy(src, OUT / dest_name)
    with ExcelFile(out) as wb:
        wb.save()
    return out


def _edit_module(src: Path, module_name: str, dest_name: str, **save_kw: object) -> Path:
    out = _copy(src, OUT / dest_name)
    with ExcelFile(out) as wb:
        current = wb.get_module(module_name)
        wb.set_module(module_name, current + EDIT_MARKER)
        wb.save(**save_kw)  # type: ignore[arg-type]
    return out


def _edit_module1(src: Path, dest_name: str, **save_kw: object) -> Path:
    return _edit_module(src, "Module1", dest_name, **save_kw)


def _add_module(src: Path, dest_name: str) -> Path:
    out = _copy(src, OUT / dest_name)
    with ExcelFile(out) as wb:
        wb.vba_project().add_module(
            "ModuleAdded",
            "Attribute VB_Name = \"ModuleAdded\"\r\n"
            "Public Sub Hello()\r\n"
            "    MsgBox \"Added by pyOpenVBA\"\r\n"
            "End Sub\r\n",
        )
        wb.save()
    return out


def _rename_module(src: Path, dest_name: str) -> Path:
    out = _copy(src, OUT / dest_name)
    with ExcelFile(out) as wb:
        wb.vba_project().rename_module("Module1", "Module1Renamed")
        wb.save()
    return out


def _delete_module(src: Path, dest_name: str) -> Path:
    out = _copy(src, OUT / dest_name)
    with ExcelFile(out) as wb:
        wb.vba_project().delete_module("Module1")
        wb.save()
    return out


def main() -> int:
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)

    produced: list[Path] = []
    produced.append(_noop(XLSM, "01_noop_xlsm.xlsm"))
    produced.append(_edit_module1(XLSM, "02_source_edit_xlsm.xlsm"))
    produced.append(_add_module(XLSM, "03_add_module_xlsm.xlsm"))
    produced.append(_rename_module(XLSM, "04_rename_module_xlsm.xlsm"))
    produced.append(_delete_module(XLSM, "05_delete_module_xlsm.xlsm"))
    produced.append(_noop(XLSB, "06_noop_xlsb.xlsb"))
    produced.append(_edit_module1(XLSB, "07_source_edit_xlsb.xlsb"))
    produced.append(_noop(PROT, "08_protected_noop.xlsm"))
    produced.append(
        _edit_module(
            PROT, "PasswordTest", "09_protected_source_edit.xlsm", allow_protected=True
        )
    )

    print(f"Wrote {len(produced)} verification workbooks to:")
    print(f"  {OUT}")
    for p in produced:
        print(f"  - {p.name}  ({p.stat().st_size:,} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
