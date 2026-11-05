"""A part with a DOCTYPE, which Excel refuses, but for its theme, which it goes without. The prediction was 1004 throughout.

    python scripts/measure_bad_xml_doctype.py

writes tests/fixtures/bad_xml/doctype.json.
"""
from __future__ import annotations

import json
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = """Function Tried(ByVal file As String) As String
    Dim wb As Workbook
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wb = Workbooks.Open(file, UpdateLinks:=0)
    Tried = Err.Number & "|" & Err.Description
    Err.Clear
    wb.Close False
End Function
"""
PARTS = ("xl/worksheets/sheet1.xml", "xl/workbook.xml", "xl/styles.xml", "docProps/core.xml", "xl/theme/theme1.xml", "_rels/.rels")
DOCTYPES = {"plain": b"<!DOCTYPE root>", "entity": b'<!DOCTYPE root [<!ENTITY e "x">]>'}


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    rows: dict[str, str] = {}
    with zipfile.ZipFile(ROOT / "tests/fixtures/bad_xml/base.xlsx") as given:
        parts = {name: given.read(name) for name in given.namelist()}
    with tempfile.TemporaryDirectory(prefix="pyopenvba_dtd_") as folder, ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
        excel.new_document()
        for part in PARTS:
            for label, doctype in DOCTYPES.items():
                file = Path(folder) / "d.xlsx"
                data = parts[part]
                at = data.find(b"?>") + 2 if data.lstrip().startswith(b"<?xml") else 0
                with zipfile.ZipFile(file, "w", zipfile.ZIP_DEFLATED) as out:
                    for name, held in parts.items():
                        out.writestr(name, data[:at] + doctype + data[at:] if name == part else held)
                read = excel.run_vba(SRC, "Tried", (str(file),), timeout=60.0, line_numbers=False)
                assert read.ok, (part, label, read.outcome)
                rows[f"{part}|{label}"] = str(read.value).split("|")[0]
                print(part, label, rows[f"{part}|{label}"], flush=True)
    (ROOT / "tests/fixtures/bad_xml/doctype.json").write_text(json.dumps(rows, indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
