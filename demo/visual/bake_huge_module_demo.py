"""Bake a stress-test demo .xlsm containing a deliberately huge VBA module.

Generates realistic-looking VBA source (thousands of small helper functions
plus a Run sub that exercises them) sized to roughly 100 KB, then writes
it as ``demo/visual/05_huge_module.xlsm``.  The point is to extend the
empirical envelope for ExcelFile.create_new() + multi-chunk MS-OVBA
compression well beyond the existing 16 KB Excel-anchored fixture.

Run from the repo root:

    python scripts/bake_huge_module_demo.py
"""

from __future__ import annotations

from pathlib import Path

from pyopenvba import ExcelFile

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "demo" / "visual" / "05_huge_module.xlsm"

TARGET_BYTES = 100_000  # aim for ~100 KB of module source
HEADER = 'Attribute VB_Name = "Module1"\r\n'

PROLOGUE = """\
Option Explicit

' Stress-test module: thousands of generated helper functions plus a
' Run sub that uses a handful of them to paint a colourful grid.  The
' point of this workbook is to prove that pyOpenVBA's MS-OVBA compressor
' can produce a large multi-chunk module that Excel will load cleanly.

Public Sub Run()
    Const SIZE As Long = 24
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Cells.Clear
    ws.Cells.ColumnWidth = 3
    ws.Cells.RowHeight = 18

    Dim r As Long, c As Long, idx As Long
    For r = 1 To SIZE
        For c = 1 To SIZE
            idx = ((r - 1) * SIZE + (c - 1)) Mod HelperCount()
            ws.Cells(r, c).Interior.Color = HelperValue(idx)
            DoEvents
        Next c
    Next r
End Sub

Public Function HelperCount() As Long
    HelperCount = {count}
End Function

Public Function HelperValue(ByVal idx As Long) As Long
    Select Case idx
"""

EPILOGUE = """\
        Case Else
            HelperValue = 0
    End Select
End Function
"""


def _build_source() -> tuple[str, int]:
    """Return (source, helper_count) sized at roughly TARGET_BYTES."""
    parts: list[str] = []
    count = 0
    while True:
        # Each helper function: real VBA, all distinct names and bodies.
        idx = count
        seed_r = (idx * 73 + 11) & 0xFF
        seed_g = (idx * 151 + 47) & 0xFF
        seed_b = (idx * 233 + 89) & 0xFF
        rgb = (seed_b << 16) | (seed_g << 8) | seed_r
        fn = (
            f"Public Function H{idx:05d}() As Long\r\n"
            f"    H{idx:05d} = &H{rgb:06X}\r\n"
            f"End Function\r\n"
        )
        case_line = f"        Case {idx}: HelperValue = H{idx:05d}()\r\n"
        parts.append(fn)
        parts.append(case_line)
        count += 1
        # Cheap running-size check based on parts; stop once we exceed target.
        if sum(len(p) for p in parts) >= TARGET_BYTES:
            break

    body = "".join(parts)
    prologue = PROLOGUE.format(count=count)
    return HEADER + prologue + body + EPILOGUE, count


def main() -> None:
    source, count = _build_source()
    print(
        f"generated module: {len(source)} bytes, {count} helper functions, "
        f"{source.count(chr(10))} lines"
    )

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with ExcelFile.create_new(OUT) as wb:
        project = wb.vba_project()
        module = project.get_module("Module1")
        module.source = source
        module.dirty = True
        wb.save()
    print(f"baked {OUT.relative_to(ROOT)} ({OUT.stat().st_size} bytes on disk)")

    # Round-trip verification.
    with ExcelFile(OUT) as wb:
        proj = wb.vba_project()
        rt = proj.get_module("Module1").source
    assert rt == source, (
        f"round-trip mismatch: wrote {len(source)} bytes, read back {len(rt)}"
    )
    print(f"round-trip OK: {len(rt)} bytes identical after reopen")


if __name__ == "__main__":
    main()
