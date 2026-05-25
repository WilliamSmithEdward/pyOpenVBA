"""
Bake a set of Access .accdb files exercising the Phase 5f
multiline-source-replace API. Open each output file in MS Access and
verify the items listed in PRINTED_INSTRUCTIONS at the bottom.

Each output starts from a known good corpus sample so the only
differences from the source are produced by pyOpenVBA writes.
"""

from __future__ import annotations

import shutil
from pathlib import Path

import pyopenvba
from pyopenvba.access import AccessFile

REPO = Path(__file__).resolve().parents[1]
CORPUS = REPO / "tests" / "live_access_test" / "re_corpus" / "samples"
OUT = REPO / "demo" / "output" / "access_phase5f"

SAMPLE_STD = CORPUS / "040__sub_msgbox_hello.accdb"     # std module "M"
SAMPLE_CLASS = CORPUS / "020__empty_ClassModule_C.accdb"  # class module "C"


def _copy(src: Path, name: str) -> Path:
    dst = OUT / name
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(src, dst)
    return dst


def bake_01_simple_multiline_replace() -> Path:
    """Std module 'M': replace body with a multiline Sub."""
    dst = _copy(SAMPLE_STD, "01_simple_multiline_replace.accdb")
    new_src = (
        "Option Explicit\r\n"
        "\r\n"
        "Sub Hello()\r\n"
        '    MsgBox "Hi from Python!"\r\n'
        '    MsgBox "This is line 2"\r\n'
        '    MsgBox "And line 3"\r\n'
        "End Sub\r\n"
    )
    db = AccessFile(dst)
    db.set_module("M", new_src)
    db.save()
    return dst


def bake_02_class_module_body_swap() -> Path:
    """Class module 'C': replace body, attribute preamble must survive."""
    dst = _copy(SAMPLE_CLASS, "02_class_module_body_swap.accdb")
    new_body = (
        "Option Explicit\r\n"
        "\r\n"
        "Private mGreeting As String\r\n"
        "\r\n"
        "Public Sub Greet(ByVal who As String)\r\n"
        '    mGreeting = "Hello, " & who & "!"\r\n'
        "    MsgBox mGreeting\r\n"
        "End Sub\r\n"
        "\r\n"
        "Public Property Get LastGreeting() As String\r\n"
        "    LastGreeting = mGreeting\r\n"
        "End Property\r\n"
    )
    db = AccessFile(dst)
    db.set_module("C", new_body)
    db.save()
    return dst


def bake_03_full_source_with_own_header() -> Path:
    """Class module 'C': caller supplies a full source with its own
    Attribute preamble (flipping PredeclaredId / Exposed to True)."""
    dst = _copy(SAMPLE_CLASS, "03_full_source_with_own_header.accdb")
    full_src = (
        'Attribute VB_Name = "C"\r\n'
        'Attribute VB_GlobalNameSpace = False\r\n'
        'Attribute VB_Creatable = False\r\n'
        'Attribute VB_PredeclaredId = True\r\n'
        'Attribute VB_Exposed = True\r\n'
        "\r\n"
        "Option Explicit\r\n"
        "\r\n"
        "Public Sub FromCaller()\r\n"
        '    MsgBox "Custom attribute header used."\r\n'
        "End Sub\r\n"
    )
    db = AccessFile(dst)
    db.set_module("C", full_src)
    db.save()
    return dst


def bake_04_push_modules_bulk() -> Path:
    """Top-level pyopenvba.push_access(): write a .bas to a folder
    and bulk-apply it to the database."""
    dst = _copy(SAMPLE_STD, "04_push_access_bulk.accdb")
    src_dir = OUT / "04_push_access_src"
    src_dir.mkdir(parents=True, exist_ok=True)
    (src_dir / "M.bas").write_bytes(
        (
            "Option Explicit\r\n"
            "\r\n"
            "Sub Pushed()\r\n"
            '    MsgBox "pushed via pyopenvba.push_access()"\r\n'
            "End Sub\r\n"
            "\r\n"
            "Sub Pushed2()\r\n"
            "    Dim i As Long\r\n"
            "    For i = 1 To 3\r\n"
            '        Debug.Print "i=" & i\r\n'
            "    Next i\r\n"
            "End Sub\r\n"
        ).encode("utf-8")
    )
    pyopenvba.push_access(src_dir, dst)
    return dst


def bake_05_rename_then_replace() -> Path:
    """Rename module 'M' to 'Renamed_M', then replace its body."""
    dst = _copy(SAMPLE_STD, "05_rename_then_replace.accdb")
    db = AccessFile(dst)
    db.rename_module("M", "Renamed_M")
    db.set_module(
        "Renamed_M",
        (
            "Sub StillWorks()\r\n"
            '    MsgBox "Renamed_M survived a body swap."\r\n'
            "End Sub\r\n"
        ),
    )
    db.save()
    return dst


def main() -> None:
    if not SAMPLE_STD.exists() or not SAMPLE_CLASS.exists():
        raise SystemExit(
            f"RE corpus samples missing under {CORPUS}; regenerate via "
            "tests/live_access_test/_corpus_generate.ps1"
        )
    OUT.mkdir(parents=True, exist_ok=True)
    bakers = [
        bake_01_simple_multiline_replace,
        bake_02_class_module_body_swap,
        bake_03_full_source_with_own_header,
        bake_04_push_modules_bulk,
        bake_05_rename_then_replace,
    ]
    produced: list[tuple[Path, str]] = []
    for fn in bakers:
        path = fn()
        # Re-open and read back what pyopenvba sees, for the
        # printed verification report.
        db = AccessFile(path)
        names = db.vba_module_names()
        produced.append((path, ", ".join(names)))
    print("\nBaked files (Python-side read-back module names):\n")
    for path, names in produced:
        print(f"  {path.relative_to(REPO).as_posix()}  ->  [{names}]")
    print()
    print(PRINTED_INSTRUCTIONS)


PRINTED_INSTRUCTIONS = r"""
================================================================
ACCESS-SIDE VERIFICATION CHECKLIST
================================================================

For each file below, open in MS Access (double-click the .accdb).
When prompted "Security Warning: Some active content has been
disabled", click "Enable Content". Then press ALT+F11 to open the
VBA editor.

NOTE on the Access engine: pyOpenVBA rewrites the OVBA *cache*
(authoritative for pyOpenVBA's own read path) plus the dir-stream
catalog. Access ALSO maintains its own compiled p-code tables.
On first open after a Python edit, Access may show EITHER:
  (a) the new source we wrote (engine recompiled from the cache), or
  (b) the stale previously-compiled source.
If you see (b), run:  Debug > Compile <ProjectName>  in the VBA
editor, then close and reopen the database. The cache-driven
recompile should then surface our edits.

----------------------------------------------------------------
01_simple_multiline_replace.accdb
----------------------------------------------------------------
VBA EDITOR:
  Modules > M  should contain:
      Option Explicit

      Sub Hello()
          MsgBox "Hi from Python!"
          MsgBox "This is line 2"
          MsgBox "And line 3"
      End Sub
  Press F5 inside Hello() -- three MsgBoxes should fire in order.

----------------------------------------------------------------
02_class_module_body_swap.accdb
----------------------------------------------------------------
VBA EDITOR:
  Class Modules > C  should contain the new Greet() / LastGreeting
  body. CRITICAL: in the Properties window (F4) with class C
  selected, confirm:
      (Name)               C
      Instancing           1 - Private        (VB_Creatable=False)
  These attribute values must survive the body swap. If the class
  loses its Instancing setting or the (Name) attribute, the Phase 5f
  attribute-preserve path is broken.

TEST IN IMMEDIATE WINDOW (Ctrl+G):
      Dim x As New C
      x.Greet "world"
      Debug.Print x.LastGreeting    ' -> Hello, world!

----------------------------------------------------------------
03_full_source_with_own_header.accdb
----------------------------------------------------------------
VBA EDITOR:
  Class Modules > C  body should be the FromCaller() sub.
PROPERTIES WINDOW (F4):
      (Name)               C
      Instancing           1 - Private
  And the class should now be PredeclaredId=True (use it without
  Dim/New, e.g. in the Immediate window:
        C.FromCaller
  should pop the "Custom attribute header used." message box without
  needing `Dim x As New C` first).

----------------------------------------------------------------
04_push_access_bulk.accdb
----------------------------------------------------------------
VBA EDITOR:
  Modules > M  should contain BOTH Pushed() and Pushed2().
  Press F5 inside Pushed()  -> "pushed via pyopenvba.push_access()".
  Press F5 inside Pushed2() -> three Debug.Print lines in
  Immediate window (Ctrl+G):  i=1  i=2  i=3

----------------------------------------------------------------
05_rename_then_replace.accdb
----------------------------------------------------------------
PROJECT EXPLORER (Ctrl+R):
  Modules > Renamed_M    (NOT "M" anymore)
VBA EDITOR:
  Renamed_M body should be the StillWorks() sub.
  Press F5 in StillWorks() -> "Renamed_M survived a body swap."

ALSO verify in the Access main window (not the VBA editor):
  Navigation pane (left side) > switch to "Modules" view
    (click the title bar drop-down). The module should appear
    as "Renamed_M". This confirms the MSysObjects write path
    (Phase 5e) renamed the catalog row too.

================================================================
If anything above fails, capture:
  - the file name
  - what the VBA editor shows vs what is expected
  - whether Debug > Compile produced a different result
================================================================
"""


if __name__ == "__main__":
    main()
