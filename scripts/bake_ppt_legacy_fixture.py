"""
Bake the binary PowerPoint (.ppt) test fixture, using PowerPoint itself.

The .ppt reader has to be tested against a container PowerPoint actually
wrote: the VBA project lives inside the 'PowerPoint Document' stream,
reached through the persist chain, and a hand-built file would prove
nothing about how PowerPoint lays that out.  So this drives live
PowerPoint through pyvbaharness and saves as ppSaveAsPresentation (1).

The presentation is given real slides on purpose.  A macro-only deck
leaves almost nothing in the document stream around the VBA record, and
the interesting failure -- content lost when the record is resized -- can
only show up when there is content to lose.

Run from the repo root, on Windows with desktop PowerPoint installed:
    python scripts/bake_ppt_legacy_fixture.py

Dev-only: pyvbaharness is a test-time oracle and is never a runtime
dependency of pyOpenVBA.
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

OUT = ROOT / "tests" / "live_powerpoint_testing" / "legacy_macros.ppt"

# ppSaveAsPresentation -- the binary 97-2003 container.
PP_SAVE_AS_PRESENTATION = 1
# ppLayoutText -- a title placeholder plus a body placeholder.
PP_LAYOUT_TEXT = 2

FIXTURE_MODULE = """\
Option Explicit

Public Function Answer() As Long
    Answer = 42
End Function

Public Function SlideTally() As Long
    SlideTally = ActivePresentation.Slides.Count
End Function

Public Function Greet(ByVal who As String) As String
    Greet = "Hello, " & who & "!"
End Function
"""

BUILD_SLIDES = f"""\
Sub Main()
    Dim i As Long, s As Object
    For i = 1 To 3
        Set s = ActivePresentation.Slides.Add(i, {PP_LAYOUT_TEXT})
        s.Shapes(1).TextFrame.TextRange.Text = "Heading " & i
        s.Shapes(2).TextFrame.TextRange.Text = String(400, "x")
    Next i
End Sub
"""

# The harness's own runner modules cannot be removed from here: the code
# doing the removing is one of them, and unwinding the running module
# aborts the save.  They are stripped afterwards instead, once PowerPoint
# has authored the container.
SAVE_AS_LEGACY = """\
Sub Main()
    ActivePresentation.SaveAs "{path}", {fmt}
End Sub
"""

HARNESS_PREFIX = "PyVba"


def main() -> int:
    import pyvbaharness

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.unlink(missing_ok=True)
    session = pyvbaharness.PowerPointSession()
    try:
        session.new_document()
        session.add_module("Fixture", FIXTURE_MODULE)
        if not session.run_vba(BUILD_SLIDES, timeout=120.0).ok:
            raise SystemExit("PowerPoint refused to build the slides")
        saver = SAVE_AS_LEGACY.format(
            path=OUT.resolve(), fmt=PP_SAVE_AS_PRESENTATION
        )
        if not session.run_vba(saver, timeout=180.0).ok:
            raise SystemExit("PowerPoint refused to save the presentation")
    finally:
        try:
            session.close()
        except Exception:
            pass
    if not OUT.exists():
        raise SystemExit(f"PowerPoint wrote no file to {OUT}")
    strip_harness_modules(OUT)
    print(f"baked {OUT.relative_to(ROOT)} ({OUT.stat().st_size} bytes)")
    return 0


def strip_harness_modules(path: Path) -> None:
    """Drop the harness's runner modules from the baked fixture.

    This is the one step PowerPoint cannot do for us, and it uses the
    library's own .ppt write path.  That is acceptable because the test
    that matters -- reading a container PowerPoint authored -- still runs
    against PowerPoint's own record layout, and the strip is verified by
    reopening below.
    """
    from pyopenvba.powerpoint import PowerPointFile

    with PowerPointFile(path) as prs:
        project = prs.vba_project()
        doomed = [n for n in prs.module_names() if n.startswith(HARNESS_PREFIX)]
        for name in doomed:
            project.delete_module(name)
        prs.save()
    with PowerPointFile(path) as prs:
        left = [n for n in prs.module_names() if n.startswith(HARNESS_PREFIX)]
        if left:
            raise SystemExit(f"harness modules survived the strip: {left}")
        if "Fixture" not in prs.module_names():
            raise SystemExit("the Fixture module did not survive the strip")


if __name__ == "__main__":
    raise SystemExit(main())
