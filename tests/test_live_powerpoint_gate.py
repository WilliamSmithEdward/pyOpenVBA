"""Live PowerPoint gate for the in-memory presentation (opt-in).

What only PowerPoint can answer: that a presentation this edited opens
without a repair prompt, and that the shapes a macro drew here are the
shapes PowerPoint reads back, with the macro still on the one that had
it.

Opt-in: set ``RUN_LIVE_POWERPOINT=1`` on a Windows machine with desktop
PowerPoint installed.  Skipped everywhere else, including CI.
"""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

import pytest

from pyopenvba.apps.powerpoint import PowerPointApplication

pytestmark = pytest.mark.skipif(
    os.environ.get("RUN_LIVE_POWERPOINT") != "1" or sys.platform != "win32",
    reason="live PowerPoint gate: set RUN_LIVE_POWERPOINT=1 on Windows with PowerPoint installed",
)

FIXTURES = Path(__file__).parent / "fixtures" / "shapes"
DECK = FIXTURES / "powerpoint_shapes.pptm"

#: What the macro does here, which PowerPoint then has to report back.
EDIT = """
Sub Draw()
    Dim sh As Object
    Set sh = ActivePresentation.Slides(1).Shapes.AddShape(9, 400, 300, 60, 70)
    sh.Name = "Fresh"
    sh.TextFrame.TextRange.Text = "Added"
    sh.ActionSettings(ppMouseClick).Run = "Clicked"
    ActivePresentation.Slides(1).Shapes("Oval").Left = 333
    ActivePresentation.Slides(1).Shapes("Rect").ActionSettings(ppMouseClick).Run = "Other"
End Sub
"""

#: Read back on both sides, in this order.
WANTED = ("Fresh", "Oval", "Rect")


def test_powerpoint_opens_the_shapes_this_wrote(tmp_path: Path) -> None:
    harness = pytest.importorskip("pyvbaharness")

    copy = tmp_path / "edited.pptm"
    shutil.copyfile(DECK, copy)
    app = PowerPointApplication.open(copy)
    app.add_module(EDIT, name="EditModule")
    app.run("Draw")
    ours = [
        str(app.evaluate(f'CStr(ActivePresentation.Slides(1).Shapes("{name}").{what})'))
        for name in WANTED
        for what in ("Name", "Type", "Left", "Top", "Width", "Height")
    ]
    ours.extend(
        str(app.evaluate(f'ActivePresentation.Slides(1).Shapes("{name}").ActionSettings(1).Run'))
        for name in WANTED
    )
    out = tmp_path / "edited_out.pptm"
    app.save(out)

    reader = (
        "Public Function Report() As String\n"
        "    Dim out As String, one As Variant, what As Variant, sh As Object\n"
        f'    Dim names As Variant: names = Array({", ".join(chr(34) + n + chr(34) for n in WANTED)})\n'
        '    what = Array("Name", "Type", "Left", "Top", "Width", "Height")\n'
        "    Dim i As Long, j As Long\n"
        "    For i = LBound(names) To UBound(names)\n"
        "        Set sh = ActivePresentation.Slides(1).Shapes(CStr(names(i)))\n"
        "        For j = LBound(what) To UBound(what)\n"
        "            out = out & CStr(CallByName(sh, CStr(what(j)), VbGet)) & vbLf\n"
        "        Next j\n"
        "    Next i\n"
        "    For i = LBound(names) To UBound(names)\n"
        "        Set sh = ActivePresentation.Slides(1).Shapes(CStr(names(i)))\n"
        "        out = out & sh.ActionSettings(1).Run & vbLf\n"
        "    Next i\n"
        "    Report = out\n"
        "End Function\n"
    )
    with harness.PowerPointSession() as powerpoint:
        powerpoint.open_document(out)
        result = powerpoint.run_vba(reader, "Report", timeout=180.0)
        assert result.ok, f"{result.outcome}: {result.message}"
        theirs = str(result.value).rstrip("\n").split("\n")

    assert [one.replace("\r", "") for one in theirs] == ours
