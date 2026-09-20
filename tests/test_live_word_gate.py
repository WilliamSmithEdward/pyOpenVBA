"""Live Word gate for the in-memory document (opt-in).

What only Word can answer: that a document this edited opens without a
repair prompt, and that the shapes a macro drew here are the shapes
Word reads back, in the places it puts them.

Opt-in: set ``RUN_LIVE_WORD=1`` on a Windows machine with desktop Word
installed.  Skipped everywhere else, including CI.
"""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

import pytest

from pyopenvba.apps.word import WordApplication

pytestmark = pytest.mark.skipif(
    os.environ.get("RUN_LIVE_WORD") != "1" or sys.platform != "win32",
    reason="live Word gate: set RUN_LIVE_WORD=1 on Windows with Word installed",
)

FIXTURES = Path(__file__).parent / "fixtures" / "shapes"
DOCUMENT = FIXTURES / "word_shapes.docm"

#: What the macro does here, which Word then has to report back.
EDIT = """
Sub Draw()
    Dim sh As Object
    Set sh = ActiveDocument.Shapes.AddShape(9, 300, 400, 60, 70)
    sh.Name = "Fresh"
    ActiveDocument.Shapes("Oval").Left = 111
    ActiveDocument.Shapes("Box").Width = 200
    ActiveDocument.Shapes("Rect").TextFrame.TextRange.Text = "Changed"
End Sub
"""

#: Read back on both sides, in this order.
WANTED = ("Fresh", "Oval", "Box", "Rect")


def test_word_opens_the_shapes_this_wrote(tmp_path: Path) -> None:
    harness = pytest.importorskip("pyvbaharness")

    copy = tmp_path / "edited.docm"
    shutil.copyfile(DOCUMENT, copy)
    app = WordApplication.open(copy)
    app.add_module(EDIT, name="EditModule")
    app.run("Draw")
    ours = [
        str(app.evaluate(f'CStr(ActiveDocument.Shapes("{name}").{what})'))
        for name in WANTED
        for what in ("Name", "Type", "Left", "Top", "Width", "Height")
    ]
    # Word's reader writes the paragraph mark as a slash, so that one
    # answer stays on one line; the same is done here.
    text = str(app.evaluate('ActiveDocument.Shapes("Rect").TextFrame.TextRange.Text'))
    ours.append(text.replace("\r", "/").replace("\v", "/"))
    ours.append(str(app.evaluate("ActiveDocument.Shapes.Count")))

    out = tmp_path / "edited_out.docm"
    app.save(out)

    reader = (
        "Public Function Report() As String\n"
        "    Dim out As String, sh As Object, i As Long, j As Long\n"
        f'    Dim names As Variant: names = Array({", ".join(chr(34) + n + chr(34) for n in WANTED)})\n'
        '    Dim what As Variant: what = Array("Name", "Type", "Left", "Top", "Width", "Height")\n'
        "    For i = LBound(names) To UBound(names)\n"
        "        Set sh = ActiveDocument.Shapes(CStr(names(i)))\n"
        "        For j = LBound(what) To UBound(what)\n"
        "            out = out & CStr(CallByName(sh, CStr(what(j)), VbGet)) & vbLf\n"
        "        Next j\n"
        "    Next i\n"
        '    out = out & Replace(Replace(ActiveDocument.Shapes("Rect").TextFrame.TextRange.Text, '
        'vbCr, "/"), Chr(11), "/") & vbLf\n'
        "    out = out & CStr(ActiveDocument.Shapes.Count) & vbLf\n"
        "    Report = out\n"
        "End Function\n"
    )
    with harness.WordSession() as word:
        word.open_document(out)
        result = word.run_vba(reader, "Report", timeout=180.0)
        assert result.ok, f"{result.outcome}: {result.message}"
        theirs = str(result.value).rstrip("\n").split("\n")

    assert [one.replace("\r", "") for one in theirs] == ours
