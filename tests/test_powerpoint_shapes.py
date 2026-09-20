"""Shapes through the presentation: loaded, edited by a macro, saved.

The slide reader and writer are held to PowerPoint's own file in
``test_shapes_pptx.py``; this is the other half, that a presentation
opened here shows the same shapes to VBA and that an edit lands in the
file.
"""

from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

import pytest

from pyopenvba.apps.powerpoint import PowerPointApplication
from pyopenvba.exceptions import VBARuntimeError

FIXTURES = Path(__file__).parent / "fixtures" / "shapes"
DECK = FIXTURES / "powerpoint_shapes.pptm"


def opened(tmp_path: Path) -> PowerPointApplication:
    copy = tmp_path / "shapes.pptm"
    shutil.copyfile(DECK, copy)
    return PowerPointApplication.open(copy)


def macro(app: PowerPointApplication, body: str, name: str = "S") -> None:
    app.add_module(f"Sub {name}()\n{body}\nEnd Sub\n", name="ShapeModule")
    app.run(name)


def slide_of(path: Path) -> str:
    return zipfile.ZipFile(path).read("ppt/slides/slide1.xml").decode("utf-8")


def test_a_presentation_opens_with_its_slides_and_shapes(tmp_path: Path) -> None:
    app = opened(tmp_path)
    assert len(app.presentation.slides_) == 1
    assert [one.name for one in app.slide(1).shapes()] == [
        "Title 1",
        "Subtitle 2",
        "Rect",
        "Oval",
        "Box",
        "Line1",
        "Group1",
    ]


def test_a_macro_sees_what_the_file_holds(tmp_path: Path) -> None:
    app = opened(tmp_path)
    assert app.evaluate("ActivePresentation.Slides(1).Shapes.Count") == 7
    assert app.evaluate('ActivePresentation.Slides(1).Shapes("Rect").TextFrame.TextRange.Text') == "Hello"
    assert app.evaluate('ActivePresentation.Slides(1).Shapes("Rect").ActionSettings(1).Run') == "Clicked"
    assert app.evaluate('ActivePresentation.Slides(1).Shapes("Rect").ActionSettings(1).Action') == 8
    assert app.evaluate('ActivePresentation.Slides(1).Shapes("Oval").ActionSettings(1).Action') == 0


def test_a_missing_shape_raises_what_powerpoint_raises(tmp_path: Path) -> None:
    app = opened(tmp_path)
    with pytest.raises(VBARuntimeError) as raised:
        app.evaluate('ActivePresentation.Slides(1).Shapes("NoSuchShape").Name')
    assert raised.value.number == -2147188160


def test_an_untouched_presentation_keeps_its_slide(tmp_path: Path) -> None:
    app = opened(tmp_path)
    out = tmp_path / "same.pptm"
    app.save(out)
    assert slide_of(out) == slide_of(DECK)


def test_a_moved_shape_reaches_the_file(tmp_path: Path) -> None:
    app = opened(tmp_path)
    macro(app, '    ActivePresentation.Slides(1).Shapes("Oval").Left = 333')
    out = tmp_path / "moved.pptm"
    app.save(out)

    again = PowerPointApplication.open(out)
    assert again.evaluate('ActivePresentation.Slides(1).Shapes("Oval").Left') == 333
    assert again.evaluate('ActivePresentation.Slides(1).Shapes("Rect").Left') == 20


def test_a_new_shape_reaches_the_file(tmp_path: Path) -> None:
    app = opened(tmp_path)
    macro(
        app,
        "    Dim sh As Object\n"
        "    Set sh = ActivePresentation.Slides(1).Shapes.AddShape(9, 400, 300, 60, 70)\n"
        '    sh.Name = "Fresh"\n'
        '    sh.TextFrame.TextRange.Text = "Added"\n'
        '    sh.ActionSettings(ppMouseClick).Run = "Clicked"',
    )
    out = tmp_path / "added.pptm"
    app.save(out)

    again = PowerPointApplication.open(out)
    assert again.evaluate('ActivePresentation.Slides(1).Shapes("Fresh").Left') == 400
    assert again.evaluate('ActivePresentation.Slides(1).Shapes("Fresh").Top') == 300
    assert again.evaluate('ActivePresentation.Slides(1).Shapes("Fresh").TextFrame.TextRange.Text') == "Added"
    assert again.evaluate('ActivePresentation.Slides(1).Shapes("Fresh").ActionSettings(1).Run') == "Clicked"
    assert again.evaluate("ActivePresentation.Slides(1).Shapes.Count") == 8


def test_a_deleted_shape_leaves_the_file(tmp_path: Path) -> None:
    app = opened(tmp_path)
    macro(app, '    ActivePresentation.Slides(1).Shapes("Oval").Delete')
    out = tmp_path / "deleted.pptm"
    app.save(out)

    again = PowerPointApplication.open(out)
    assert again.evaluate("ActivePresentation.Slides(1).Shapes.Count") == 6
    assert '"Oval"' not in slide_of(out)


def test_the_macro_a_click_runs_can_be_changed(tmp_path: Path) -> None:
    app = opened(tmp_path)
    macro(
        app,
        '    ActivePresentation.Slides(1).Shapes("Rect").ActionSettings(1).Run = "Other"\n'
        '    ActivePresentation.Slides(1).Shapes("Oval").ActionSettings(1).Run = "Clicked"',
    )
    out = tmp_path / "wired.pptm"
    app.save(out)

    again = PowerPointApplication.open(out)
    assert again.evaluate('ActivePresentation.Slides(1).Shapes("Rect").ActionSettings(1).Run') == "Other"
    assert again.evaluate('ActivePresentation.Slides(1).Shapes("Oval").ActionSettings(1).Run') == "Clicked"


def test_the_presentations_macros_are_loaded(tmp_path: Path) -> None:
    """The fixture carries the module the measuring run wrote into it."""
    app = opened(tmp_path)
    assert app.interpreter.modules
