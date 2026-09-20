"""Shapes through the document: loaded, edited by a macro, saved.

The document reader and writer are held to Word's own file in
``test_shapes_docx.py``; this is the other half, that a document opened
here shows the same shapes to VBA and that an edit lands in the file.
"""

from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

import pytest

from pyopenvba.apps.word import WordApplication
from pyopenvba.exceptions import VBARuntimeError

FIXTURES = Path(__file__).parent / "fixtures" / "shapes"
DOCUMENT = FIXTURES / "word_shapes.docm"


def opened(tmp_path: Path) -> WordApplication:
    copy = tmp_path / "shapes.docm"
    shutil.copyfile(DOCUMENT, copy)
    return WordApplication.open(copy)


def macro(app: WordApplication, body: str, name: str = "S") -> None:
    app.add_module(f"Sub {name}()\n{body}\nEnd Sub\n", name="ShapeModule")
    app.run(name)


def body_of(path: Path) -> str:
    return zipfile.ZipFile(path).read("word/document.xml").decode("utf-8")


def test_a_document_opens_with_its_shapes_and_text(tmp_path: Path) -> None:
    """The counts and the text are what Word answered for this file.

    Asked of live Word: four paragraphs, 38 characters, and the first
    is the in-line rule, which is one Chr(1) and the paragraph mark.
    """
    app = opened(tmp_path)
    assert [one.name for one in app.shapes()] == ["Rect", "Oval", "Box", "Line1", "Plain1"]
    assert len(app.inline_shapes()) == 1
    assert app.document.text == "\x01\rFirst paragraph.\rSecond paragraph.\r\r"
    assert len(app.document.text) == 38


def test_a_macro_sees_what_the_file_holds(tmp_path: Path) -> None:
    app = opened(tmp_path)
    assert app.evaluate("ActiveDocument.Shapes.Count") == 5
    assert app.evaluate("ActiveDocument.InlineShapes.Count") == 1
    # A Word range reads with its paragraph mark, shape or not.
    assert app.evaluate('ActiveDocument.Shapes("Rect").TextFrame.TextRange.Text') == "Hello\r"
    assert app.evaluate('ActiveDocument.Shapes("Rect").Left') == -32
    assert app.evaluate("ActiveDocument.Paragraphs.Count") == 4
    assert app.evaluate("Len(ActiveDocument.Paragraphs(1).Range.Text)") == 2


def test_a_missing_shape_raises_what_word_raises(tmp_path: Path) -> None:
    app = opened(tmp_path)
    with pytest.raises(VBARuntimeError) as raised:
        app.evaluate('ActiveDocument.Shapes("NoSuchShape").Name')
    assert raised.value.number == -2147024809


def test_an_untouched_document_keeps_its_body(tmp_path: Path) -> None:
    app = opened(tmp_path)
    out = tmp_path / "same.docm"
    app.save(out)
    assert body_of(out) == body_of(DOCUMENT)


def test_a_moved_shape_reaches_the_file(tmp_path: Path) -> None:
    app = opened(tmp_path)
    macro(app, '    ActiveDocument.Shapes("Oval").Left = 111')
    out = tmp_path / "moved.docm"
    app.save(out)

    again = WordApplication.open(out)
    assert again.evaluate('ActiveDocument.Shapes("Oval").Left') == 111
    assert again.evaluate('ActiveDocument.Shapes("Rect").Left') == -32


def test_a_resized_shape_reaches_the_file(tmp_path: Path) -> None:
    app = opened(tmp_path)
    macro(app, '    ActiveDocument.Shapes("Box").Width = 200')
    out = tmp_path / "resized.docm"
    app.save(out)

    again = WordApplication.open(out)
    assert again.evaluate('ActiveDocument.Shapes("Box").Width') == 200


def test_a_new_shape_reaches_the_file(tmp_path: Path) -> None:
    """A shape a macro adds is placed from the page, as Word places one."""
    app = opened(tmp_path)
    macro(
        app,
        "    Dim sh As Object\n"
        "    Set sh = ActiveDocument.Shapes.AddShape(9, 100, 200, 60, 70)\n"
        '    sh.Name = "Fresh"',
    )
    assert app.evaluate('ActiveDocument.Shapes("Fresh").Left') == 100 - 72
    out = tmp_path / "added.docm"
    app.save(out)

    again = WordApplication.open(out)
    assert again.evaluate("ActiveDocument.Shapes.Count") == 6
    assert again.evaluate('ActiveDocument.Shapes("Fresh").Left') == 28
    assert again.evaluate('ActiveDocument.Shapes("Fresh").Width') == 60


def test_a_deleted_shape_leaves_the_file(tmp_path: Path) -> None:
    app = opened(tmp_path)
    macro(app, '    ActiveDocument.Shapes("Oval").Delete')
    out = tmp_path / "deleted.docm"
    app.save(out)

    again = WordApplication.open(out)
    assert again.evaluate("ActiveDocument.Shapes.Count") == 4
    assert '"Oval"' not in body_of(out)


def test_a_shapes_text_can_be_changed(tmp_path: Path) -> None:
    app = opened(tmp_path)
    macro(app, '    ActiveDocument.Shapes("Rect").TextFrame.TextRange.Text = "Changed"')
    out = tmp_path / "retexted.docm"
    app.save(out)

    again = WordApplication.open(out)
    assert again.evaluate('ActiveDocument.Shapes("Rect").TextFrame.TextRange.Text') == "Changed\r"


def test_assigning_the_text_takes_the_shapes_with_it(tmp_path: Path) -> None:
    """Which is what Word does: asked of it, five shapes became none."""
    app = opened(tmp_path)
    macro(app, '    ActiveDocument.Content.Text = "Only this."')
    out = tmp_path / "retyped.docm"
    app.save(out)

    again = WordApplication.open(out)
    assert again.document.text == "Only this.\r"
    assert again.evaluate("ActiveDocument.Shapes.Count") == 0
    assert again.evaluate("ActiveDocument.InlineShapes.Count") == 0
