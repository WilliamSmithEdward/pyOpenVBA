"""Word's object model, with real state behind it.

What each member answers is what live Word answered in
``tests/fixtures/word_model/probes.txt``.  Word differs from the other
two hosts in ways that only measuring shows: a shape is placed from the
page but reported against the column and the paragraph, so the numbers
come back shifted by the margins; a new shape's number comes from a
counter that does not go back down when shapes are deleted; and there
is no macro on a shape at all.

A shape here is the same :class:`pyopenvba.shapes.Shape` the file layer
reads and writes.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, VBAObject, member, method, setter
from pyopenvba.interpreter._values import (
    EMPTY,
    MISSING,
    NOTHING,
    VBAInt,
    VBASingle,
    error,
    to_number,
    to_text,
)
from pyopenvba.shapes._docx_text import DEFAULT_MARGIN_POINTS
from pyopenvba.shapes._values import PRESET_GEOMETRY, WORD_SHAPE_NAMES, Shape

#: Word raises this for a shape that is not there, by name or number.
ERR_NO_SUCH_SHAPE = -2147024809

#: What Word calls a new shape, by the msoShapeType asked for.  Word's
#: own names, which are not Excel's: a rounded rectangle is "Rectangle:
#: Rounded Corners", a can is a "Cylinder", and a text box is "Text
#: Box", with the space.  Measured by scripts/measure_shape_types.py.
AUTO_SHAPE_NAMES = WORD_SHAPE_NAMES

#: wdRelativeHorizontalPositionColumn and its vertical twin, which is
#: what a shape added by a macro is placed against.
RELATIVE_TO_COLUMN = 2

#: wdWrapNone, which a new AutoShape gets.
WRAP_NONE = 3


class WordObject(VBAObject):
    """Anything in Word's model, for a common repr."""

    def __repr__(self) -> str:
        return f"<{type(self).__name__}>"


class Application(WordObject):
    """The Word application, with its documents."""

    vba_type_name = "Application"

    def __init__(self) -> None:
        self.documents_ = Documents(self)
        self.active: Document | None = None
        self.visible = -1

    @member
    def Name(self) -> object:
        return "Microsoft Word"

    @member
    def Version(self) -> object:
        return "16.0"

    @member
    def Documents(self, Index: object = MISSING) -> object:
        if Index is MISSING:
            return self.documents_
        return self.documents_.vba_get("Item", [Index])

    @member
    def ActiveDocument(self) -> object:
        return self.active if self.active is not None else NOTHING

    @member
    def Visible(self) -> object:
        return VBAInt(self.visible, "Long")

    @setter("Visible")
    def _set_visible(self, value: object) -> None:
        self.visible = -1 if to_number(value) else 0

    @member
    def Selection(self) -> object:
        raise VBAUnsupportedError(
            "Application.Selection is real Word that pyOpenVBA does not implement: there is "
            "no window here for anything to be selected in"
        )

    @method
    def Quit(self, *args: object) -> object:
        del args
        return EMPTY


class Documents(VBACollection, WordObject):
    """The documents that are open."""

    vba_type_name = "Documents"

    def __init__(self, application: Application) -> None:
        self.application = application
        self.items: list[Document] = []

    def vba_items(self) -> list[object]:
        return list(self.items)

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            for one in self.items:
                if one.name.lower() == index.lower():
                    return one
            raise error(ERR_NO_SUCH_SHAPE, f"there is no document called {index}")
        return super().vba_lookup(index, items)

    @method
    def Add(self, *args: object, **named: object) -> object:
        del args, named
        document = Document(self.application, "Document1")
        self.items.append(document)
        self.application.active = document
        return document

    @method
    def Open(self, FileName: object = MISSING, *rest: object) -> object:
        del rest
        from pyopenvba.apps.word._io import load_document

        document = load_document(self.application, Path(to_text(FileName)))
        self.items.append(document)
        self.application.active = document
        return document


class Document(WordObject):
    """One open document: its text, and the shapes on it."""

    vba_type_name = "Document"

    def __init__(self, application: Application, name: str, folder: str = "") -> None:
        self.application = application
        self.name = name
        self.folder = folder
        #: The text, one entry per paragraph.  Word always has at least
        #: one, even in an empty document.
        self.paragraphs_: list[str] = [""]
        self.shapes_: list[Shape] = []
        #: The ones a macro deleted, so a save can take their runs out
        #: of the markup they are still in.
        self.deleted_shapes: list[Shape] = []
        self.saved = False
        #: What the body looked like when it was read, so a document
        #: nobody rewrote is written back exactly as it arrived.
        self.document_xml = ""
        self.text_dirty = False
        self.shapes_dirty = False
        #: How many shapes have ever been added, which is where the next
        #: one's number comes from.  It does not go back down.
        self.shape_count = 0
        self.left_margin = DEFAULT_MARGIN_POINTS
        self.top_margin = DEFAULT_MARGIN_POINTS
        self.package: Any = None

    # -- Python side

    @property
    def text(self) -> str:
        """Everything in the document, as Word's Content.Text reads it."""
        return "".join(f"{one}\r" for one in self.paragraphs_)

    def set_text(self, value: str) -> None:
        body = value.replace("\r\n", "\r").replace("\n", "\r")
        self.paragraphs_ = body.split("\r")
        if self.paragraphs_ and self.paragraphs_[-1] == "":
            self.paragraphs_.pop()
        if not self.paragraphs_:
            self.paragraphs_ = [""]
        # Assigning the text takes the shapes with it, as Word does.
        self.deleted_shapes.extend(self.shapes_)
        self.shapes_ = []
        self.text_dirty = True
        self.saved = False

    def touched(self) -> None:
        self.shapes_dirty = True
        self.saved = False

    # -- VBA side

    @member
    def Name(self) -> object:
        return self.name

    @member
    def FullName(self) -> object:
        return str(Path(self.folder) / self.name) if self.folder else self.name

    @member
    def Path(self) -> object:
        return self.folder

    @member
    def Saved(self) -> object:
        return self.saved

    @setter("Saved")
    def _set_saved(self, value: object) -> None:
        self.saved = bool(to_number(value))

    @member
    def Content(self) -> object:
        return Range(self, 0, len(self.text))

    @member
    def Range(self, Start: object = MISSING, End: object = MISSING) -> object:
        whole = len(self.text)
        start = int(to_number(Start)) if Start is not MISSING else 0
        finish = int(to_number(End)) if End is not MISSING else whole
        return Range(self, max(0, start), min(finish, whole))

    @member
    def Paragraphs(self, Index: object = MISSING) -> object:
        paragraphs = Paragraphs(self)
        return paragraphs if Index is MISSING else paragraphs.vba_get("Item", [Index])

    @member
    def Shapes(self, Index: object = MISSING) -> object:
        shapes = Shapes(self)
        return shapes if Index is MISSING else shapes.vba_get("Item", [Index])

    @member
    def InlineShapes(self, Index: object = MISSING) -> object:
        shapes = InlineShapes(self)
        return shapes if Index is MISSING else shapes.vba_get("Item", [Index])

    @method
    def Save(self) -> object:
        from pyopenvba.apps.word._io import save_document

        if not self.folder:
            raise error(ERR_NO_SUCH_SHAPE, "this document has never been saved")
        save_document(self, Path(self.folder) / self.name)
        self.saved = True
        return EMPTY

    @method
    def SaveAs(self, FileName: object = MISSING, *rest: object) -> object:
        del rest
        from pyopenvba.apps.word._io import save_document

        target = Path(to_text(FileName))
        save_document(self, target)
        self.name, self.folder = target.name, str(target.parent)
        self.saved = True
        return EMPTY

    #: SaveAs2 takes the same arguments for what this models.
    SaveAs2 = SaveAs

    @method
    def Close(self, *args: object) -> object:
        del args
        if self in self.application.documents_.items:
            self.application.documents_.items.remove(self)
        if self.application.active is self:
            self.application.active = None
        return EMPTY


class Range(WordObject):
    """A stretch of the document's text."""

    vba_type_name = "Range"

    def __init__(self, document: Document, start: int, end: int) -> None:
        self.document = document
        self.start = start
        self.end = end

    @member(default=True)
    def Text(self) -> object:
        return self.document.text[self.start : self.end]

    @setter("Text")
    def _set_text(self, value: object) -> None:
        whole = self.document.text
        self.document.set_text(whole[: self.start] + to_text(value) + whole[self.end :])

    @member
    def Start(self) -> object:
        return VBAInt(self.start, "Long")

    @member
    def End(self) -> object:
        return VBAInt(self.end, "Long")

    @member
    def Paragraphs(self, Index: object = MISSING) -> object:
        return self.document.Paragraphs(Index)

    @method
    def Delete(self, *args: object) -> object:
        del args
        whole = self.document.text
        self.document.set_text(whole[: self.start] + whole[self.end :])
        return EMPTY

    @method
    def InsertAfter(self, Text: object = MISSING) -> object:
        whole = self.document.text
        self.document.set_text(whole[: self.end] + to_text(Text) + whole[self.end :])
        return EMPTY


class Paragraphs(VBACollection, WordObject):
    """The paragraphs of a document."""

    vba_type_name = "Paragraphs"

    def __init__(self, document: Document) -> None:
        self.document = document

    def vba_items(self) -> list[object]:
        return [Paragraph(self.document, index) for index in range(len(self.document.paragraphs_))]


class Paragraph(WordObject):
    """One paragraph."""

    vba_type_name = "Paragraph"

    def __init__(self, document: Document, index: int) -> None:
        self.document = document
        self.index = index

    @member
    def Range(self) -> object:
        start = sum(len(one) + 1 for one in self.document.paragraphs_[: self.index])
        return Range(self.document, start, start + len(self.document.paragraphs_[self.index]) + 1)


class Shapes(VBACollection, WordObject):
    """The floating shapes of a document, in the order Word lists them."""

    vba_type_name = "Shapes"

    def __init__(self, document: Document) -> None:
        self.document = document

    def floating(self) -> list[Shape]:
        return [one for one in self.document.shapes_ if one.placement != "inline"]

    def vba_items(self) -> list[object]:
        return [ShapeObject(self.document, one) for one in self.floating()]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            for shape in self.floating():
                if shape.name.lower() == index.lower():
                    return ShapeObject(self.document, shape)
            raise error(ERR_NO_SUCH_SHAPE, f"there is no shape called {index}")
        position = int(to_number(index))
        if 1 <= position <= len(items):
            return items[position - 1]
        raise error(ERR_NO_SUCH_SHAPE, f"there is no shape number {position}")

    @method
    def AddShape(
        self,
        Type: object = MISSING,
        Left: object = MISSING,
        Top: object = MISSING,
        Width: object = MISSING,
        Height: object = MISSING,
        Anchor: object = MISSING,
    ) -> object:
        del Anchor
        wanted = int(to_number(Type)) if Type is not MISSING else 1
        shape = self._made(
            kind="shape",
            geometry=PRESET_GEOMETRY.get(wanted, "rect"),
            stem=AUTO_SHAPE_NAMES.get(wanted, "Rectangle"),
            box=(Left, Top, Width, Height),
        )
        shape.auto_shape_type = wanted
        return ShapeObject(self.document, shape)

    @method
    def AddTextbox(
        self,
        Orientation: object = MISSING,
        Left: object = MISSING,
        Top: object = MISSING,
        Width: object = MISSING,
        Height: object = MISSING,
        Anchor: object = MISSING,
    ) -> object:
        del Orientation, Anchor
        shape = self._made(
            kind="textBox", geometry="rect", stem="Text Box", box=(Left, Top, Width, Height)
        )
        return ShapeObject(self.document, shape)

    @method
    def AddLine(
        self,
        BeginX: object = MISSING,
        BeginY: object = MISSING,
        EndX: object = MISSING,
        EndY: object = MISSING,
        Anchor: object = MISSING,
    ) -> object:
        del Anchor
        start_x, start_y = _number(BeginX), _number(BeginY)
        end_x, end_y = _number(EndX), _number(EndY)
        shape = self._made(
            kind="line",
            geometry="line",
            stem="Straight Connector",
            box=(
                min(start_x, end_x),
                min(start_y, end_y),
                abs(end_x - start_x),
                abs(end_y - start_y),
            ),
        )
        return ShapeObject(self.document, shape)

    @method
    def AddPicture(self, *args: object, **named: object) -> object:
        del args, named
        raise VBAUnsupportedError(
            "Shapes.AddPicture is real Word that pyOpenVBA does not implement"
        )

    @method
    def Range(self, Index: object = MISSING) -> object:
        del Index
        raise VBAUnsupportedError("Shapes.Range is real Word that pyOpenVBA does not implement")

    def _made(
        self, *, kind: str, geometry: str, stem: str, box: tuple[object, object, object, object]
    ) -> Shape:
        left, top, width, height = (_number(one) for one in box)
        self.document.shape_count += 1
        shape = Shape(
            name=f"{stem} {self.document.shape_count}",
            kind=kind,
            geometry=geometry,
            # Word is given a place on the page and reports it against
            # the column and the paragraph.
            left=left - self.document.left_margin,
            top=top - self.document.top_margin,
            width=width,
            height=height,
            placement="anchored",
            shape_id=max((one.shape_id for one in self.document.shapes_), default=0) + 1,
        )
        shape.z_order = max((one.z_order for one in self.document.shapes_), default=0) + 1
        self.document.shapes_.append(shape)
        self.document.touched()
        return shape


class InlineShapes(VBACollection, WordObject):
    """The shapes that sit in the text rather than beside it."""

    vba_type_name = "InlineShapes"

    def __init__(self, document: Document) -> None:
        self.document = document

    def vba_items(self) -> list[object]:
        return [
            ShapeObject(self.document, one)
            for one in self.document.shapes_
            if one.placement == "inline"
        ]


class ShapeObject(WordObject):
    """One shape on a document."""

    vba_type_name = "Shape"

    def __init__(self, document: Document, shape: Shape) -> None:
        self.document = document
        self.shape = shape

    @member
    def Name(self) -> object:
        return self.shape.name

    @setter("Name")
    def _set_name(self, value: object) -> None:
        self.shape.name = to_text(value)
        self.document.touched()

    @member
    def Type(self) -> object:
        return VBAInt(self.shape.mso_type, "Long")

    @member
    def AutoShapeType(self) -> object:
        return VBAInt(self.shape.auto_shape_type, "Long")

    @member
    def ZOrderPosition(self) -> object:
        for index, one in enumerate(
            [one for one in self.document.shapes_ if one.placement != "inline"], start=1
        ):
            if one is self.shape:
                return VBAInt(index, "Long")
        return VBAInt(1, "Long")

    @member
    def RelativeHorizontalPosition(self) -> object:
        return VBAInt(RELATIVE_TO_COLUMN, "Long")

    @member
    def RelativeVerticalPosition(self) -> object:
        return VBAInt(RELATIVE_TO_COLUMN, "Long")

    @member
    def WrapFormat(self) -> object:
        return WrapFormat()

    @member
    def Left(self) -> object:
        return VBASingle(self.shape.left)

    @setter("Left")
    def _set_left(self, value: object) -> None:
        self.shape.left = _number(value)
        self.document.touched()

    @member
    def Top(self) -> object:
        return VBASingle(self.shape.top)

    @setter("Top")
    def _set_top(self, value: object) -> None:
        self.shape.top = _number(value)
        self.document.touched()

    @member
    def Width(self) -> object:
        return VBASingle(self.shape.width)

    @setter("Width")
    def _set_width(self, value: object) -> None:
        self.shape.width = _number(value)
        self.document.touched()

    @member
    def Height(self) -> object:
        return VBASingle(self.shape.height)

    @setter("Height")
    def _set_height(self, value: object) -> None:
        self.shape.height = _number(value)
        self.document.touched()

    @member
    def TextFrame(self) -> object:
        return TextFrame(self)

    @method
    def Delete(self) -> object:
        self.document.shapes_ = [one for one in self.document.shapes_ if one is not self.shape]
        self.document.deleted_shapes.append(self.shape)
        self.document.touched()
        return EMPTY

    @method
    def Select(self, *args: object) -> object:
        del args
        return EMPTY


class WrapFormat(WordObject):
    """How the text flows around a shape."""

    vba_type_name = "WrapFormat"

    @member
    def Type(self) -> object:
        return VBAInt(WRAP_NONE, "Long")


class TextFrame(WordObject):
    """A shape's text frame."""

    vba_type_name = "TextFrame"

    def __init__(self, shape: ShapeObject) -> None:
        self.owner = shape

    @member
    def TextRange(self) -> object:
        return TextRange(self.owner)

    @member
    def HasText(self) -> object:
        return VBAInt(-1 if self.owner.shape.text else 0, "Long")


class TextRange(WordObject):
    """The text inside a shape.

    Reading it gives the paragraph mark back with the text, the way
    every Word range does: a shape holding "Hello" reads as "Hello" and
    a carriage return.  Assigning does not want one.
    """

    vba_type_name = "Range"

    def __init__(self, shape: ShapeObject) -> None:
        self.owner = shape

    @member(default=True)
    def Text(self) -> object:
        return self.owner.shape.text + "\r"

    @setter("Text")
    def _set_text(self, value: object) -> None:
        self.owner.shape.text = to_text(value).rstrip("\r")
        self.owner.document.touched()


def _number(value: object) -> float:
    if value is MISSING or value is None:
        return 0.0
    return float(to_number(value))
