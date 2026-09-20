"""PowerPoint's object model, with real state behind it.

What each member answers is what live PowerPoint answered in
``tests/fixtures/powerpoint_model/probes.txt``: the names it gives a
new shape, the Type numbers, the Single a position is reported in, the
error a shape that is not there raises, and what an ActionSetting says
before anything has been assigned to it.

A shape here is the same :class:`pyopenvba.shapes.Shape` the file layer
reads and writes, so a macro's edit and the slide's markup are never
two models of the same thing.
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
from pyopenvba.shapes._values import PRESET_GEOMETRY, SHAPE_NAMES, Shape

#: PowerPoint raises this for a shape that is not there, by name or by
#: number.  Measured: it is not Excel's number and not 1004.
ERR_NO_SUCH_SHAPE = -2147188160

#: What PowerPoint calls a new shape, by the msoShapeType asked for.
#: Measured in live PowerPoint, which names every one of them exactly
#: as Excel does; Word is the one that differs.
AUTO_SHAPE_NAMES = SHAPE_NAMES

#: A 16:9 slide, which is what PowerPoint makes now: 960 by 540 points.
DEFAULT_SLIDE_WIDTH = 960.0
DEFAULT_SLIDE_HEIGHT = 540.0

#: ppActionRunMacro, the only action this models.
PP_ACTION_RUN_MACRO = 8


class PowerPointObject(VBAObject):
    """Anything in PowerPoint's model, for a common repr."""

    def __repr__(self) -> str:
        return f"<{type(self).__name__}>"


class Application(PowerPointObject):
    """The PowerPoint application, with its presentations."""

    vba_type_name = "Application"

    def __init__(self) -> None:
        self.presentations_ = Presentations(self)
        self.active: Presentation | None = None
        self.visible = -1

    @member
    def Name(self) -> object:
        return "Microsoft PowerPoint"

    @member
    def Version(self) -> object:
        return "16.0"

    @member
    def Presentations(self, Index: object = MISSING) -> object:
        if Index is MISSING:
            return self.presentations_
        return self.presentations_.vba_get("Item", [Index])

    @member
    def ActivePresentation(self) -> object:
        return self.active if self.active is not None else NOTHING

    @member
    def ActiveWindow(self) -> object:
        if self.active is None:
            raise error(-2147188160, "there is no active window")
        return DocumentWindow(self.active)

    @member
    def Visible(self) -> object:
        return VBAInt(self.visible, "Long")

    @setter("Visible")
    def _set_visible(self, value: object) -> None:
        self.visible = -1 if to_number(value) else 0

    @method
    def Quit(self) -> object:
        return EMPTY

    @member
    def Windows(self, Index: object = MISSING) -> object:
        del Index
        raise VBAUnsupportedError(
            "Application.Windows is real PowerPoint that pyOpenVBA does not implement"
        )


class DocumentWindow(PowerPointObject):
    """The window a presentation is shown in.

    Only enough of one to answer ``ActiveWindow.View.Slide``, which is
    how a macro reaches the slide it is looking at.
    """

    vba_type_name = "DocumentWindow"

    def __init__(self, presentation: Presentation) -> None:
        self.presentation = presentation

    @member
    def View(self) -> object:
        return View(self.presentation)

    @member
    def Presentation(self) -> object:
        return self.presentation

    @member
    def Selection(self) -> object:
        raise VBAUnsupportedError(
            "DocumentWindow.Selection is real PowerPoint that pyOpenVBA does not implement"
        )


class View(PowerPointObject):
    """What the window is showing."""

    vba_type_name = "View"

    def __init__(self, presentation: Presentation) -> None:
        self.presentation = presentation

    @member
    def Slide(self) -> object:
        slide = self.presentation.active_slide
        if slide is None:
            raise error(ERR_NO_SUCH_SHAPE, "there is no slide in view")
        return slide

    @setter("Slide")
    def _set_slide(self, value: object) -> None:
        if isinstance(value, Slide):
            self.presentation.active_slide = value


class Presentations(VBACollection, PowerPointObject):
    """The presentations that are open."""

    vba_type_name = "Presentations"

    def __init__(self, application: Application) -> None:
        self.application = application
        self.items: list[Presentation] = []

    def vba_items(self) -> list[object]:
        return list(self.items)

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            for one in self.items:
                if one.name.lower() == index.lower():
                    return one
            raise error(ERR_NO_SUCH_SHAPE, f"there is no presentation called {index}")
        return super().vba_lookup(index, items)

    @method
    def Add(self, WithWindow: object = MISSING) -> object:
        del WithWindow
        presentation = Presentation(self.application, "Presentation1")
        presentation.add_slide()
        self.items.append(presentation)
        self.application.active = presentation
        return presentation

    @method
    def Open(self, FileName: object = MISSING, *rest: object) -> object:
        del rest
        from pyopenvba.apps.powerpoint._io import load_presentation

        presentation = load_presentation(self.application, Path(to_text(FileName)))
        self.items.append(presentation)
        self.application.active = presentation
        return presentation


class Presentation(PowerPointObject):
    """One open presentation."""

    vba_type_name = "Presentation"

    def __init__(self, application: Application, name: str, folder: str = "") -> None:
        self.application = application
        self.name = name
        self.folder = folder
        self.slides_: list[Slide] = []
        self.active_slide: Slide | None = None
        self.saved = 0
        #: The package the presentation was read from, so every part
        #: this does not model survives a save.
        self.package: Any = None

    def add_slide(self, layout: int = 12) -> Slide:
        slide = Slide(self, len(self.slides_) + 1, layout)
        self.slides_.append(slide)
        if self.active_slide is None:
            self.active_slide = slide
        return slide

    def touched(self) -> None:
        self.saved = 0

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
        return VBAInt(self.saved, "Long")

    @setter("Saved")
    def _set_saved(self, value: object) -> None:
        self.saved = -1 if to_number(value) else 0

    @member
    def Slides(self, Index: object = MISSING) -> object:
        slides = Slides(self)
        return slides if Index is MISSING else slides.vba_get("Item", [Index])

    @member
    def PageSetup(self) -> object:
        return PageSetup(self)

    @member
    def SlideMaster(self) -> object:
        raise VBAUnsupportedError(
            "Presentation.SlideMaster is real PowerPoint that pyOpenVBA does not implement"
        )

    @method
    def Save(self) -> object:
        from pyopenvba.apps.powerpoint._io import save_presentation

        if not self.folder:
            raise error(-2147188160, "this presentation has never been saved")
        save_presentation(self, Path(self.folder) / self.name)
        self.saved = -1
        return EMPTY

    @method
    def SaveAs(self, FileName: object = MISSING, *rest: object) -> object:
        del rest
        from pyopenvba.apps.powerpoint._io import save_presentation

        target = Path(to_text(FileName))
        save_presentation(self, target)
        self.name, self.folder = target.name, str(target.parent)
        self.saved = -1
        return EMPTY

    @method
    def Close(self) -> object:
        if self in self.application.presentations_.items:
            self.application.presentations_.items.remove(self)
        if self.application.active is self:
            self.application.active = None
        return EMPTY


class PageSetup(PowerPointObject):
    """How big a slide is."""

    vba_type_name = "PageSetup"

    def __init__(self, presentation: Presentation) -> None:
        self.presentation = presentation
        self.width = DEFAULT_SLIDE_WIDTH
        self.height = DEFAULT_SLIDE_HEIGHT

    @member
    def SlideWidth(self) -> object:
        return VBASingle(self.width)

    @setter("SlideWidth")
    def _set_slide_width(self, value: object) -> None:
        self.width = float(to_number(value))

    @member
    def SlideHeight(self) -> object:
        return VBASingle(self.height)

    @setter("SlideHeight")
    def _set_slide_height(self, value: object) -> None:
        self.height = float(to_number(value))


class Slides(VBACollection, PowerPointObject):
    """The slides of one presentation, in the order they are shown."""

    vba_type_name = "Slides"

    def __init__(self, presentation: Presentation) -> None:
        self.presentation = presentation

    def vba_items(self) -> list[object]:
        return list(self.presentation.slides_)

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            for one in self.presentation.slides_:
                if one.name.lower() == index.lower():
                    return one
            raise error(ERR_NO_SUCH_SHAPE, f"there is no slide called {index}")
        return super().vba_lookup(index, items)

    @method
    def Add(self, Index: object = MISSING, Layout: object = MISSING) -> object:
        layout = int(to_number(Layout)) if Layout is not MISSING else 12
        slide = Slide(self.presentation, 0, layout)
        at = int(to_number(Index)) if Index is not MISSING else len(self.presentation.slides_) + 1
        at = max(1, min(at, len(self.presentation.slides_) + 1))
        self.presentation.slides_.insert(at - 1, slide)
        self.presentation.touched()
        return slide

    @method
    def Range(self, Index: object = MISSING) -> object:
        del Index
        raise VBAUnsupportedError(
            "Slides.Range is real PowerPoint that pyOpenVBA does not implement"
        )


class Slide(PowerPointObject):
    """One slide, and the shapes on it."""

    vba_type_name = "Slide"

    def __init__(self, presentation: Presentation, number: int, layout: int = 12) -> None:
        self.presentation = presentation
        self.layout = layout
        self.slide_name = f"Slide{number or len(presentation.slides_) + 1}"
        self.shapes_: list[Shape] = []
        #: How many shapes have been added here, which is where the next
        #: one's number comes from.  Per slide, and it does not go back
        #: down when one is deleted.
        self.shape_count = 0
        #: The part this slide came from and its markup, so a slide
        #: nobody touched is written back exactly as it arrived.
        self.part_name = ""
        self.slide_xml = ""
        self.layout_xml = ""
        self.dirty = False

    @property
    def name(self) -> str:
        return self.slide_name

    def touched(self) -> None:
        self.dirty = True
        self.presentation.touched()

    @member
    def Name(self) -> object:
        return self.slide_name

    @setter("Name")
    def _set_name(self, value: object) -> None:
        self.slide_name = to_text(value)
        self.touched()

    @member
    def SlideIndex(self) -> object:
        return VBAInt(self._position(), "Long")

    @member
    def SlideNumber(self) -> object:
        return VBAInt(self._position(), "Long")

    @member
    def Layout(self) -> object:
        return VBAInt(self.layout, "Long")

    @setter("Layout")
    def _set_layout(self, value: object) -> None:
        self.layout = int(to_number(value))
        self.touched()

    @member
    def Shapes(self, Index: object = MISSING) -> object:
        shapes = Shapes(self)
        return shapes if Index is MISSING else shapes.vba_get("Item", [Index])

    @method
    def Select(self) -> object:
        self.presentation.active_slide = self
        return EMPTY

    @method
    def Delete(self) -> object:
        if self in self.presentation.slides_:
            self.presentation.slides_.remove(self)
        if self.presentation.active_slide is self:
            self.presentation.active_slide = (
                self.presentation.slides_[0] if self.presentation.slides_ else None
            )
        self.presentation.touched()
        return EMPTY

    def _position(self) -> int:
        for index, one in enumerate(self.presentation.slides_, start=1):
            if one is self:
                return index
        return 1


class Shapes(VBACollection, PowerPointObject):
    """The shapes on one slide."""

    vba_type_name = "Shapes"

    def __init__(self, slide: Slide) -> None:
        self.slide = slide

    def vba_items(self) -> list[object]:
        return [ShapeObject(self.slide, one) for one in self.slide.shapes_]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            for shape in self.slide.shapes_:
                if shape.name.lower() == index.lower():
                    return ShapeObject(self.slide, shape)
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
    ) -> object:
        wanted = int(to_number(Type)) if Type is not MISSING else 1
        shape = self._made(
            kind="shape",
            geometry=PRESET_GEOMETRY.get(wanted, "rect"),
            stem=AUTO_SHAPE_NAMES.get(wanted, "Rectangle"),
            box=(Left, Top, Width, Height),
        )
        shape.auto_shape_type = wanted
        return ShapeObject(self.slide, shape)

    @method
    def AddTextbox(
        self,
        Orientation: object = MISSING,
        Left: object = MISSING,
        Top: object = MISSING,
        Width: object = MISSING,
        Height: object = MISSING,
    ) -> object:
        del Orientation
        shape = self._made(
            kind="textBox", geometry="rect", stem="TextBox", box=(Left, Top, Width, Height)
        )
        return ShapeObject(self.slide, shape)

    @method
    def AddLine(
        self,
        BeginX: object = MISSING,
        BeginY: object = MISSING,
        EndX: object = MISSING,
        EndY: object = MISSING,
    ) -> object:
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
        return ShapeObject(self.slide, shape)

    @method
    def AddPicture(self, *args: object, **named: object) -> object:
        del args, named
        raise VBAUnsupportedError(
            "Shapes.AddPicture is real PowerPoint that pyOpenVBA does not implement"
        )

    @method
    def Range(self, Index: object = MISSING) -> object:
        del Index
        raise VBAUnsupportedError(
            "Shapes.Range is real PowerPoint that pyOpenVBA does not implement"
        )

    def _made(
        self, *, kind: str, geometry: str, stem: str, box: tuple[object, object, object, object]
    ) -> Shape:
        left, top, width, height = (_number(one) for one in box)
        shape = Shape(
            name=self._free_name(stem),
            kind=kind,
            geometry=geometry,
            left=left,
            top=top,
            width=width,
            height=height,
            shape_id=max((one.shape_id for one in self.slide.shapes_), default=1) + 1,
        )
        self.slide.shapes_.append(shape)
        self.slide.touched()
        return shape

    def _free_name(self, stem: str) -> str:
        """The name PowerPoint would give the next shape.

        One counter per slide, over every kind, which does not go back
        down when shapes are deleted; a new slide starts again at one.
        Measured, and the same rule as Excel's and Word's.
        """
        self.slide.shape_count += 1
        return f"{stem} {self.slide.shape_count}"


class ShapeObject(PowerPointObject):
    """One shape on a slide."""

    vba_type_name = "Shape"

    def __init__(self, slide: Slide, shape: Shape) -> None:
        self.slide = slide
        self.shape = shape

    @member
    def Name(self) -> object:
        return self.shape.name

    @setter("Name")
    def _set_name(self, value: object) -> None:
        self.shape.name = to_text(value)
        self.slide.touched()

    @member
    def Type(self) -> object:
        return VBAInt(self.shape.mso_type, "Long")

    @member
    def AutoShapeType(self) -> object:
        return VBAInt(self.shape.auto_shape_type, "Long")

    @member
    def ZOrderPosition(self) -> object:
        for index, one in enumerate(self.slide.shapes_, start=1):
            if one is self.shape:
                return VBAInt(index, "Long")
        return VBAInt(1, "Long")

    @member
    def HasTextFrame(self) -> object:
        return VBAInt(0 if self.shape.kind == "line" else -1, "Long")

    @member
    def Left(self) -> object:
        return VBASingle(self.shape.left)

    @setter("Left")
    def _set_left(self, value: object) -> None:
        self.shape.left = _number(value)
        self.slide.touched()

    @member
    def Top(self) -> object:
        return VBASingle(self.shape.top)

    @setter("Top")
    def _set_top(self, value: object) -> None:
        self.shape.top = _number(value)
        self.slide.touched()

    @member
    def Width(self) -> object:
        return VBASingle(self.shape.width)

    @setter("Width")
    def _set_width(self, value: object) -> None:
        self.shape.width = _number(value)
        self.slide.touched()

    @member
    def Height(self) -> object:
        return VBASingle(self.shape.height)

    @setter("Height")
    def _set_height(self, value: object) -> None:
        self.shape.height = _number(value)
        self.slide.touched()

    @member
    def TextFrame(self) -> object:
        return TextFrame(self)

    @member
    def TextFrame2(self) -> object:
        return TextFrame(self)

    @member
    def ActionSettings(self, Index: object = MISSING) -> object:
        settings = ActionSettings(self)
        return settings if Index is MISSING else settings.vba_get("Item", [Index])

    @member
    def GroupItems(self) -> object:
        if self.shape.kind != "group":
            raise error(ERR_NO_SUCH_SHAPE, "this shape is not a group")
        return GroupItems(self.slide, self.shape)

    @method
    def Delete(self) -> object:
        self.slide.shapes_ = [one for one in self.slide.shapes_ if one is not self.shape]
        self.slide.touched()
        return EMPTY

    @method
    def Select(self, Replace: object = MISSING) -> object:
        del Replace
        return EMPTY


class GroupItems(VBACollection, PowerPointObject):
    """The shapes inside a group."""

    vba_type_name = "GroupShapes"

    def __init__(self, slide: Slide, group: Shape) -> None:
        self.slide = slide
        self.group = group

    def vba_items(self) -> list[object]:
        return [ShapeObject(self.slide, one) for one in self.group.children]


class ActionSettings(VBACollection, PowerPointObject):
    """What a shape does when it is clicked or hovered over."""

    vba_type_name = "ActionSettings"

    def __init__(self, shape: ShapeObject) -> None:
        self.owner = shape

    def vba_items(self) -> list[object]:
        # ppMouseClick and ppMouseOver, in that order.
        return [ActionSetting(self.owner, 1), ActionSetting(self.owner, 2)]


class ActionSetting(PowerPointObject):
    """One of them: the click, or the hover.

    Only the click is modelled, because that is where the macro is;
    what PowerPoint holds for a hover is a separate hyperlink this does
    not write.
    """

    vba_type_name = "ActionSetting"

    def __init__(self, shape: ShapeObject, which: int) -> None:
        self.owner = shape
        self.which = which

    @member
    def Action(self) -> object:
        if self.which != 1:
            return VBAInt(0, "Long")
        shape = self.owner.shape
        return VBAInt(PP_ACTION_RUN_MACRO if (shape.action or shape.macro) else 0, "Long")

    @setter("Action")
    def _set_action(self, value: object) -> None:
        wanted = int(to_number(value))
        if wanted not in (0, PP_ACTION_RUN_MACRO):
            raise VBAUnsupportedError(
                f"ActionSetting.Action = {wanted} is real PowerPoint that pyOpenVBA "
                "does not implement; only ppActionRunMacro is"
            )
        if self.which != 1:
            return
        # PowerPoint remembers ppActionRunMacro before a macro is named,
        # so the action is held apart from the name.
        self.owner.shape.action = wanted
        if wanted == 0:
            self.owner.shape.macro = ""
        self.owner.slide.touched()

    @member
    def Run(self) -> object:
        return self.owner.shape.macro if self.which == 1 else ""

    @setter("Run")
    def _set_run(self, value: object) -> None:
        if self.which != 1:
            raise VBAUnsupportedError(
                "a macro on mouse-over is real PowerPoint that pyOpenVBA does not implement"
            )
        self.owner.shape.macro = to_text(value)
        self.owner.shape.action = PP_ACTION_RUN_MACRO if self.owner.shape.macro else 0
        self.owner.slide.touched()

    @member
    def Hyperlink(self) -> object:
        raise VBAUnsupportedError(
            "ActionSetting.Hyperlink is real PowerPoint that pyOpenVBA does not implement"
        )


class TextFrame(PowerPointObject):
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


class TextRange(PowerPointObject):
    """The text itself."""

    vba_type_name = "TextRange"

    def __init__(self, shape: ShapeObject) -> None:
        self.owner = shape

    @member(default=True)
    def Text(self) -> object:
        return self.owner.shape.text

    @setter("Text")
    def _set_text(self, value: object) -> None:
        self.owner.shape.text = to_text(value)
        self.owner.slide.touched()

    @member
    def Length(self) -> object:
        return VBAInt(len(self.owner.shape.text), "Long")


def _number(value: object) -> float:
    if value is MISSING or value is None:
        return 0.0
    return float(to_number(value))
