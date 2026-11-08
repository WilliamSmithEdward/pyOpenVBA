"""Worksheet.TextBoxes, Pictures, Rectangles, Ovals, Lines and Arcs, and DrawingObjects: the old names of the sheet's drawings
(scripts/measure_old_drawings.py, tests/fixtures/old_drawings.json).

Like the old control collections (see _old_controls), these are views of the sheet's shapes, not a model of their own: an item
answers through its shape, so what a macro sets one way reads back the other. What only reads here is what Excel answers for a new
drawing; setting those reports itself until what they write is measured.
"""
from __future__ import annotations

import re
from collections.abc import Sequence
from pathlib import Path
from typing import TYPE_CHECKING, ClassVar

from pyopenvba.apps.excel._held import HeldPart
from pyopenvba.apps.excel._old_controls import ListedControls, OldControls, OldObject, _object  # pyright: ignore[reportPrivateUsage]
from pyopenvba.apps.excel._styles import Color
from pyopenvba.apps.excel._shapes import ShapeObject, ShapeRange, TextFrame, _image_points  # pyright: ignore[reportPrivateUsage]
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBAObject, member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, VBAInt, error, to_bool, to_number, to_text
from pyopenvba.shapes._values import FillLine, Shape

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet

_NOT_THERE = (1004, 'Unable to get the property')
_XL_NONE = -4142
_XL_DASH = -4115
#: What each XlBorderWeight writes, in EMU.
_BORDER_WIDTHS = {1: 1, 2: 9525, -4138: 17145, 4: 24765}
_LINE_STYLES = {'lgDash': _XL_DASH}
#: MsoTextOrientation, which the shape's TextFrame takes, to XlOrientation, for those measured.
_ORIENTATIONS = {1: -4128, 2: -4171}
#: Interior and Border's ColorIndex while the style still gives the colour, by what the shape is.
_STYLE_INDEXES = {'shape': (14, 49), 'textBox': (-4105, 15), 'line': (_XL_NONE, 14), 'arc': (_XL_NONE, 14), 'picture': (_XL_NONE, _XL_NONE)}


def _style_of(shape: Shape) -> str:
    return 'arc' if shape.geometry == 'arc' else shape.kind


def _colour_index(drawing: OldDrawing, colour: int) -> VBAInt:
    """The palette's place of a colour set with RGB. A colour not in the palette reads as the nearest, by a rule not worked out."""
    wanted = f'{colour & 255:02X}{colour >> 8 & 255:02X}{colour >> 16 & 255:02X}'
    palette = list(drawing.sheet.book.stylesheet.colors.indexed[8:64])
    if wanted not in palette:
        raise VBAUnsupportedError("The ColorIndex of a drawing's colour outside the palette is not implemented")
    return VBAInt(palette.index(wanted) + 1, 'Long')


class _Paint(HeldPart):
    """A drawing's Interior or Border, in terms of its shape's Fill or Line."""

    vba_type_name = ''
    slot = 0

    def __init__(self, owner: OldDrawing) -> None:
        self.owner = owner

    def _format(self) -> VBAObject:
        return _object(self.owner.drawn.vba_get('Line' if self.slot else 'Fill', []))

    def _look(self) -> FillLine:
        """What the shape holds of its own, read from its markup if it came from a file."""
        return getattr(self._format(), 'look')()

    def _refuse_picture(self) -> None:
        if self.owner.shape.kind == 'picture':
            raise VBAUnsupportedError("Setting the fill or line of a picture through its old object is not implemented")

    def _bare_picture(self) -> bool:
        """A picture with no fill or line of its own, as this part goes."""
        look = self._look()
        own = (look.line_rgb or look.line_scheme or look.line_width) if self.slot else (look.fill_rgb or look.fill_scheme)
        return self.owner.shape.kind == 'picture' and not own

    def _shown(self) -> bool:
        return not self._bare_picture() and bool(to_number(self._format().vba_get('Visible', [])))

    def _set_colour(self, name: str, value: object) -> None:
        self._refuse_picture()
        _object(self._format().vba_get('ForeColor', [])).vba_set(name, value)

    @member
    def Parent(self) -> object:
        return self.owner

    @member
    def Color(self) -> object:
        if self._bare_picture():
            return VBAInt(16777215, 'Long')
        look = self._look()
        if not self.slot and look.fill_rgb is None and look.fill_scheme is None:
            # The style's fill, or white where it has none, whether or not it is hidden.
            return VBAInt({'shape': 8544277}.get(_style_of(self.owner.shape), 16777215), 'Long')
        return _object(self._format().vba_get('ForeColor', [])).vba_get('RGB', [])

    @setter('Color')
    def _color(self, value: object) -> None:
        self._set_colour('RGB', value)

    @member
    def ColorIndex(self) -> object:
        look = self._look()
        own = look.line_rgb if self.slot else look.fill_rgb
        if self._bare_picture() or (own is None and not self._shown()):
            return VBAInt(_XL_NONE, 'Long')
        colour = int(to_number(self.vba_get('Color', [])))
        if self.owner.shape.kind == 'picture':
            # Its colour reads as the nearest in the palette (tests/fixtures/shape_settings_kept.json).
            colors = self.owner.sheet.book.stylesheet.colors
            return VBAInt(colors.tab_index(Color('rgb', f'FF{colour & 255:02X}{colour >> 8 & 255:02X}{colour >> 16 & 255:02X}')), 'Long')
        scheme = look.line_scheme if self.slot else look.fill_scheme
        if own is None and scheme and colour != getattr(self._format(), 'style')()[2 if self.slot else 0]:
            raise VBAUnsupportedError("The ColorIndex of a drawing's theme colour is not implemented")
        if own is None:
            return VBAInt(_STYLE_INDEXES[_style_of(self.owner.shape)][self.slot], 'Long')
        return _colour_index(self.owner, int(to_number(self.vba_get('Color', []))))

    @setter('ColorIndex')
    def _color_index(self, value: object) -> None:
        wanted = int(to_number(value))
        if wanted == _XL_NONE:
            self._refuse_picture()
            self._format().vba_set('Visible', VBAInt(0, 'Long'))
            return
        if not 1 <= wanted <= 56:
            raise VBAUnsupportedError("This ColorIndex of a drawing's old object is not implemented")
        # A SchemeColor counts the palette from its first eight colours, which ColorIndex leaves out.
        self._set_colour('SchemeColor', VBAInt(wanted + 7, 'Long'))


class Interior(_Paint):
    vba_type_name = 'Interior'

    @member
    def Pattern(self) -> object:
        # Hiding the fill leaves it: only a style with no fill reads none.
        look = self._look()
        given = look.fill_rgb is not None or look.fill_scheme is not None
        return VBAInt(1 if given or _style_of(self.owner.shape) in ('shape', 'textBox') else _XL_NONE, 'Long')

    @setter('Pattern')
    def _pattern(self, value: object) -> None:
        if int(to_number(value)) != _XL_NONE:
            raise VBAUnsupportedError("This Pattern of a drawing's old object is not implemented")
        self._refuse_picture()
        self._format().vba_set('Visible', VBAInt(0, 'Long'))


class Border(_Paint):
    vba_type_name = 'Border'
    slot = 1

    def vba_get(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        # Of a shape that has gone, these two read as nothing, and the rest fail.
        if name.lower() in ('linestyle', 'weight') and self.held_gone():
            return ''
        return super().vba_get(name, args, named)

    @member
    def LineStyle(self) -> object:
        if not self._shown():
            return VBAInt(_XL_NONE, 'Long')
        dash = self._look().line_dash
        if dash is not None and dash not in _LINE_STYLES:
            raise VBAUnsupportedError("This LineStyle of a drawing's old object is not implemented")
        return VBAInt(1 if dash is None else _LINE_STYLES[dash], 'Long')

    @setter('LineStyle')
    def _line_style(self, value: object) -> None:
        wanted = int(to_number(value))
        self._refuse_picture()
        if wanted == _XL_NONE:
            self._format().vba_set('Visible', VBAInt(0, 'Long'))
        elif wanted == _XL_DASH:
            getattr(self._format(), 'change')().line_dash = 'lgDash'
        else:
            raise VBAUnsupportedError("This LineStyle of a drawing's old object is not implemented")

    @member
    def Weight(self) -> object:
        if self._bare_picture():
            return VBAInt(1, 'Long')
        width = float(to_number(self._format().vba_get('Weight', [])))
        if width <= 0:
            # No line shows.
            return VBAInt(2 if self.owner.shape.kind == 'textBox' else -4138, 'Long')
        # By the line's width, with the widths between these steps not measured (tests/fixtures/shape_settings_kept.json).
        if width < 0.001:
            return VBAInt(1, 'Long')
        if 0.25 <= width <= 1:
            return VBAInt(2, 'Long')
        if 1.25 <= width <= 1.5:
            return VBAInt(-4138, 'Long')
        if width >= 1.95:
            return VBAInt(4, 'Long')
        raise VBAUnsupportedError(f"The Weight of a drawing's line {width} points wide is not implemented")

    @setter('Weight')
    def _weight(self, value: object) -> None:
        width = _BORDER_WIDTHS.get(int(to_number(value)))
        if width is None:
            raise VBAUnsupportedError("This Weight of a drawing's old object is not implemented")
        self._refuse_picture()
        getattr(self._format(), 'change')().line_width = width


class OldDrawing(OldObject):
    """A drawing under its old name."""

    vba_type_name = ''

    def holds(self, shape: Shape) -> bool:
        return _kind_of(shape) is type(self)

    @member
    def Enabled(self) -> object:
        raise error(*_NOT_THERE)

    @setter('Enabled')
    def _enabled(self, value: object) -> None:
        raise error(*_NOT_THERE)

    @member
    def PrintObject(self) -> object:
        return self.shape.prints

    @setter('PrintObject')
    def _print_object(self, value: object) -> None:
        self.shape.prints = to_bool(value)
        self.sheet.drawing_changed()

    @member
    def ShapeRange(self) -> object:
        return ShapeRange([self.drawn])

    @member
    def Border(self) -> object:
        return Border(self)


class _Filled(OldDrawing):
    """A drawing with an inside to fill."""

    vba_type_name = ''

    @member
    def Interior(self) -> object:
        return Interior(self)


class _Texted(_Filled):
    """A drawing that holds text, through its shape's TextFrame."""

    vba_type_name = ''

    def _frame(self) -> TextFrame:
        return TextFrame(self.drawn)

    def _whole(self) -> VBAObject:
        return _object(self._frame().vba_get('Characters', []))

    @member
    def Text(self) -> object:
        return self._whole().vba_get('Text', [])

    @setter('Text')
    def _text(self, value: object) -> None:
        self._whole().vba_set('Text', value)

    @member
    def Caption(self) -> object:
        return self._whole().vba_get('Text', [])

    @setter('Caption')
    def _caption(self, value: object) -> None:
        self._whole().vba_set('Text', value)

    @method
    def Characters(self, Start: object = MISSING, Length: object = MISSING) -> object:  # noqa: N803 - native named arguments
        return self._frame().vba_get('Characters', [Start, Length])

    @member
    def Font(self) -> object:
        return self._whole().vba_get('Font', [])

    @member
    def HorizontalAlignment(self) -> object:
        return self._frame().vba_get('HorizontalAlignment', [])

    @setter('HorizontalAlignment')
    def _horizontal_alignment(self, value: object) -> None:
        self._frame().vba_set('HorizontalAlignment', value)

    @member
    def VerticalAlignment(self) -> object:
        return self._frame().vba_get('VerticalAlignment', [])

    @setter('VerticalAlignment')
    def _vertical_alignment(self, value: object) -> None:
        self._frame().vba_set('VerticalAlignment', value)

    @member
    def Orientation(self) -> object:
        wanted = int(to_number(self._frame().vba_get('Orientation', [])))
        if wanted not in _ORIENTATIONS:
            raise VBAUnsupportedError("This Orientation of a drawing's old object is not implemented")
        return VBAInt(_ORIENTATIONS[wanted], 'Long')

    @setter('Orientation')
    def _orientation(self, value: object) -> None:
        wanted = int(to_number(value))
        if wanted not in _ORIENTATIONS.values():
            raise VBAUnsupportedError("This Orientation of a drawing's old object is not implemented")
        self._frame().vba_set('Orientation', VBAInt(next(key for key, one in _ORIENTATIONS.items() if one == wanted), 'Long'))

    @member
    def AutoSize(self) -> object:
        # Only a text box, or a shape with text, has one to read.
        if self.shape.kind != 'textBox' and not self.shape.text:
            raise error(*_NOT_THERE)
        return False

    @setter('AutoSize')
    def _auto_size(self, value: object) -> None:
        if self.shape.kind != 'textBox':
            raise error(*_NOT_THERE)
        if to_bool(value):
            raise VBAUnsupportedError('Fitting a text box to its text is not implemented')

    @member
    def AddIndent(self) -> object:
        return False

    @member
    def LockedText(self) -> object:
        return True

    @setter('LockedText')
    def _locked_text(self, value: object) -> None:
        raise VBAUnsupportedError("Setting LockedText on a drawing's old object is not implemented")

    @member
    def ReadingOrder(self) -> object:
        return VBAInt(-2, 'Long')

    @member
    def Formula(self) -> object:
        # From a file Excel reads the link back with a space after it (tests/fixtures/old_collection_members.json).
        found = re.search(r'\btextlink="([^"]+)"', self.shape.source)
        return f'{found.group(1)} ' if found else ''

    @setter('Formula')
    def _formula(self, value: object) -> None:
        raise VBAUnsupportedError("Linking a drawing's text to a cell is not implemented")


class _Shadowed(HeldPart):
    """Shadow as a Boolean, on those that have it."""

    vba_type_name = ''

    @member
    def Shadow(self) -> object:
        return bool(to_number(_object(getattr(self, 'drawn').vba_get('Shadow', [])).vba_get('Visible', [])))

    @setter('Shadow')
    def _shadow(self, value: object) -> None:
        if not to_bool(value):
            raise VBAUnsupportedError("Taking the Shadow off a drawing's old object is not implemented")
        _object(getattr(self, 'drawn').vba_get('Shadow', [])).vba_set('Visible', VBAInt(-1, 'Long'))


class _Cornered(HeldPart):
    """RoundedCorners, which only reads here: setting it renames a rectangle, by a rule measured once."""

    vba_type_name = ''

    @member
    def RoundedCorners(self) -> object:
        return getattr(self, 'shape').geometry == 'roundRect'

    @setter('RoundedCorners')
    def _rounded_corners(self, value: object) -> None:
        raise VBAUnsupportedError("Setting RoundedCorners on a drawing's old object is not implemented")


class TextBox(_Shadowed, _Cornered, _Texted):
    vba_type_name = 'TextBox'


class Rectangle(_Shadowed, _Cornered, _Texted):
    vba_type_name = 'Rectangle'

    def holds(self, shape: Shape) -> bool:
        return _is_rectangle(shape)


class Oval(_Shadowed, _Texted):
    vba_type_name = 'Oval'


class Arc(_Texted):
    vba_type_name = 'Arc'


class Line(OldDrawing):
    vba_type_name = 'Line'

    def holds(self, shape: Shape) -> bool:
        return _is_collected_line(shape)

    def vba_get(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        # An elbow or a curve has an Interior, as none shows (tests/fixtures/shape_settings_kept.json); a straight line has none.
        if name.lower() == 'interior' and self.shape.geometry not in ('line', 'straightConnector1'):
            from pyopenvba.apps.excel._text_paragraphs import NoEffect

            return NoEffect(self, 'Interior', {'Color': VBAInt(0, 'Long'), 'ColorIndex': VBAInt(_XL_NONE, 'Long'), 'Pattern': VBAInt(_XL_NONE, 'Long')})
        return super().vba_get(name, args, named)

    @member
    def ArrowHeadStyle(self) -> object:
        return VBAInt(_XL_NONE, 'Long')

    @setter('ArrowHeadStyle')
    def _arrow_head_style(self, value: object) -> None:
        raise VBAUnsupportedError("Setting a Line's arrowhead through its old object is not implemented")

    @member
    def ArrowHeadLength(self) -> object:
        return VBAInt(-4138, 'Long')

    @setter('ArrowHeadLength')
    def _arrow_head_length(self, value: object) -> None:
        raise VBAUnsupportedError("Setting a Line's arrowhead through its old object is not implemented")

    @member
    def ArrowHeadWidth(self) -> object:
        return VBAInt(-4138, 'Long')

    @setter('ArrowHeadWidth')
    def _arrow_head_width(self, value: object) -> None:
        raise VBAUnsupportedError("Setting a Line's arrowhead through its old object is not implemented")


class Picture(_Filled):
    vba_type_name = 'Picture'

    @member
    def Shadow(self) -> object:
        return False

    @setter('Shadow')
    def _shadow(self, value: object) -> None:
        raise VBAUnsupportedError("Setting a picture's Shadow through its old object is not implemented")

    @member
    def Formula(self) -> object:
        return ''

    @setter('Formula')
    def _formula(self, value: object) -> None:
        raise VBAUnsupportedError('Linking a picture to cells is not implemented')


def _is_rectangle(shape: Shape) -> bool:
    return shape.kind == 'shape' and shape.geometry in ('rect', 'roundRect')


def _is_collected_line(shape: Shape) -> bool:
    return shape.kind == 'line' and shape.geometry in ('line', 'straightConnector1')


def _kind_of(shape: Shape) -> type[OldDrawing] | None:
    """The old object a shape answers as. Any other AutoShape answers as a Rectangle, though Rectangles does not hold it."""
    if shape.kind == 'textBox':
        return TextBox
    if shape.kind == 'picture':
        return Picture
    if shape.kind == 'line':
        return Line
    if shape.kind != 'shape':
        return None
    return {'ellipse': Oval, 'arc': Arc}.get(shape.geometry, Rectangle)


def drawing_of(sheet: Worksheet, shape: Shape) -> OldDrawing:
    kind = _kind_of(shape)
    if kind is None:
        raise VBAUnsupportedError(f'The old object of a {shape.kind} is not implemented')
    return kind(sheet, shape)


class _Drawings(OldControls):
    """One kind of a sheet's drawings. Reading a member of its items from several fails, though setting one reaches them all."""

    vba_type_name = ''
    item: ClassVar[type[OldDrawing]]
    parts = frozenset({'Border', 'Font', 'Interior'})
    spread_alone = False
    several = {'addindent': False, 'autoscalefont': False}
    doubled = frozenset()

    def holds(self, shape: Shape) -> bool:
        return self.item.holds(self.item.__new__(self.item), shape)

    def item_for(self, shape: Shape) -> object:
        return drawing_of(self.sheet, shape)

    def _added(self, stem: str, *arguments: object) -> object:
        found = _object(self.sheet.vba_get('Shapes', [])).vba_call(stem, list(arguments))
        return drawing_of(self.sheet, getattr(found, 'shape'))

    def nothing_held(self, name: str) -> object:
        raise error(*_NOT_THERE)

    def read_part(self, part: str, name: str) -> object:
        if part == 'Font' and name.lower() == 'colorindex':
            raise error(-2147467259, 'Automation error')
        if part in ('Font', 'Characters'):
            raise error(*_NOT_THERE)
        # An Interior's Pattern fails with its colours, whatever the patterns are.
        asked = ['ColorIndex', name] if part == 'Interior' and name.lower() == 'pattern' else [name]
        answers: list[object] = []
        for each in asked:
            answers = [_object(_object(one).vba_get(part, [])).vba_get(each, []) for one in self.vba_items()]
            if len({str(one) for one in answers}) > 1:
                raise error(-2147467259, 'Automation error')
        return answers[0]


class _BoxAdded(_Drawings):
    """Those whose Add takes a box."""

    vba_type_name = ''
    missing = 450
    shape_type = 1

    @method
    def Add(self, Left: object = MISSING, Top: object = MISSING, Width: object = MISSING, Height: object = MISSING) -> object:  # noqa: N803 - native named arguments
        if MISSING in (Left, Top, Width, Height):
            raise error(self.missing, 'Argument not optional' if self.missing == 449 else 'Wrong number of arguments or invalid property assignment')
        if float(to_number(Width)) < 0 or float(to_number(Height)) < 0:
            raise error(-2147024809, 'The parameter is incorrect.')
        return self._added('AddShape', VBAInt(self.shape_type, 'Long'), Left, Top, Width, Height)


class TextBoxes(_BoxAdded):
    spread_reads = frozenset({'Border', 'Caption', 'Characters', 'Enabled', 'Font', 'Interior', 'Locked', 'Placement', 'PrintObject', 'ShapeRange', 'Text', 'Visible', 'ZOrder'})
    spread_sets = frozenset({'Caption', 'Enabled', 'Locked', 'Placement', 'PrintObject', 'Text', 'Visible'})
    vba_type_name = 'TextBoxes'
    item = TextBox
    missing = 449

    @method
    def Add(self, Left: object = MISSING, Top: object = MISSING, Width: object = MISSING, Height: object = MISSING) -> object:  # noqa: N803 - native named arguments
        if MISSING in (Left, Top, Width, Height):
            raise error(449, 'Argument not optional')
        return self._added('AddTextbox', VBAInt(1, 'Long'), Left, Top, Width, Height)


class Rectangles(_BoxAdded):
    spread_reads = frozenset({'AutoSize', 'Border', 'Caption', 'Characters', 'Enabled', 'Font', 'HorizontalAlignment', 'Interior', 'Locked', 'Orientation', 'Placement', 'PrintObject', 'ShapeRange', 'Text', 'VerticalAlignment', 'Visible', 'ZOrder'})
    spread_sets = frozenset({'Caption', 'Enabled', 'HorizontalAlignment', 'Locked', 'Orientation', 'Placement', 'PrintObject', 'Text', 'VerticalAlignment', 'Visible'})
    vba_type_name = 'Rectangles'
    parts = frozenset({'Border', 'Characters', 'Font', 'Interior'})
    several = {'addindent': False, 'autoscalefont': False, 'readingorder': VBAInt(27, 'Long')}
    item = Rectangle


class Ovals(_BoxAdded):
    spread_reads = frozenset({'AutoSize', 'Border', 'Caption', 'Characters', 'Enabled', 'Font', 'HorizontalAlignment', 'Interior', 'Locked', 'Orientation', 'Placement', 'PrintObject', 'ShapeRange', 'Text', 'VerticalAlignment', 'Visible', 'ZOrder'})
    spread_sets = frozenset({'Caption', 'Enabled', 'HorizontalAlignment', 'Locked', 'Orientation', 'Placement', 'PrintObject', 'Text', 'VerticalAlignment', 'Visible'})
    vba_type_name = 'Ovals'
    item = Oval
    shape_type = 9


class Arcs(_Drawings):
    spread_reads = frozenset({'AutoSize', 'Border', 'Caption', 'Characters', 'Enabled', 'Font', 'HorizontalAlignment', 'Interior', 'Locked', 'Orientation', 'Placement', 'PrintObject', 'ShapeRange', 'Text', 'VerticalAlignment', 'Visible', 'ZOrder'})
    spread_sets = frozenset({'AutoSize', 'Caption', 'Enabled', 'HorizontalAlignment', 'Locked', 'Orientation', 'Placement', 'PrintObject', 'Text', 'VerticalAlignment', 'Visible'})
    vba_type_name = 'Arcs'
    item = Arc

    @method
    def Add(self, X1: object = MISSING, Y1: object = MISSING, X2: object = MISSING, Y2: object = MISSING) -> object:  # noqa: N803 - native named arguments
        if MISSING in (X1, Y1, X2, Y2):
            raise error(450, 'Wrong number of arguments or invalid property assignment')
        return self._added('AddShape', VBAInt(25, 'Long'), X1, Y1, X2, Y2)


class Lines(_Drawings):
    spread_reads = frozenset({'Border', 'Enabled', 'Locked', 'Placement', 'PrintObject', 'ShapeRange', 'Visible', 'ZOrder'})
    spread_sets = frozenset({'Enabled', 'Locked', 'Placement', 'PrintObject', 'Visible'})
    vba_type_name = 'Lines'
    parts = frozenset({'Border'})
    several = {}
    item = Line

    @method
    def Add(self, X1: object = MISSING, Y1: object = MISSING, X2: object = MISSING, Y2: object = MISSING) -> object:  # noqa: N803 - native named arguments
        if MISSING in (X1, Y1, X2, Y2):
            raise error(450, 'Wrong number of arguments or invalid property assignment')
        return self._added('AddLine', X1, Y1, X2, Y2)


class Pictures(_Drawings):
    spread_reads = frozenset({'Border', 'Enabled', 'Interior', 'Locked', 'Placement', 'PrintObject', 'ShapeRange', 'Visible', 'ZOrder'})
    spread_sets = frozenset({'Enabled', 'Locked', 'Placement', 'PrintObject', 'Visible'})
    vba_type_name = 'Pictures'
    parts = frozenset({'Border', 'Interior'})
    several = {}
    item = Picture

    @method
    def Insert(self, Filename: object = MISSING, Converter: object = MISSING) -> object:  # noqa: N803 - native named arguments
        """A picture linked to its file, not kept in the workbook, at its own size at the active cell."""
        from pyopenvba.apps.excel._protection import guard_drawing

        del Converter
        self.sheet.book.application.outside('Pictures.Insert')
        guard_drawing(self.sheet, 'Pictures.Insert')
        if Filename is MISSING:
            raise error(449, 'Argument not optional')
        source = Path(to_text(Filename))
        try:
            data = source.read_bytes()
        except OSError:
            raise error(1004, 'Unable to get the Insert property of the Pictures class') from None
        wide, tall, _ = _image_points(data)
        if self.sheet.book.active_sheet is not self.sheet:
            raise VBAUnsupportedError('Pictures.Insert on a sheet that is not active is not implemented')
        cell = _object(self.sheet.active_cell)
        shapes = _object(self.sheet.vba_get('Shapes', []))
        self.sheet.shape_count += 1
        shape = Shape(name=getattr(shapes, '_free_name')('Picture'), kind='picture', left=float(to_number(cell.vba_get('Left', []))), top=float(to_number(cell.vba_get('Top', []))),
                      width=wide, height=tall, shape_id=getattr(shapes, '_free_id')() + 1, edit_as='oneCell', lock_aspect=True, image_link=str(source))
        self.sheet.shapes_.append(shape)
        self.sheet.drawing_changed()
        return drawing_of(self.sheet, shape)


class DrawingObjects(_Drawings, ListedControls):
    parts = frozenset()
    several = {}
    spread_reads = frozenset({'Accelerator', 'Border', 'CancelButton', 'Caption', 'Characters', 'DefaultButton', 'DismissButton', 'Display3DShading', 'DisplayVerticalScrollBar', 'DropDownLines', 'Enabled', 'Font', 'HelpButton', 'InputType', 'Interior', 'LargeChange', 'LinkedCell', 'ListFillRange', 'ListIndex', 'Locked', 'Max', 'Min', 'MultiLine', 'MultiSelect', 'PhoneticAccelerator', 'Placement', 'PrintObject', 'ShapeRange', 'SmallChange', 'Text', 'Value', 'ZOrder'})

    def vba_get(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        """A member of the collection; of its items' members, the few Excel answers over a sheet's every kind, whatever they hold, and 1004 for the rest
        (tests/fixtures/old_collection_reads.json). Visible is not followed yet."""
        lower = name.lower()
        if lower in {key.lower() for key in self._vba_members} and lower not in {key.lower() for key in self.spread_reads}:
            return super().vba_get(name, args, named)
        if lower in ('border', 'characters', 'font', 'interior'):
            raise error(438, "Object doesn't support this property or method")
        if lower == 'shaperange':
            raise error(450, 'Wrong number of arguments or invalid property assignment')
        if lower == 'readingorder' and self.vba_items():
            return VBAInt(27, 'Long')
        if lower == 'visible':
            raise VBAUnsupportedError('DrawingObjects.Visible, which follows no one item, is not implemented')
        raise error(*_NOT_THERE)

    """All of a sheet's drawings and controls, each as its old object."""

    vba_type_name = 'DrawingObjects'

    def holds(self, shape: Shape) -> bool:
        return True

    def item_for(self, shape: Shape) -> object:
        from pyopenvba.apps.excel._old_controls import item_of

        if shape.kind in ('shape', 'textBox', 'line', 'picture', 'formControl'):
            return item_of(self.sheet, shape)
        # A chart answers as its ChartObject. A group's old object is not built, and it answers as the shape.
        return _object(self.sheet.vba_get('Shapes', [])).vba_get('Item', [shape.name])

    @method
    def Delete(self) -> object:
        for shape in list(self._shapes()):
            ShapeObject(self.sheet, shape).vba_call('Delete', [])
        return EMPTY

    # Its items' list members, called on all of a sheet's drawings, are 1004 as nearly every member read over them is
    # (tests/fixtures/old_collection_methods.json).
    @method
    def List(self, Index: object = MISSING) -> object:  # noqa: N803 - native named argument
        raise error(1004, 'Unable to get the List property')

    @method
    def Selected(self, Index: object = MISSING) -> object:  # noqa: N803 - native named argument
        raise error(1004, 'Unable to get the Selected property')

    @method
    def LinkCombo(self, Link: object = MISSING) -> object:  # noqa: N803 - native named argument
        raise error(1004, 'Unable to get the LinkCombo property')


__all__ = ['Arc', 'Arcs', 'DrawingObjects', 'Line', 'Lines', 'Oval', 'Ovals', 'Picture', 'Pictures', 'Rectangle', 'Rectangles', 'TextBox', 'TextBoxes', 'drawing_of']
