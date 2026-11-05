"""The Border of a chart's elements, the older way to reach their line, over the same look as Format.Line
(tests/fixtures/chart_borders.json). Elements fall into three families, which read and write differently: those with a
grey line of their own, a series or point with none, and the boxes around text and the plot with none. Trendlines and
error bars have a black line and a family each."""
from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Protocol, cast

from pyopenvba.apps.excel._creator import Created
from pyopenvba.interpreter._objects import VBAObject, member, method, setter
from pyopenvba.interpreter._values import VBAInt, error, to_integer, to_number
from pyopenvba.shapes._values import FillLine
from pyopenvba.exceptions import VBAUnsupportedError

if TYPE_CHECKING:
    from pyopenvba.apps.excel._styles import Colors

_AUTOMATIC, _NONE, _CONTINUOUS = -4105, -4142, 1
_MEDIUM = -4138
#: Border.Weight and the width it writes.
_WIDTHS = {1: 3175, 2: 12700, _MEDIUM: 25400, 4: 38100}
#: Border.LineStyle and the dash it writes.
_DASHES = {_CONTINUOUS: 'solid', -4115: 'lgDash', -4118: 'sysDash', 4: 'lgDashDot', 5: 'lgDashDotDot'}


@dataclass(frozen=True)
class Family:
    #: What Color reads with nothing set.
    colour: int
    #: The width a hidden line writes, and the Weight it reads.
    hidden_width: int
    hidden_weight: int
    #: Whether Weight or LineStyle on a line with nothing set makes a black line, and of what width for LineStyle.
    draws: bool
    #: Whether Weight on a line with nothing set is written at all.
    weighs: bool
    #: What Weight and LineStyle read with nothing set.
    weight: int = 2
    style: int = _AUTOMATIC
    #: Whether LineStyle gives a line with nothing set a width.
    styles_width: bool = True
    #: Whether a line drawn by Weight still reads as automatic.
    weight_reads_automatic: bool = False
    #: Whether the line can be hidden.
    hides: bool = True


GREY = Family(9013641, 12700, 2, False, True)
SERIES = Family(-2, 25400, 2, True, True)
BOX = Family(-2, 25400, _MEDIUM, False, False)
TRENDLINE = Family(0, 12700, 2, True, True, styles_width=False, weight_reads_automatic=True)
ERROR_BARS = Family(0, 12700, 2, False, True, hides=False)
#: A chart group's lines and bars (tests/fixtures/chart_group_looks.json).
GROUP_LINES = Family(0, 12700, 2, False, True)


def family_of_series(kind: str, accent: int) -> Family:
    """A series' family: a line chart's draws a line of its accent, and a scatter chart's none."""
    if kind in ('line', 'lineMarkers'):
        return Family(accent, 38100, 2, True, True, 4)
    if kind == 'scatter':
        return Family(-2, 38100, 2, True, True, 4, _NONE)
    return SERIES


def family_of(key: str) -> Family:
    return GREY if key == 'area' or key.startswith(('axis', 'grid')) else BOX


class Line(Protocol):
    """What a Format.Line offers: the look as it stands, and the look to change."""

    def look(self) -> FillLine: ...

    def change(self) -> FillLine: ...


def _rgb(colour: int) -> str:
    return f'{colour & 0xFF:02X}{colour >> 8 & 0xFF:02X}{colour >> 16 & 0xFF:02X}'


def _weight_of(width: int) -> int:
    return 1 if width <= 6350 else 2 if width <= 19050 else _MEDIUM if width <= 31750 else 4


def _clear(look: FillLine) -> None:
    look.line_rgb, look.line_width, look.line_dash, look.line_cap = None, None, None, None
    look.line_compound = None
    look.line_off = look.line_by_weight = False


#: What GetProperty and SetProperty raise on a chart element, by the element (tests/fixtures/chart_group_looks.json).
NOT_A_PROPERTY = -2147220480
UNSPECIFIED = -2147467259
INVALID_CALL = 445


class ElementProperties(Created):
    """GetProperty and SetProperty, which take no name on these elements, those in the file's own terms included."""

    #: A base, not a type of its own.
    vba_type_name = ''
    get_error = NOT_A_PROPERTY
    set_error = NOT_A_PROPERTY

    @method
    def GetProperty(self, Name: object) -> object:  # noqa: N803 - native named argument
        del Name
        raise error(self.get_error, 'Automation error')

    @method
    def SetProperty(self, Name: object, Value: object) -> object:  # noqa: N803 - native named argument
        del Name, Value
        raise error(self.set_error, 'Automation error')


class PictureMembers(Created):
    """What a series or point reads of a picture fill it does not have; none of it can be set
    (tests/fixtures/chart_members_more.json)."""

    #: A base, not a type of its own.
    vba_type_name = ''

    @member
    def ApplyPictToEnd(self) -> object:
        raise error(UNSPECIFIED, 'Automation error')

    @setter('ApplyPictToEnd')
    def _apply_pict_to_end(self, value: object) -> None:
        del value
        raise error(UNSPECIFIED, 'Automation error')

    @member
    def ApplyPictToFront(self) -> object:
        raise error(UNSPECIFIED, 'Automation error')

    @setter('ApplyPictToFront')
    def _apply_pict_to_front(self, value: object) -> None:
        del value
        raise error(UNSPECIFIED, 'Automation error')

    @member
    def ApplyPictToSides(self) -> object:
        raise error(UNSPECIFIED, 'Automation error')

    @setter('ApplyPictToSides')
    def _apply_pict_to_sides(self, value: object) -> None:
        del value
        raise error(UNSPECIFIED, 'Automation error')

    @member
    def PictureType(self) -> object:
        return VBAInt(1, 'Long')

    @setter('PictureType')
    def _picture_type(self, value: object) -> None:
        del value
        raise error(UNSPECIFIED, 'Automation error')

    @member
    def PictureUnit2(self) -> object:
        return 0.0

    @setter('PictureUnit2')
    def _picture_unit2(self, value: object) -> None:
        del value
        raise error(UNSPECIFIED, 'Automation error')


class ChartBorder(Created):
    vba_type_name = 'Border'

    def __init__(self, line: Line, family: Family, parent: object, colours: Colors) -> None:
        self.line, self.family, self.parent, self.colours = line, family, parent, colours

    def look(self) -> FillLine:
        """The line's look. A loaded chart with a style part answers Border otherwise than its Format, which is not measured, so
        it refuses (tests/fixtures/loaded_chart_looks.json)."""
        chart = getattr(self.parent, 'chart', None) or getattr(getattr(self.parent, 'axis', None), 'chart', None)
        state = getattr(chart, 'state', None)
        if state is not None and state.original and state.style_part:
            raise VBAUnsupportedError('The Border of an element of an imported styled chart is not implemented')
        return self.line.look()

    @member
    def Parent(self) -> object:
        return self.parent

    @member
    def Application(self) -> object:
        return cast(VBAObject, self.parent).vba_get('Application')

    @member
    def Color(self) -> object:
        look = self.look()
        if look.line_off or look.line_hidden:
            return VBAInt(-2, 'Long')
        if look.line_rgb is not None:
            text = look.line_rgb
            return VBAInt(int(text[4:6] + text[2:4] + text[:2], 16), 'Long')
        return VBAInt(self.family.colour, 'Long')

    @setter('Color')
    def _color(self, value: object) -> None:
        self._paint(_rgb(int(to_number(value))))

    def _paint(self, rgb: str) -> None:
        self.look()
        look = self.line.change()
        look.line_off = look.line_by_weight = False
        look.line_rgb = rgb
        look.line_dash = look.line_dash or 'solid'

    @member
    def ColorIndex(self) -> object:
        look = self.look()
        if look.line_off or look.line_hidden:
            return VBAInt(_NONE, 'Long')
        if look.line_rgb is None or look.line_by_weight:
            return VBAInt(_AUTOMATIC, 'Long')
        return VBAInt(self.colours.nearest(look.line_rgb), 'Long')

    @setter('ColorIndex')
    def _color_index(self, value: object) -> None:
        wanted = int(to_integer(value, 'Long'))
        if wanted == _AUTOMATIC:
            self._automatic()
        elif wanted == _NONE:
            if not self.family.hides:
                raise error(-2147467259, 'Automation error\nUnspecified error ')
            self._hide()
        elif 1 <= wanted <= 56:
            self._paint(self.colours.indexed[wanted + 7])
        else:
            raise error(1004, 'Unable to set the ColorIndex property of the Border class')

    @member
    def Weight(self) -> object:
        look = self.look()
        if look.line_off or look.line_hidden:
            return VBAInt(self.family.hidden_weight, 'Long')
        return VBAInt(self.family.weight if look.line_width is None else _weight_of(look.line_width), 'Long')

    @setter('Weight')
    def _weight(self, value: object) -> None:
        wanted = int(to_integer(value, 'Long'))
        if wanted not in _WIDTHS:
            raise error(1004, 'Unable to set the Weight property of the Border class')
        before = self.look()
        untouched = before.line_rgb is None and before.line_dash is None and not before.line_off
        if untouched and not self.family.weighs:
            return
        look = self.line.change()
        look.line_off = False
        look.line_width = _WIDTHS[wanted]
        if untouched and self.family.draws:
            look.line_rgb, look.line_dash = '000000', 'solid'
            look.line_by_weight = self.family.weight_reads_automatic

    @member
    def LineStyle(self) -> object:
        look = self.look()
        if look.line_off or look.line_hidden:
            return VBAInt(_NONE, 'Long')
        if look.line_by_weight or (look.line_rgb is None and look.line_dash is not None):
            return VBAInt(_AUTOMATIC, 'Long')
        if look.line_dash is not None:
            return VBAInt(next((key for key, name in _DASHES.items() if name == look.line_dash), _CONTINUOUS), 'Long')
        return VBAInt(_CONTINUOUS if look.line_rgb is not None else self.family.style, 'Long')

    @setter('LineStyle')
    def _line_style(self, value: object) -> None:
        wanted = int(to_integer(value, 'Long'))
        if wanted == _AUTOMATIC:
            self._automatic()
        elif wanted == _NONE:
            if not self.family.hides:
                raise error(1004, 'Unable to set the LineStyle property of the Border class')
            self._hide()
        elif wanted in _DASHES:
            before = self.look()
            look = self.line.change()
            if self.family.draws and self.family.styles_width and before.line_width is None:
                look.line_width = 12700
            look.line_off = look.line_by_weight = False
            look.line_rgb = look.line_rgb or '000000'
            look.line_dash = _DASHES[wanted]
        else:
            raise error(1004, 'Unable to set the LineStyle property of the Border class')

    def _automatic(self) -> None:
        before = self.look()
        if (before.line_rgb is not None or before.line_width is not None or before.line_dash is not None
                or before.line_off):
            _clear(self.line.change())

    def _hide(self) -> None:
        self.look()
        look = self.line.change()
        _clear(look)
        look.line_off = True
        look.line_width = self.family.hidden_width
