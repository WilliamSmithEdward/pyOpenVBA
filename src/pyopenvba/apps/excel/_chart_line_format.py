"""The Format of a chart's trendlines and error bars, which are lines and nothing else
(tests/fixtures/chart_line_format.json). It shares its look with their Border."""
from __future__ import annotations

from pyopenvba.apps.excel._chart_borders import Line
from pyopenvba.apps.excel._chart_format import plain_line
from pyopenvba.apps.excel._creator import Created
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import member, setter
from pyopenvba.interpreter._values import VBAInt, VBASingle, error, to_integer, to_number
from pyopenvba.shapes._values import FillLine

_MIXED = -2147483648.0
_INVALID = (5, 'Invalid procedure call or argument')
_NO_MEMBER = (-2147467259, 'Automation error\nUnspecified error ')
_EMU_PER_POINT = 12700
#: LineFormat.DashStyle and how the file writes it: its prstDash and the cap it takes.
_DASHES = {1: ('solid', None), 2: ('dot', 'sq'), 3: ('dot', 'rnd'), 4: ('dash', None), 5: ('dashDot', None),
           6: ('sysDashDotDot', None), 7: ('lgDash', None), 8: ('lgDashDot', None), 9: ('lgDashDotDot', None),
           10: ('sysDash', None), 11: ('sysDot', None), 12: ('sysDashDot', None)}
#: LineFormat.Style and the compound line the file writes.
_STYLES = {1: 'sng', 2: 'dbl', 3: 'thinThick', 4: 'thickThin', 5: 'tri'}
_WINDOW_TEXT = '<a:sysClr val="windowText" lastClr="000000"/>'
_A14 = 'xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"'


def _unmeasured(what: str) -> VBAUnsupportedError:
    return VBAUnsupportedError(f'{what} on a trendline or error bars is not implemented')


def shape_properties(look: FillLine, cap: str, painted: bool) -> str:
    """The element's spPr: what was set of the line, or, hidden, the stroke Excel writes out in full, with the colour it
    keeps. Hidden error bars are painted as well, and their line's cap differs."""
    if look.line_hidden:
        colour = _WINDOW_TEXT if look.line_rgb is None else f'<a:srgbClr val="{look.line_rgb}"/>'
        stroke = f'w="12700" cap="{cap}" cmpd="sng" algn="ctr"'
        fill = f'<a:solidFill>{_WINDOW_TEXT}</a:solidFill>' if painted else ''
        return (f'<c:spPr>{fill}<a:ln {stroke}><a:noFill/><a:prstDash val="solid"/><a:round/></a:ln><a:effectLst/>'
                '<a:extLst><a:ext uri="{91240B29-F687-4F45-9708-019B960494DF}">'
                f'<a14:hiddenLine {_A14} {stroke}><a:solidFill>{colour}</a:solidFill>'
                '<a:prstDash val="solid"/><a:round/>'
                '</a14:hiddenLine></a:ext></a:extLst></c:spPr>')
    written = plain_line(look)
    return f'<c:spPr>{written}</c:spPr>' if written else ''


class LineColor(Created):
    vba_type_name = 'ColorFormat'

    def __init__(self, line: Line) -> None:
        self.line = line

    @member
    def RGB(self) -> object:
        look = self.line.look()
        if look.line_off:
            return VBAInt(16777215, 'Long')
        text = look.line_rgb or '000000'
        return VBAInt(int(text[4:6] + text[2:4] + text[:2], 16), 'Long')

    @setter('RGB')
    def _rgb(self, value: object) -> None:
        colour = int(to_integer(value, 'Long')) & 0xFFFFFF
        before = self.line.look()
        if before.line_off:
            raise _unmeasured('A colour once its Border was hidden')
        if not before.line_hidden:
            look = self.line.change()
            look.line_rgb = f'{colour & 0xFF:02X}{colour >> 8 & 0xFF:02X}{colour >> 16 & 0xFF:02X}'
            look.line_by_weight = False

    @member
    def Type(self) -> object:
        return VBAInt(1, 'Long')

    @member
    def ObjectThemeColor(self) -> object:
        return VBAInt(0, 'Long')


class LineLineFormat(Created):
    vba_type_name = 'LineFormat'

    def __init__(self, line: Line) -> None:
        self.line = line

    def _touched(self, look: FillLine) -> bool:
        return (look.line_rgb is not None or look.line_width is not None or look.line_dash is not None
                or look.line_compound is not None or look.line_hidden)

    def _plain(self, what: str) -> FillLine:
        before = self.line.look()
        if before.line_off or before.line_hidden:
            raise _unmeasured(f'{what} once the line was hidden')
        return self.line.change()

    @member
    def Visible(self) -> object:
        look = self.line.look()
        return VBAInt(0 if look.line_hidden or look.line_off else -1, 'Long')

    @setter('Visible')
    def _visible(self, value: object) -> None:
        wanted = int(to_integer(value, 'Long'))
        look = self.line.look()
        hidden = look.line_hidden or look.line_off
        if wanted == -1 and not hidden:
            return
        if wanted == 0 and look.line_hidden:
            return
        if wanted != 0 or look.line_off:
            raise _unmeasured('Showing the line again, or Visible set to msoCTrue,')
        self.line.change().line_hidden = True

    @member
    def ForeColor(self) -> object:
        return LineColor(self.line)

    @member
    def Weight(self) -> object:
        look = self.line.look()
        if look.line_off:
            return VBASingle(_MIXED)
        return VBASingle(1.0 if look.line_width is None or look.line_hidden else look.line_width / _EMU_PER_POINT)

    @setter('Weight')
    def _weight(self, value: object) -> None:
        wanted = float(VBASingle(float(to_number(value))))
        if not 0 <= wanted <= 1584:
            raise error(*_INVALID)
        before = self.line.look()
        if before.line_off:
            raise _unmeasured('A weight once its Border was hidden')
        if not before.line_hidden:
            self.line.change().line_width = round(wanted * _EMU_PER_POINT)

    @member
    def DashStyle(self) -> object:
        look = self.line.look()
        if look.line_off:
            return VBAInt(-2, 'Long')
        if look.line_dash == 'dot' and look.line_cap == 'sq':
            return VBAInt(2, 'Long')
        return VBAInt(next((key for key, (name, cap) in _DASHES.items()
                            if name == look.line_dash and cap == look.line_cap), 1), 'Long')

    @setter('DashStyle')
    def _dash_style(self, value: object) -> None:
        wanted = int(to_integer(value, 'Long'))
        if wanted not in _DASHES:
            raise error(*_INVALID)
        look = self._plain('A dash style')
        look.line_dash, look.line_cap = _DASHES[wanted]

    @member
    def Style(self) -> object:
        look = self.line.look()
        if look.line_off:
            return VBAInt(-2, 'Long')
        return VBAInt(next((key for key, name in _STYLES.items() if name == look.line_compound), 1), 'Long')

    @setter('Style')
    def _style(self, value: object) -> None:
        wanted = int(to_integer(value, 'Long'))
        if wanted not in _STYLES:
            raise error(*_INVALID)
        self._plain('A style').line_compound = _STYLES[wanted]

    @member
    def Transparency(self) -> object:
        look = self.line.look()
        return VBASingle(0.0 if self._touched(look) and not look.line_off else _MIXED)

    @setter('Transparency')
    def _transparency(self, value: object) -> None:
        wanted = float(to_number(value))
        if not 0 <= wanted <= 1:
            raise error(*_INVALID)
        if self._touched(self.line.look()):
            raise _unmeasured('Transparency on a line with a look')

    @member
    def BeginArrowheadStyle(self) -> object:
        self.line.look()
        return VBAInt(1, 'Long')

    @member
    def EndArrowheadStyle(self) -> object:
        self.line.look()
        return VBAInt(1, 'Long')


class LineChartFormat(Created):
    vba_type_name = 'ChartFormat'

    def __init__(self, parent: object, line: Line) -> None:
        self.parent, self.line = parent, line

    @member
    def Parent(self) -> object:
        return self.parent

    @member
    def Line(self) -> object:
        return LineLineFormat(self.line)

    @member
    def Fill(self) -> object:
        raise error(*_NO_MEMBER)

    @member
    def ThreeD(self) -> object:
        raise error(*_NO_MEMBER)

    @member
    def TextFrame2(self) -> object:
        raise error(*_NO_MEMBER)

    @member
    def PictureEffects(self) -> object:
        raise error(438, "Object doesn't support this property or method")
