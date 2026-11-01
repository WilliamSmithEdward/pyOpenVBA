"""A chart title's link and formatting, kept as Excel keeps them: each setting is either left alone or set, and only
the settings once set reach the file (tests/fixtures/chart_title.json)."""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyopenvba._xml import escape_text
from pyopenvba.apps.excel._creator import Created
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import member, setter
from pyopenvba.interpreter._values import VBAInt, error, to_bool, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts_file import XmlElement as Element
    from pyopenvba.apps.excel._model import Worksheet

#: ChartTitle.HorizontalAlignment and how a paragraph spells it.
ALIGNMENTS = {-4131: 'l', -4152: 'r', -4108: 'ctr', -4130: 'just', -4117: 'dist'}
#: ChartTitle.Orientation constants, as angles; xlVertical stacks the letters instead.
ORIENTATIONS = {-4128: 0, -4171: 90, -4170: -90}
VERTICAL = -4166
#: ChartTitle.ReadingOrder: xlContext leaves it to the text, xlRTL and xlLTR set it.
READING_ORDERS = {-5002: None, -5004: True, -5003: False}
VERTICAL_ALIGNMENTS = (-4160, -4108, -4107, -4130, -4117)
_SHADOW = ('<c:spPr><a:effectLst><a:outerShdw dist="35921" dir="2700000" algn="br"><a:srgbClr val="000000"/>'
           '</a:outerShdw></a:effectLst></c:spPr>')


@dataclass
class TitleLook:
    #: The cell the title shows, as 'Sheet1!$B$1'; empty when it shows text.
    formula: str = ''
    #: Degrees, VERTICAL, or None when never set.
    rotation: int | None = None
    alignment: str | None = None
    right_to_left: bool | None = None
    overlay: bool = False
    shadow: bool | None = None

    def turned(self, default: int) -> int:
        """The angle the title shows at: its own once set, else the one its place gives it."""
        return default if self.rotation is None else self.rotation


class TitleMembers(Created):
    """The members a chart title and an axis title share. Each owner says where its look lives, how a change is
    admitted, which sheet its links resolve against, and the angle it shows at until one is set; Text and Delete stay
    with the owner."""

    #: The angle the title shows at until Orientation is set: 90 for the title of an upright axis.
    title_default_angle = 0

    def title_look(self) -> TitleLook:
        raise NotImplementedError

    def title_restyle(self) -> TitleLook:
        raise NotImplementedError

    def title_sheet(self) -> Worksheet:
        raise NotImplementedError

    def title_link(self, formula: str) -> None:
        raise NotImplementedError

    def title_angle(self) -> int:
        return self.title_default_angle

    def _text(self, value: object) -> None:
        raise NotImplementedError

    def Text(self) -> object:  # noqa: N802 - the owner's VBA member
        raise NotImplementedError

    @member
    def Caption(self) -> object:
        return self.Text()

    @setter('Caption')
    def _caption(self, value: object) -> None:
        self._text(value)

    @member
    def Formula(self) -> object:
        formula = self.title_look().formula
        return '=' + formula if formula else self.Text()

    @setter('Formula')
    def _formula(self, value: object) -> None:
        self.link(to_text(value), r1c1=False)

    @member
    def FormulaR1C1(self) -> object:
        from pyopenvba.apps.excel._charts import r1c1_series

        formula = self.title_look().formula
        return r1c1_series('=' + formula) if formula else self.Text()

    @setter('FormulaR1C1')
    def _formula_r1c1(self, value: object) -> None:
        self.link(to_text(value), r1c1=True)

    @member
    def FormulaLocal(self) -> object:
        return self.Formula()

    @setter('FormulaLocal')
    def _formula_local(self, value: object) -> None:
        self._formula(value)

    @member
    def FormulaR1C1Local(self) -> object:
        return self.FormulaR1C1()

    @setter('FormulaR1C1Local')
    def _formula_r1c1_local(self, value: object) -> None:
        self._formula_r1c1(value)

    def link(self, text: str, *, r1c1: bool) -> None:
        """Formula: a reference to one cell links the title to it; anything not starting with = is typed text."""
        from pyopenvba._a1 import Area, column_number
        from pyopenvba.apps.excel._charts import reference

        owner = self.title_sheet()
        if not text.startswith('='):
            self._text(text)
            return
        found = re.fullmatch(r"=((?:'[^']+'|[^!'=]+))!(?:\$?([A-Za-z]{1,3})\$?(\d+)|R(\d+)C(\d+))", text)
        if found is None:
            if re.fullmatch(r"=[^!]*![^!]*", text):
                raise VBAUnsupportedError('Linking a chart title to this reference is not implemented')
            raise error(1004, 'Unable to set the Formula property of the title')
        sheet_name = found.group(1)[1:-1].replace("''", "'") if found.group(1).startswith("'") else found.group(1)
        sheet = next((item for item in owner.book.sheets_ if item.name.lower() == sheet_name.lower()), None)
        if sheet is None or (found.group(4) is not None) != r1c1:
            raise VBAUnsupportedError('Linking a chart title to this reference is not implemented')
        row, column = ((int(found.group(4)), int(found.group(5))) if r1c1
                       else (int(found.group(3)), column_number(found.group(2).upper())))
        self.title_link(reference(sheet, Area(row, column, row, column)))

    @member
    def Orientation(self) -> object:
        rotation = self.title_look().turned(self.title_angle())
        if rotation == VERTICAL:
            return VBAInt(VERTICAL, 'Long')
        named = {angle: number for number, angle in ORIENTATIONS.items()}
        return VBAInt(named.get(rotation, rotation), 'Long')

    @setter('Orientation')
    def _orientation(self, value: object) -> None:
        wanted = int(to_integer(value, 'Long'))
        angle = VERTICAL if wanted == VERTICAL else ORIENTATIONS.get(wanted, wanted)
        if angle != VERTICAL and not -90 <= angle <= 90:
            raise error(-2147467259, 'Automation error\nUnspecified error')
        self.title_restyle().rotation = angle

    @member
    def HorizontalAlignment(self) -> object:
        spelled = self.title_look().alignment or 'ctr'
        return VBAInt(next(number for number, name in ALIGNMENTS.items() if name == spelled), 'Long')

    @setter('HorizontalAlignment')
    def _horizontal_alignment(self, value: object) -> None:
        wanted = ALIGNMENTS.get(int(to_integer(value, 'Long')))
        if wanted is None:
            raise error(-2147467259, 'Automation error\nUnspecified error')
        angle = self.title_look().turned(self.title_angle())
        if angle == VERTICAL:
            raise VBAUnsupportedError('Aligning a vertical chart title is not implemented')
        if angle == 0:
            # Excel ignores the alignment of text it has turned.
            self.title_restyle().alignment = wanted

    @member
    def VerticalAlignment(self) -> object:
        self.title_look()
        return VBAInt(-4108, 'Long')

    @setter('VerticalAlignment')
    def _vertical_alignment(self, value: object) -> None:
        if int(to_integer(value, 'Long')) not in VERTICAL_ALIGNMENTS:
            raise VBAUnsupportedError('This chart title vertical alignment is not implemented')
        self.title_look()  # Excel keeps a title centred whatever is set.

    @member
    def ReadingOrder(self) -> object:
        setting = self.title_look().right_to_left
        return VBAInt(next(number for number, value in READING_ORDERS.items() if value is setting), 'Long')

    @setter('ReadingOrder')
    def _reading_order(self, value: object) -> None:
        wanted = int(to_integer(value, 'Long'))
        if wanted not in READING_ORDERS:
            raise VBAUnsupportedError('This chart title reading order is not implemented')
        self.title_restyle().right_to_left = READING_ORDERS[wanted]

    @member
    def Position(self) -> object:
        self.title_look()
        return VBAInt(-4105, 'Long')

    @setter('Position')
    def _position(self, value: object) -> None:
        if int(to_integer(value, 'Long')) != -4105:
            raise VBAUnsupportedError('Placing a chart title requires native layout')
        self.title_look()

    @member
    def IncludeInLayout(self) -> object:
        return not self.title_look().overlay

    @setter('IncludeInLayout')
    def _include_in_layout(self, value: object) -> None:
        wanted = to_bool(value)
        self.title_restyle().overlay = not wanted

    @member
    def Shadow(self) -> object:
        return bool(self.title_look().shadow)

    @setter('Shadow')
    def _shadow(self, value: object) -> None:
        wanted = to_bool(value)
        self.title_restyle().shadow = wanted


def title_element(look: TitleLook, *, text: str | None, linked_text: str = '') -> str:
    """The title element: ``text`` typed in, the linked cell's text, or with neither Excel's automatic title."""
    if look.rotation == VERTICAL:
        body = '<a:bodyPr rot="0" vert="wordArtVert"/>'
    elif look.rotation is not None:
        body = f'<a:bodyPr rot="{-look.rotation * 60000}" vert="horz"/>'
    else:
        body = '<a:bodyPr/>'
    settings = (f' algn="{look.alignment}"' if look.alignment else '') + (
        '' if look.right_to_left is None else f' rtl="{int(look.right_to_left)}"')
    paragraph = f'<a:pPr{settings}><a:defRPr/></a:pPr>'
    formatted = look.rotation is not None or bool(settings)
    properties = ''
    if text is not None:
        run = f'<a:r><a:rPr lang="en-US"/><a:t>{escape_text(text)}</a:t></a:r>' if text else '<a:endParaRPr lang="en-US"/>'
        source = f'<c:tx><c:rich>{body}<a:lstStyle/><a:p>{paragraph}{run}</a:p></c:rich></c:tx>'
    else:
        source = ''
        if look.formula:
            source = (f'<c:tx><c:strRef><c:f>{escape_text(look.formula)}</c:f><c:strCache><c:ptCount val="1"/>'
                      f'<c:pt idx="0"><c:v>{escape_text(linked_text)}</c:v></c:pt></c:strCache></c:strRef></c:tx>')
        if formatted:
            properties = f'<c:txPr>{body}<a:lstStyle/><a:p>{paragraph}<a:endParaRPr lang="en-US"/></a:p></c:txPr>'
    shadow = '' if look.shadow is None else _SHADOW if look.shadow else '<c:spPr><a:effectLst/></c:spPr>'
    return f'<c:title>{source}<c:overlay val="{int(look.overlay)}"/>{shadow}{properties}</c:title>'


def read_look(title: Element, namespaces: dict[str, str]) -> TitleLook:
    """The link and formatting of an imported title, as far as they are spelled as Excel spells them."""
    look = TitleLook()
    formula = title.find('c:tx/c:strRef/c:f', namespaces)
    look.formula = formula.text or '' if formula is not None else ''
    overlay = title.find('c:overlay', namespaces)
    look.overlay = overlay is not None and overlay.get('val', '1') in ('1', 'true')
    shadow = title.find('c:spPr/a:effectLst', namespaces)
    if shadow is not None:
        look.shadow = shadow.find('a:outerShdw', namespaces) is not None
    body = title.find('c:tx/c:rich/a:bodyPr', namespaces)
    if body is None:
        body = title.find('c:txPr/a:bodyPr', namespaces)
    if body is not None and body.get('vert') == 'wordArtVert':
        look.rotation = VERTICAL
    elif body is not None and re.fullmatch(r'-?[0-9]{1,8}', body.get('rot', '')):
        look.rotation = -int(body.get('rot', '0')) // 60000
    paragraph = title.find('c:tx/c:rich/a:p/a:pPr', namespaces)
    if paragraph is None:
        paragraph = title.find('c:txPr/a:p/a:pPr', namespaces)
    if paragraph is not None:
        if paragraph.get('algn') in ALIGNMENTS.values():
            look.alignment = paragraph.get('algn')
        if paragraph.get('rtl') in ('0', '1'):
            look.right_to_left = paragraph.get('rtl') == '1'
    return look
