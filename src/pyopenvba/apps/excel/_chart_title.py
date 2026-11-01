"""A chart title's link and formatting, kept as Excel keeps them: each setting is either left alone or set, and only
the settings once set reach the file (tests/fixtures/chart_title.json)."""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyopenvba._xml import escape_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts_file import XmlElement as Element

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

    def rotated(self) -> bool:
        return self.rotation is not None and self.rotation != 0


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
