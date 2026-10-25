"""A sheet's printed headers and footers: the strings Excel keeps, how it reads and sets their sections, and the
headerFooter element a save writes.

Measured in live Excel (scripts/measure_page_headers.py, tests/fixtures/page_headers.json):

- Each header and footer is one string, its sections marked &L, &C and &R
  in either case. LeftHeader, CenterHeader and RightHeader, their footers,
  and the first and even pages' HeaderFooter.Text read those sections.
- The codes are found left to right, && an ampersand and no code. A code's
  section runs to the first code found from its second character on, so
  &L&C&R reads &C on the left and &R in the center, and &C&&Rb reads & in
  the center and loses Rb. Text before the first code is the center's,
  and a section marked twice reads both its parts.
- Setting a section cuts each of its parts, from its code to the next
  code, the center's text before the first code too, and adds the code and
  the new text at the end: setting the right, then the left, then the
  center header writes &R..&L..&C... A code in the text set reads as
  another section.
- The string is at most 255 characters, codes and all. Beyond that a
  setting is error 1004 and changes nothing, and a file holding a longer
  one is a file Excel does not open.
- The flags are headerFooter's attributes, written where they are not
  their defaults. The first and even pages' strings are written whether or
  not their flag is on, a carriage return as _x000D_ and a line feed as
  CR LF. A header with no strings and the flags at their defaults is not
  written at all, and one a macro left alone is written as the file had it.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from pyopenvba._opc_xml import opc_part
from pyopenvba._xml import escape_text
from pyopenvba.apps.excel import _xstring
from pyopenvba.exceptions import VBAUnsupportedError

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
#: The headerFooter element's children, in the schema's order.
CHILDREN = ('oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter')
#: Each flag's attribute and its value where the attribute is not written, in the schema's order.
FLAGS = {'OddAndEvenPagesHeaderFooter': ('differentOddEven', False),
         'DifferentFirstPageHeaderFooter': ('differentFirst', False),
         'ScaleWithDocHeaderFooter': ('scaleWithDoc', True),
         'AlignMarginsHeaderFooter': ('alignWithMargins', True)}
#: The longest header or footer string, its codes counted.
LIMIT = 255
#: The worksheet children a headerFooter element comes before, in the schema's order.
AFTER = ('rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing',
         'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts',
         'extLst')
_SECTIONS = 'LlCcRr'


def unsupported(detail: str) -> VBAUnsupportedError:
    return VBAUnsupportedError('PageSetup headers and footers ' + detail + ' are not implemented')


@dataclass
class Headers:
    """A sheet's header and footer strings by child, the flags a file or macro set, and whether a macro changed
    either since the sheet was read."""

    texts: dict[str, str] = field(default_factory=lambda: {})
    flags: dict[str, bool] = field(default_factory=lambda: {})
    changed: bool = False
    #: Whether the sheet draws a picture in a header, through a legacyDrawingHF part.
    pictures: bool = False

    def flag(self, attribute: str, default: bool) -> bool:
        return self.flags.get(attribute, default)


def _codes(text: str) -> list[tuple[int, str]]:
    """Each section code in ``text``, left to right, with its section: && is an ampersand, not a code."""
    found: list[tuple[int, str]] = []
    at = 0
    while at < len(text) - 1:
        if text[at] == '&' and text[at + 1] in _SECTIONS:
            found.append((at, text[at + 1].upper()))
            at += 2
        elif text[at] == '&' and text[at + 1] == '&':
            at += 2
        else:
            at += 1
    return found


def _end(text: str, start: int) -> int:
    """Where a section whose text starts at ``start`` ends: at the first code found from its second character on."""
    at = start + 1
    while at < len(text) - 1:
        if text[at] == '&' and text[at + 1] in _SECTIONS:
            return at
        at += 2 if text[at] == '&' and text[at + 1] == '&' else 1
    return len(text)


def _leads(codes: list[tuple[int, str]], text: str) -> bool:
    """Whether ``text`` has center text before its first code."""
    return bool(text) and (not codes or codes[0][0] > 0)


def section(text: str, code: str) -> str:
    """The section ``code`` (L, C or R) of a header or footer string, as Excel reads it."""
    codes = _codes(text)
    parts = [text[:codes[0][0] if codes else len(text)]] if code == 'C' and _leads(codes, text) else []
    parts += [text[at + 2:_end(text, at + 2)] for at, name in codes if name == code]
    return ''.join(parts)


def with_section(text: str, code: str, value: str) -> str:
    """A header or footer string with its section ``code`` set to ``value``, as Excel sets it."""
    codes = _codes(text)
    cuts = [(0, codes[0][0] if codes else len(text))] if code == 'C' and _leads(codes, text) else []
    cuts += [(at, codes[index + 1][0] if index + 1 < len(codes) else len(text))
             for index, (at, name) in enumerate(codes) if name == code]
    kept: list[str] = []
    place = 0
    for start, end in cuts:
        kept.append(text[place:start])
        place = end
    kept.append(text[place:])
    return ''.join(kept) + ('&' + code + value if value else '')


def read(data: bytes) -> Headers:
    """A sheet's headers and footers as its part ``data`` holds them."""
    from pyopenvba.apps.excel._charts_file import xml_tree

    part = opc_part(data, NS, 'worksheet')
    found = [(start, end) for name, _attrs, start, end in part.children if name == NS + '}headerFooter']
    result = Headers(pictures=any(name == NS + '}legacyDrawingHF' for name, *_rest in part.children))
    if not found:
        return result
    if len(found) > 1:
        raise unsupported('with more than one headerFooter element')
    start, end = found[0]
    try:
        root = xml_tree(data[start:end])
    except Exception:
        raise unsupported('with headerFooter markup the model cannot read') from None
    for flag, default in FLAGS.values():
        if flag in root.attrib:
            spelled = root.attrib[flag]
            if spelled not in ('1', '0', 'true', 'false'):
                raise unsupported('with a flag that is not a Boolean')
            if (spelled in ('1', 'true')) != default:
                result.flags[flag] = not default
    for child in root:
        if child.tag not in CHILDREN or child.tag in result.texts or len(child):
            raise unsupported('with headerFooter children outside the schema')
        text = _xstring.decode(child.text or '')
        if len(text) > LIMIT:
            raise unsupported('longer than 255 characters, in a file Excel does not open,')
        result.texts[child.tag] = text
    return result


def _spelled(text: str) -> str:
    """A header or footer string between tags as Excel writes it."""
    return escape_text(_xstring.escape(text)).replace('\n', '\r\n')


def element(headers: Headers) -> str | None:
    """The headerFooter element a save writes, or None where there is nothing to write."""
    flags = ''.join(f' {attribute}="{int(not default)}"' for attribute, default in FLAGS.values()
                    if headers.flag(attribute, default) != default)
    children = ''.join(f'<{name}>{_spelled(headers.texts[name])}</{name}>' for name in CHILDREN
                       if headers.texts.get(name))
    if not flags and not children:
        return None
    return f'<headerFooter{flags}>{children}</headerFooter>' if children else f'<headerFooter{flags}/>'


def with_headers(xml: str, headers: Headers) -> str:
    """The sheet's XML with its headerFooter element as ``headers`` has it."""
    data = xml.encode('utf-8')
    part = opc_part(data, NS, 'worksheet')
    wanted = element(headers)
    markup = (wanted or '').encode('utf-8')
    for name, _attrs, start, end in part.children:
        if name == NS + '}headerFooter':
            return (data[:start] + markup + data[end:]).decode('utf-8')
    if wanted is None:
        return xml
    at = next((start for name, _attrs, start, _end in part.children
               if name.rpartition('}')[2] in AFTER), part.closing_start)
    return part.insert(at, markup).decode('utf-8')
