"""Notes in a file: the comments part that holds their text, and the VML part that holds their boxes.

Measured in live Excel (scripts/measure_notes.py, tests/fixtures/notes.json):

- The comments part lists every author once and the notes row by row,
  each with shapeId="0", a fresh random xr:uid, and its text in one run of
  Tahoma 9; a note with no text is <text/>. A line feed is written as a
  CRLF, a carriage return as _x000D_, and text with a space at either end
  is xml:space="preserve".
- The VML part claims a block of 1024 shape ids for the sheet, its idmap,
  and draws each note from shape type 202, which comes before the first
  note it draws; the notes and form controls follow in the order they were
  made, each with a z-index counting them. Excel's VML writer breaks a
  line, two spaces in, before any piece of a shape's opening tag once the
  line has reached 70 characters.
- A sheet with notes names the VML part in <legacyDrawing>, and writes an
  empty <row> for each row that holds one; a note's cell counts in the
  sheet's dimension and its rows' spans.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from pyopenvba._a1 import parse_area
from pyopenvba._xml import attributes, escape_text, unescape

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet
    from pyopenvba.apps.excel._notes import Anchor, Note

COMMENTS_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
COMMENTS_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"

_HEAD = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
         '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
         ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr"'
         ' xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">')
#: The run every note's text is written in.
_RUN = '<rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr>'

#: A VML part with nothing drawn yet, for the sheet's block of ids.
_VML_HEAD = ('<xml xmlns:v="urn:schemas-microsoft-com:vml"\r\n'
             ' xmlns:o="urn:schemas-microsoft-com:office:office"\r\n'
             ' xmlns:x="urn:schemas-microsoft-com:office:excel">\r\n'
             ' <o:shapelayout v:ext="edit">\r\n'
             '  <o:idmap v:ext="edit" data="{block}"/>\r\n'
             ' </o:shapelayout>')
#: The shape type a note's box is drawn from.
NOTE_SHAPE_TYPE = ('<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"\r\n'
                   '  path="m,l,21600r21600,l21600,xe">\r\n'
                   '  <v:stroke joinstyle="miter"/>\r\n'
                   '  <v:path gradientshapeok="t" o:connecttype="rect"/>\r\n'
                   ' </v:shapetype>')
#: How long a line of a shape's opening tag gets before Excel breaks it.
_LINE = 70

_COMMENT = re.compile(r"<comment\b[^>]*?(?:/>|>.*?</comment>)", re.DOTALL)
_AUTHOR = re.compile(r"<author>(.*?)</author>|<author/>", re.DOTALL)
_TEXT = re.compile(r"<t\b[^>]*?(?:/>|>(.*?)</t>)", re.DOTALL)
_SHAPE = re.compile(r"<v:shape\b.*?</v:shape>", re.DOTALL)
_ESCAPED = re.compile(r"_x([0-9A-Fa-f]{4})_")


def empty_vml(block: int) -> str:
    return _VML_HEAD.format(block=block) + "</xml>"


# --- text ------------------------------------------------------------------------------------------------------


def encoded(text: str) -> str:
    """A note's text as the comments part spells it: a line feed CRLF, a carriage return _x000D_, and any other
    control character, or anything that reads as such an escape, escaped the same way."""
    text = re.sub(r"_(x[0-9A-Fa-f]{4}_)", r"_x005F_\1", text)
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", lambda match: f"_x{ord(match.group(0)):04X}_", text)
    return escape_text(text).replace("\r", "_x000D_").replace("\n", "\r\n")


def decoded(stored: str) -> str:
    """The text a stored note stands for: a line break read as a line feed, as XML reads one."""
    text = unescape(stored.replace("\r\n", "\n").replace("\r", "\n"))
    return _ESCAPED.sub(lambda match: chr(int(match.group(1), 16)), text)


def comment_xml(cell: str, author_id: int, note: Note) -> str:
    """One note's <comment>."""
    text = note.text
    if not text:
        body = "<text/>"
    else:
        space = ' xml:space="preserve"' if text != text.strip() else ""
        body = f"<text><r>{_RUN}<t{space}>{encoded(text)}</t></r></text>"
    return f'<comment ref="{cell}" authorId="{author_id}" shapeId="0" xr:uid="{note.uid}">{body}</comment>'


def comments_part(notes: list[tuple[str, Note]]) -> str:
    """The comments part for these notes, each with its cell, row by row."""
    authors: list[str] = []
    for _, note in notes:
        if note.author not in authors:
            authors.append(note.author)
    listed = "".join(f"<author>{escape_text(author)}</author>" for author in authors)
    entries = "".join(_retargeted(note.comment_xml, cell, authors.index(note.author)) if note.comment_xml
                      else comment_xml(cell, authors.index(note.author), note) for cell, note in notes)
    return f"{_HEAD}<authors>{listed}</authors><commentList>{entries}</commentList></comments>"


def _retargeted(markup: str, cell: str, author_id: int) -> str:
    """A <comment> a file gave, its formatting kept, naming the cell and the author it now has."""
    head_end = markup.find(">") + 1
    head = re.sub(r'\bref="[^"]*"', f'ref="{cell}"', markup[:head_end], count=1)
    head = re.sub(r'\bauthorId="[^"]*"', f'authorId="{author_id}"', head, count=1)
    return head + markup[head_end:]


# --- boxes -----------------------------------------------------------------------------------------------------


def _points(pixels: int) -> str:
    """Pixels as the points VML writes, three quarters of one each, with no trailing zeros."""
    quarters = pixels * 3
    if quarters % 4 == 0:
        return str(quarters // 4)
    return f"{quarters / 4:.2f}".rstrip("0")


def _wrapped(pieces: list[str]) -> str:
    """A shape's opening tag as Excel's VML writer lays it out: a line breaks, two spaces in, before any piece once
    it has reached 70 characters, the space a piece starts with dropped."""
    lines = [pieces[0]]
    for piece in pieces[1:]:
        if len(lines[-1]) >= _LINE:
            lines.append("  " + piece.lstrip(" "))
        else:
            lines[-1] += piece
    return "\r\n".join(lines)


def shape_xml(sheet: Worksheet, row: int, column: int, note: Note, z_index: int) -> str:
    """A note's box, as Excel writes one."""
    from pyopenvba.apps.excel._notes import anchor_of, box

    left, top, right, bottom = box(sheet, note)
    shown = "visible" if note.visible else "hidden"
    tag = _wrapped([
        "<v:shape", f' id="_x0000_s{note.shape_id}"', ' type="#_x0000_t202"', " style='position:absolute;",
        f"margin-left:{_points(left)}pt;", f"margin-top:{_points(top)}pt;", f"width:{_points(right - left)}pt;",
        f"height:{_points(bottom - top)}pt;", f"z-index:{z_index};", f"visibility:{shown}'",
        ' fillcolor="infoBackground [80]"', ' strokecolor="none [81]"', ' o:insetmode="auto"', ">"])
    corners = ", ".join(str(number) for number in anchor_of(sheet, note))
    flag = "\r\n   <x:Visible/>" if note.visible else ""
    return (tag + "\r\n"
            '  <v:fill color2="infoBackground [80]"/>\r\n'
            '  <v:shadow color="none [81]" obscured="t"/>\r\n'
            '  <v:path o:connecttype="none"/>\r\n'
            "  <v:textbox style='mso-direction-alt:auto'>\r\n"
            "   <div style='text-align:left'></div>\r\n"
            "  </v:textbox>\r\n"
            '  <x:ClientData ObjectType="Note">\r\n'
            "   <x:MoveWithCells/>\r\n"
            "   <x:SizeWithCells/>\r\n"
            "   <x:Anchor>\r\n"
            f"    {corners}</x:Anchor>\r\n"
            "   <x:AutoFill>False</x:AutoFill>\r\n"
            f"   <x:Row>{row - 1}</x:Row>\r\n"
            f"   <x:Column>{column - 1}</x:Column>{flag}\r\n"
            "  </x:ClientData>\r\n"
            " </v:shape>")


def _shape_id(shape: str) -> int:
    found = re.search(r'\bid="_x0000_s(\d+)"', shape) or re.search(r'\bo:spid="_x0000_s(\d+)"', shape)
    return int(found.group(1)) if found else 0


def with_notes(sheet: Worksheet, vml: str) -> str:
    """The VML part with the sheet's notes drawn as they now are: each note's box where it was, taken out where
    the note has gone, and a new note's among the other shapes by its id, which counts them in the order they
    were made; each shape type in front of the first shape drawn from it."""
    notes = {note.shape_id: (key, note) for key, note in sheet.notes.items()}
    # The part in pieces: the text between shapes, and each shape by its id, None for a note to be drawn afresh.
    pieces: list[str | tuple[int, str | None]] = []
    at = 0
    for match in _SHAPE.finditer(vml):
        pieces.append(vml[at:match.start()])
        at = match.end()
        shape, shape_id = match.group(0), _shape_id(match.group(0))
        if 'ObjectType="Note"' not in shape:
            pieces.append((shape_id, shape))
        elif shape_id in notes:
            pieces.append((shape_id, None))
    tail = vml[at:]
    closing = tail.rfind("</xml>")
    pieces += [tail[:closing] if closing >= 0 else tail, tail[closing:] if closing >= 0 else "</xml>"]
    drawn_ids = {piece[0] for piece in pieces if isinstance(piece, tuple)}
    for shape_id in sorted(notes):
        if shape_id in drawn_ids:
            continue
        later = next((index for index, piece in enumerate(pieces) if isinstance(piece, tuple) and piece[0] > shape_id),
                     len(pieces) - 1)
        pieces.insert(later, (shape_id, None))
    # Every shape is numbered in the order it is drawn, the notes and the controls alike.
    out: list[str] = []
    z_index = 0
    for piece in pieces:
        if isinstance(piece, str):
            out.append(piece)
            continue
        z_index += 1
        shape_id, shape = piece
        if shape is None:
            key, note = notes[shape_id]
            shape = note.shape_xml or shape_xml(sheet, *key, note, z_index)
        out.append(shape)
    from pyopenvba.shapes._xlsx import CONTROL_SHAPE_TYPE

    return _typed_first(_typed_first("".join(out), "202", NOTE_SHAPE_TYPE), "201", CONTROL_SHAPE_TYPE)


def _typed_first(vml: str, number: str, shape_type: str) -> str:
    """The part with shape type ``number`` just in front of the first shape drawn from it, as Excel writes it:
    moved there, or put there when the part has none (tests/fixtures/notes.json)."""
    first = vml.find(f'type="#_x0000_t{number}"')
    if first < 0:
        return vml
    user = vml.rfind("<v:shape ", 0, first)
    found = re.search(rf'<v:shapetype id="_x0000_t{number}".*?</v:shapetype>', vml, re.DOTALL)
    if found is not None:
        if found.end() == user:
            return vml
        shape_type = found.group(0)
        vml = vml[:found.start()] + vml[found.end():]
        if found.start() < user:
            user -= len(shape_type)
    return vml[:user] + shape_type + vml[user:]


# --- reading -----------------------------------------------------------------------------------------------------


def read(sheet: Worksheet, comments: str, vml: str) -> None:
    """The notes a sheet's comments and VML parts hold, and the block of ids and the last id its VML part used."""
    from pyopenvba.apps.excel._notes import BLOCK, Note, size_of

    block_found = re.search(r'<o:idmap\b[^>]*\bdata="(\d+)', vml)
    shapes = [match.group(0) for match in _SHAPE.finditer(vml)]
    ids = [_shape_id(shape) for shape in shapes]
    if block_found is not None:
        sheet.vml_block = int(block_found.group(1))
    elif ids:
        sheet.vml_block = max(ids) // BLOCK
    sheet.vml_last_id = max(ids, default=0)
    if not comments:
        return
    authors = [decoded(found.group(1) or "") for found in _AUTHOR.finditer(comments.split("</authors>")[0])]
    entries: dict[tuple[int, int], tuple[str, str, str, str]] = {}
    for element in _COMMENT.finditer(comments):
        markup = element.group(0)
        fields = attributes(markup[:markup.find(">") + 1])
        try:
            area = parse_area(fields.get("ref", ""))
        except ValueError:
            continue
        index = int(fields.get("authorId", "0") or 0)
        author = authors[index] if 0 <= index < len(authors) else ""
        text = "".join(decoded(piece or "") for piece in _TEXT.findall(markup))
        entries[(area.top, area.left)] = (text, author, fields.get("xr:uid", ""), markup)
    boxes: dict[tuple[int, int], tuple[str, int, bool, Anchor]] = {}
    for shape, shape_id in zip(shapes, ids, strict=True):
        if 'ObjectType="Note"' not in shape:
            continue
        row = re.search(r"<x:Row>\s*(\d+)\s*</x:Row>", shape)
        column = re.search(r"<x:Column>\s*(\d+)\s*</x:Column>", shape)
        anchor = re.search(r"<x:Anchor>\s*([\d\s,]+?)\s*</x:Anchor>", shape)
        if row is None or column is None or anchor is None:
            continue
        numbers = [int(number) for number in anchor.group(1).split(",")]
        if len(numbers) != 8:
            continue
        corners: Anchor = (numbers[0], numbers[1], numbers[2], numbers[3], numbers[4], numbers[5], numbers[6],
                           numbers[7])
        visible = re.search(r"<x:Visible\s*/>", shape) is not None
        boxes[(int(row.group(1)) + 1, int(column.group(1)) + 1)] = (shape, shape_id, visible, corners)
    start = sheet.vml_block * BLOCK
    for key in [*[key for key in boxes if key in entries], *[key for key in entries if key not in boxes]]:
        text, author, uid, markup = entries[key]
        shape, shape_id, visible, corners = boxes.get(key, ("", 0, False, (0, 0, 0, 0, 0, 0, 0, 0)))
        width, height = size_of(sheet, corners)
        # Excel names a note after its shape's place in the sheet's block of ids.
        sheet.notes[key] = Note(text=text, author=author, visible=visible, name=f"Comment {shape_id - start}",
                                shape_id=shape_id, anchor=corners, width=width, height=height, uid=uid,
                                comment_xml=markup, shape_xml=shape)
    sheet.shape_count = max([sheet.shape_count, *(note.shape_id - start for note in sheet.notes.values())])


__all__ = ["COMMENTS_RELATIONSHIP", "COMMENTS_TYPE", "NOTE_SHAPE_TYPE", "comments_part", "decoded", "empty_vml",
           "encoded", "read", "shape_xml", "with_notes"]
