"""The text a shape in a file holds, as the TextLook and marks a new shape's TextFrame and TextRange2 work from.

Measured in live Excel (scripts/measure_loaded_shape_text.py, tests/fixtures/loaded_shape_text.json). Letters painted with a gradient,
a picture or a pattern are not read: a macro reaching them is told so, rather than given a wrong answer.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field, fields
from pyopenvba._xml import unescape
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.shapes import _drawingml as dml
from pyopenvba.shapes._read_look import _attribute, _paint  # pyright: ignore[reportPrivateUsage]
from pyopenvba.shapes._values import TextLine, TextLook

_PARAGRAPH = re.compile(r"<a:p(?:\s[^>]*)?(?:/>|>(.*?)</a:p>)", re.DOTALL)
_PROPERTIES = re.compile(r"<a:pPr\b([^>]*?)(?:/>|>(.*?)</a:pPr>)", re.DOTALL)
_PIECE = re.compile(
    r"<a:(r|fld)\b[^>]*>(?:<a:rPr\b([^>]*?)(?:/>|>(.*?)</a:rPr>))?(?:<a:pPr\b.*?(?:/>|</a:pPr>))?(?:<a:t>(.*?)</a:t>|<a:t/>)?</a:\1>"
    r"|<a:br\b[^>]*?(?:/>|>(?:<a:rPr\b([^>]*?)(?:/>|>(.*?)</a:rPr>))?</a:br>)"
    r"|<a:endParaRPr\b([^>]*?)(?:/>|>(.*?)</a:endParaRPr>)", re.DOTALL)
_INDEX = re.compile(r'a14:legacySpreadsheetColorIndex="(\d+)"')
_TYPEFACE = re.compile(r'\btypeface="([^"]*)"')
#: What a run's rPr says, by attribute, as the field it sets.
_NUMBERS = {"kern": "font_kern", "spc": "font_spacing", "baseline": "font_baseline"}
_WORDS = {"u": "font_underline", "strike": "font_strike", "cap": "font_cap"}
_FLAGS = {"b": "font_bold", "i": "font_italic"}
_DEFAULTS: dict[str, object] = {one.name: one.default for one in fields(TextLook) if one.name.startswith("font_")}


@dataclass(slots=True)
class ReadText:
    """A shape's text: its look, each character's own settings where they differ, its paragraph's end, and each paragraph's pPr
    where they are not all the look's."""

    look: TextLook
    marks: list[dict[str, object]] = field(default_factory=lambda: [])
    end: dict[str, object] | None = None
    paragraphs: list[str] | None = None


def _body(markup: str) -> str | None:
    span = dml.element_span(markup, "xdr:txBody")
    return None if span is None else markup[span[0]:span[1]]


def _run_text(raw: str | None) -> str:
    """A run's text as the model holds it: a line break within a paragraph, which the file spells CR LF, is a vertical tab."""
    return unescape(raw or "").replace("\r\n", "\x0b").replace("\r", "\x0b").replace("\n", "\x0b")


def loaded_text(markup: str) -> str:
    """A loaded shape's text: \\n ends a paragraph, \\x0b breaks a line within one as Excel writes it, and \\x0c is an a:br, which
    reads as a vertical tab where Excel's own break reads as a line feed."""
    body = _body(markup)
    if body is None:
        return ""
    lines: list[str] = []
    for paragraph in _PARAGRAPH.finditer(body):
        line = ""
        for piece in _PIECE.finditer(paragraph.group(1) or ""):
            if piece.group(0).startswith("<a:br"):
                line += "\x0c"
            elif not piece.group(0).startswith("<a:endParaRPr"):
                line += _run_text(piece.group(4))
        lines.append(line)
    return "\n".join(lines)


def _text_line(piece: str) -> TextLine:
    head = piece[:piece.find(">")]
    width, compound = _attribute(head, "w"), _attribute(head, "cmpd")
    if "<a:noFill" in piece:
        return TextLine(hidden=True)
    if "<a:gradFill" in piece or "<a:pattFill" in piece:
        raise VBAUnsupportedError("The line of letters with a gradient or a pattern is not implemented")
    dash = re.search(r'<a:prstDash\b[^>]*\bval="([^"]*)"', piece)
    full = _attribute(head, "cap") is not None
    solid = dml.element_span(piece, "a:solidFill")
    inside = "" if solid is None else piece[solid[0]:solid[1]]
    if "<a:prstClr" in inside:
        mods = tuple((name, int(value)) for name, value in re.findall(r'<a:(\w+) val="(-?\d+)"/>', inside))
        kept = tuple(one for one in mods if one[0] != "alpha")
        alpha = next((value for name, value in mods if name == "alpha"), None)
        return TextLine(preset=True, mods=kept, alpha=alpha, width=None if width is None else int(width),
                        compound=compound, dash=None if dash is None else dash.group(1), full=full)
    paint = _paint(inside) if inside else None
    return TextLine(rgb=None if paint is None else paint.rgb, scheme=None if paint is None else paint.scheme,
                    mods=() if paint is None else paint.mods, alpha=None if paint is None else paint.alpha,
                    width=None if width is None else int(width), compound=compound,
                    dash=None if dash is None else dash.group(1), full=full)


def _settings(attributes: str | None, inside: str | None, scheme: str) -> dict[str, object]:
    """What a run's rPr or a paragraph's endParaRPr sets, as TextLook's fields."""
    found: dict[str, object] = {}
    attributes, inside = attributes or "", inside or ""
    size = _attribute(attributes, "sz")
    if size is not None:
        found["font_size"] = int(size) / 100
    for name, target in _FLAGS.items():
        value = _attribute(attributes, name)
        if value is not None:
            found[target] = value in ("1", "true")
    for name, target in _WORDS.items():
        value = _attribute(attributes, name)
        if value is not None:
            found[target] = value
    for name, target in _NUMBERS.items():
        value = _attribute(attributes, name)
        if value is not None:
            found[target] = int(value)
    if _attribute(attributes, "normalizeH") in ("1", "true"):
        found["font_normalize"] = True
    rest = inside
    for tag, target in (("a:ln", "font_line"), ("a:highlight", "font_highlight"), ("a:uLn", "font_underline_line"), ("a:uFill", "font_underline_fill")):
        span = dml.element_span(rest, tag)
        if span is not None:
            piece = rest[span[0]:span[1]]
            found[target] = _text_line(piece) if tag == "a:ln" else piece
            rest = rest[:span[0]] + rest[span[1]:]
    for tag in ("a:blipFill", "a:grpFill"):
        if f"<{tag}" in rest:
            raise VBAUnsupportedError(f"The text of a shape whose letters have {tag[2:]} is not implemented")
    for tag, kind in (("a:gradFill", 3), ("a:pattFill", 2)):
        span = dml.element_span(rest, tag)
        if span is not None:
            # Read as its first colour.
            found["font_fill_type"] = kind
            found.update(_ink(rest[span[0]:span[1]], scheme))
    if "<a:noFill" in rest:
        found["font_no_fill"] = True
    solid = dml.element_span(rest, "a:solidFill")
    if solid is not None:
        found.update(_ink(rest[solid[0]:solid[1]], scheme))
    effects = dml.element_span(rest, "a:effectLst")
    if effects is not None:
        piece = rest[effects[0]:effects[1]]
        if "<a:outerShdw" in piece:
            found["font_shadow"] = True
        if re.search(r"<a:(glow|reflection|softEdge|innerShdw|prstShdw|blur|fillOverlay)\b", piece):
            found["font_effects"] = piece
    for tag, target in (("a:latin", "font_latin"), ("a:ea", "font_ea"), ("a:cs", "font_cs"), ("a:sym", "font_sym")):
        typeface = re.search(rf"<{tag}\b[^>]*/>", rest)
        if typeface is not None:
            found[target] = typeface.group(0)
    return found


def _ink(piece: str, scheme: str) -> dict[str, object]:
    paint = _paint(piece)
    if paint is None:
        return {"font_solid": True}
    found: dict[str, object] = {}
    index = _INDEX.search(piece)
    if index is not None:
        found["font_index"] = int(index.group(1))
    if paint.rgb is not None:
        found["font_rgb"] = paint.rgb
    elif paint.scheme != scheme:
        found["font_theme"] = paint.scheme
    elif not paint.mods and paint.alpha is None:
        found["font_solid"] = True
    if len(paint.mods) == 1 and paint.mods[0][0] == "tint":
        found["font_tint"] = paint.mods[0][1]
    elif paint.mods:
        found["font_mods"] = paint.mods
    if paint.alpha is not None:
        found["font_alpha"] = paint.alpha
    return found


def typeface(markup: str | None) -> str | None:
    """The name a latin, ea, cs or sym element gives."""
    found = None if markup is None else _TYPEFACE.search(markup)
    return None if found is None else unescape(found.group(1))


def _frame(look: TextLook, attributes: str, inside: str) -> None:
    """What bodyPr says."""
    margins = [_attribute(attributes, name) for name in ("lIns", "tIns", "rIns", "bIns")]
    look.margins = [None if value is None else int(value) / 12700 for value in margins]
    look.anchor = _attribute(attributes, "anchor") or "t"
    look.vert = _attribute(attributes, "vert")
    look.horizontal_overflow = _attribute(attributes, "horzOverflow") or "overflow"
    look.vertical_overflow = _attribute(attributes, "vertOverflow") or "overflow"
    look.wrap = _attribute(attributes, "wrap")
    centred = _attribute(attributes, "anchorCtr")
    look.anchor_center = None if centred is None else centred == "1"
    look.upright = _attribute(attributes, "upright") == "1"
    columns, gap = _attribute(attributes, "numCol"), _attribute(attributes, "spcCol")
    look.columns = 1 if columns is None else int(columns)
    look.column_spacing = 0.0 if gap is None else int(gap) / 12700
    look.full = _attribute(attributes, "rot") is not None
    scale = re.search(r'<a:normAutofit\b[^>]*\bfontScale="(\d+)"', inside)
    look.font_scale = None if scale is None else int(scale.group(1))
    warp = re.search(r'<a:prstTxWarp\b[^>]*\bprst="(\w+)"', inside)
    look.warp = None if warp is None or warp.group(1) == "textNoShape" else warp.group(1)
    look.autofit = "no" if "<a:noAutofit" in inside else "norm" if "<a:normAutofit" in inside else "sp" if "<a:spAutoFit" in inside else None


def read_text(markup: str) -> ReadText:
    """All a shape's markup holds of its text. With no txBody, what a new shape of no text reads as."""
    style = dml.element_span(markup, "xdr:style")
    reference = None if style is None else re.search(r"<a:fontRef\b[^>]*>(.*?)</a:fontRef>", markup[style[0]:style[1]], re.DOTALL)
    paint = None if reference is None else _paint(reference.group(1))
    look = TextLook(font_scheme=paint.scheme if paint is not None and paint.scheme else "tx1")
    if paint is not None and paint.rgb is not None:
        look.font_rgb = paint.rgb
    body = _body(markup)
    if body is None:
        return ReadText(look)
    head = re.search(r"<a:bodyPr\b([^>]*?)(?:/>|>(.*?)</a:bodyPr>)", body, re.DOTALL)
    if head is not None:
        _frame(look, head.group(1), head.group(2) or "")
    characters: list[dict[str, object]] = []
    ends: list[dict[str, object] | None] = []
    formats: list[str] = []
    listed = re.search(r"<a:lstStyle>(.*?)</a:lstStyle>", body, re.DOTALL)
    for paragraph in _PARAGRAPH.finditer(body):
        inside = paragraph.group(1) or ""
        properties = _PROPERTIES.search(inside)
        own = "" if properties is None else properties.group(0)
        level = int(paragraph_settings(own).get("lvl", 0))
        level_style = None if listed is None else re.search(rf"<a:lvl{level + 1}pPr\b([^>]*?)(?:/>|>(.*?)</a:lvl{level + 1}pPr>)", listed.group(1), re.DOTALL)
        default = None if level_style is None else re.search(r"<a:defRPr\b([^>]*?)(?:/>|>(.*?)</a:defRPr>)", level_style.group(2) or "", re.DOTALL)
        inherited = {} if default is None else _settings(default.group(1), default.group(2), look.font_scheme)
        if level_style is not None:
            # What a level's style says of a paragraph, under what the paragraph says.
            own = _under(own, level_style.group(1))
        formats.append(own)
        end: dict[str, object] | None = None
        last: dict[str, object] = {}
        for piece in _PIECE.finditer(inside):
            whole = piece.group(0)
            if whole.startswith("<a:endParaRPr"):
                end = {**inherited, **_settings(piece.group(7), piece.group(8), look.font_scheme)}
            elif whole.startswith("<a:br"):
                last = {**inherited, **_settings(piece.group(5), piece.group(6), look.font_scheme)}
                characters.append(last)
            else:
                last = {**inherited, **_settings(piece.group(2), piece.group(3), look.font_scheme)}
                characters.extend(dict(last) for _ in _run_text(piece.group(4)))
        ends.append(end)
        # A paragraph's ending takes its end's settings, or those of its last run.
        characters.append(end if end is not None else last)
    if characters:
        characters.pop()
    found = ReadText(look, end=ends[-1] if ends else None)
    if not characters and found.end is not None:
        _apply(look, found.end)
        found.end = None
    elif characters and all(one == characters[0] for one in characters):
        base = characters[0]
        _apply(look, base)
        if found.end is not None:
            found.end = _against(found.end, base)
    else:
        found.marks = characters
    _layout(found, formats)
    return found


def _under(own: str, attributes: str) -> str:
    """A pPr with the attributes a level's style gives where it says nothing itself."""
    said = paragraph_settings(own)
    extra = "".join(f' {name}="{value}"' for name, value in re.findall(r'\b(\w+)="([^"]*)"', attributes) if name not in said)
    if not extra:
        return own
    return re.sub(r"<a:pPr", f"<a:pPr{extra}", own, count=1) if own else f"<a:pPr{extra}/>"


def _apply(look: TextLook, settings: dict[str, object]) -> None:
    for name, value in settings.items():
        setattr(look, name, value)


def _against(settings: dict[str, object], base: dict[str, object]) -> dict[str, object]:
    """Settings as a mark over a look holding ``base``."""
    found = dict(settings)
    for name in base:
        if name not in found:
            found[name] = _DEFAULTS[name]
    return {name: value for name, value in found.items() if base.get(name, _DEFAULTS[name]) != value}


def _layout(found: ReadText, formats: list[str]) -> None:
    """The first paragraph's algn and rtl are the look's; paragraphs that are not all alike, or say more, keep their own pPr."""
    look = found.look
    first = formats[0] if formats else ""
    head = first[:first.find(">") + 1] if first else ""
    look.align = _attribute(head, "algn")
    rtl = _attribute(head, "rtl")
    look.rtl = None if rtl is None else rtl == "1"
    settings = (f' algn="{look.align}"' if look.align else "") + ("" if rtl is None else f' rtl="{rtl}"')
    plain = f"<a:pPr{settings}/>" if settings else ""
    if any(one != plain for one in formats):
        found.paragraphs = formats


def paragraph_settings(markup: str) -> dict[str, str]:
    """The attributes a pPr gives."""
    head = markup[:markup.find(">") + 1] if markup else ""
    return dict(re.findall(r'\b(\w+)="([^"]*)"', head))
