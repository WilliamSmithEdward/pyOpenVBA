"""The fill and line a shape in a file holds, as the FillLine a new shape's Fill and Line work from, and what its style gives.

Measured in live Excel (scripts/measure_object_roundtrip.py, tests/fixtures/object_roundtrip.json). A shape's own spPr wins over its
style's fillRef and lnRef, whose colours come from the theme and whose line width comes from the theme's list. Gradients, pictures
and custom dashes are not read: a macro reaching them is told so, rather than given a wrong answer.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.shapes import _drawingml as dml
from pyopenvba.shapes._values import ConnectorLink, FillLine, Shape

_COLOUR = re.compile(r'<a:(srgbClr|schemeClr|sysClr|prstClr|scrgbClr|hslClr)\b([^>]*?)(?:/>|>(.*?)</a:\1>)', re.DOTALL)
_MOD = re.compile(r'<a:(\w+)\b[^>]*?\bval="(-?\d+)"')
_REF = re.compile(r'<a:(fillRef|lnRef)\b[^>]*?\bidx="(\d+)"[^>]*?(?:/>|>(.*?)</a:\1>)', re.DOTALL)
_FILLS = ("a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill")
_GEOMETRIES = ("a:prstGeom", "a:custGeom")


@dataclass(frozen=True)
class Paint:
    """A colour in a file: its RRGGBB or theme name, its lumMod and the rest in order, and its alpha."""

    rgb: str | None = None
    scheme: str | None = None
    mods: tuple[tuple[str, int], ...] = ()
    alpha: int | None = None


@dataclass(frozen=True)
class StyleRefs:
    """What a shape's xdr:style gives: each paint, None for an index of 0 or no style, and the line's place in the theme's list."""

    fill: Paint | None = None
    line: Paint | None = None
    line_index: int = 0


def _attribute(attributes: str, name: str) -> str | None:
    found = re.search(rf'\b{name}="([^"]*)"', attributes)
    return found.group(1) if found else None


def _paint(markup: str) -> Paint | None:
    found = _COLOUR.search(markup)
    if found is None:
        return None
    kind, attributes, inside = found.group(1), found.group(2), found.group(3) or ""
    mods = [(name, int(value)) for name, value in _MOD.findall(inside)]
    alpha = next((value for name, value in mods if name == "alpha"), None)
    kept = tuple((name, value) for name, value in mods if name != "alpha")
    if kind == "srgbClr":
        return Paint(rgb=(_attribute(attributes, "val") or "000000").upper(), mods=kept, alpha=alpha)
    if kind == "schemeClr":
        return Paint(scheme=_attribute(attributes, "val"), mods=kept, alpha=alpha)
    if kind == "sysClr":
        return Paint(rgb=(_attribute(attributes, "lastClr") or "000000").upper(), mods=kept, alpha=alpha)
    raise VBAUnsupportedError(f"The fill and line of a shape coloured by {kind} are not implemented")


def _top(markup: str, tags: tuple[str, ...]) -> tuple[str, int, int] | None:
    """The first of ``tags`` in the markup, and where it is."""
    spans = [(span[0], tag, span[1]) for tag in tags if (span := dml.element_span(markup, tag)) is not None]
    if not spans:
        return None
    start, tag, end = min(spans)
    return tag, start, end


def properties(markup: str) -> str:
    """What the shape's spPr holds after its transform and geometry."""
    span = dml.element_span(markup, "xdr:spPr")
    if span is None:
        return ""
    body = markup[span[0]:span[1]]
    geometry = _top(body, _GEOMETRIES)
    transform = dml.element_span(body, "a:xfrm")
    start = geometry[2] if geometry is not None else transform[1] if transform is not None else body.find(">") + 1
    return body[start:]


def read_look(markup: str) -> FillLine:
    """What the shape's own spPr sets."""
    look = FillLine()
    rest = properties(markup)
    line = dml.element_span(rest, "a:ln")
    outside = rest if line is None else rest[:line[0]]
    fill = _top(outside, _FILLS)
    if fill is not None:
        tag, start, end = fill
        piece = outside[start:end]
        if tag == "a:noFill":
            look.fill_hidden = True
        elif tag == "a:solidFill":
            paint = _paint(piece)
            if paint is not None:
                look.fill_rgb, look.fill_scheme, look.fill_mods, look.fill_alpha = paint.rgb, paint.scheme, paint.mods, paint.alpha
                look.fill_solid = paint.rgb is None and paint.scheme is None
        elif tag == "a:pattFill":
            _read_pattern(look, piece)
        else:
            raise VBAUnsupportedError(f"The fill and line of a shape with {tag[2:]} are not implemented")
    if line is not None:
        _read_line(look, rest[line[0]:line[1]])
    return look


def _read_pattern(look: FillLine, piece: str) -> None:
    look.fill_pattern = _attribute(piece[:piece.find(">")], "prst")
    fore, back = dml.element_span(piece, "a:fgClr"), dml.element_span(piece, "a:bgClr")
    paint = None if fore is None else _paint(piece[fore[0]:fore[1]])
    if paint is not None:
        look.fill_rgb, look.fill_scheme, look.fill_mods = paint.rgb, paint.scheme, paint.mods
    paint = None if back is None else _paint(piece[back[0]:back[1]])
    if paint is not None:
        look.fill_back_rgb, look.fill_back_scheme, look.fill_back_mods = paint.rgb, paint.scheme, paint.mods


def _read_line(look: FillLine, piece: str) -> None:
    head = piece[:piece.find(">")]
    width = _attribute(head, "w")
    look.line_width = None if width is None else int(width)
    look.line_cap, look.line_align = _attribute(head, "cap"), _attribute(head, "algn")
    if "<a:custDash" in piece:
        raise VBAUnsupportedError("The line of a shape with a custom dash is not implemented")
    if "<a:gradFill" in piece:
        raise VBAUnsupportedError("The line of a shape with a gradient is not implemented")
    if "<a:noFill" in piece:
        look.line_hidden = True
    solid = dml.element_span(piece, "a:solidFill")
    paint = None if solid is None else _paint(piece[solid[0]:solid[1]])
    if paint is not None:
        look.line_rgb, look.line_scheme, look.line_mods, look.line_alpha = paint.rgb, paint.scheme, paint.mods, paint.alpha
    dash = re.search(r'<a:prstDash\b[^>]*\bval="([^"]*)"', piece)
    look.line_dash = None if dash is None or dash.group(1) == "solid" else dash.group(1)
    for tag, field in (("a:headEnd", "line_head"), ("a:tailEnd", "line_tail")):
        found = re.search(rf"<{tag}\b([^>]*)/?>", piece)
        if found is not None:
            setattr(look, field, (_attribute(found.group(1), "type"), _attribute(found.group(1), "w"), _attribute(found.group(1), "len")))


def read_style(markup: str) -> StyleRefs:
    """What the shape's xdr:style gives."""
    span = dml.element_span(markup, "xdr:style")
    if span is None:
        return StyleRefs()
    fill: Paint | None = None
    line: Paint | None = None
    line_index = 0
    for name, index, inside in _REF.findall(markup[span[0]:span[1]]):
        if not int(index):
            continue  # its colour is not used
        if name == "fillRef":
            if int(index) > 1:
                raise VBAUnsupportedError("The fill of a shape styled by a theme's gradient is not implemented")
            fill = _paint(inside)
        else:
            line, line_index = _paint(inside), int(index)
    return StyleRefs(fill, line, line_index)


def read_links(markup: str) -> tuple[ConnectorLink | None, ConnectorLink | None]:
    """The shape and site each end of a connector in a file holds, with where the site stands not yet worked out."""
    found: list[ConnectorLink | None] = []
    for tag in ("a:stCxn", "a:endCxn"):
        link = re.search(rf"<{tag}\b([^>]*)/>", markup)
        found.append(None if link is None else ConnectorLink(int(_attribute(link.group(1), "id") or 0), int(_attribute(link.group(1), "idx") or 0) + 1, (0.0, 0.0), ""))
    return found[0], found[1]


def has_own_shadow(markup: str, theme: str) -> bool | None:
    """Whether a shape has a shadow, by its own effect list or else the theme effect its style names; None for a theme with no
    such effect to look at."""
    rest = properties(markup)
    if "<a:effectLst" in rest:
        return "<a:outerShdw" in rest
    reference = re.search(r'<a:effectRef\b[^>]*?\bidx="(\d+)"', markup)
    if reference is None or reference.group(1) == "0":
        return False
    effects = re.findall(r"<a:effectStyle>.*?</a:effectStyle>", theme, re.DOTALL)
    index = int(reference.group(1))
    return "<a:outerShdw" in effects[index - 1] if 0 < index <= len(effects) else None


def read_picture(markup: str, shape: Shape) -> None:
    """What PictureFormat reads of a picture in a file: its blip's effects, its crop, and the fill behind it."""
    blip = re.search(r"<a:blip\b[^>]*/>|<a:blip\b.*?</a:blip>", markup, re.DOTALL)
    effects = "" if blip is None else blip.group(0)
    lum = re.search(r"<a:lum\b([^>]*)/>", effects)
    if lum is not None:
        shape.picture_lum = (int(_attribute(lum.group(1), "bright") or 0), int(_attribute(lum.group(1), "contrast") or 0))
    shape.picture_colour_type = 2 if "<a:grayscl" in effects else 3 if "<a:biLevel" in effects else 4 if shape.picture_lum == (70000, -70000) else 1
    clear = re.search(r'<a:clrFrom>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"', effects)
    if clear is not None:
        shape.picture_clear = clear.group(1).upper()
    rect = re.search(r"<a:srcRect\b([^>]*)/>", markup)
    if rect is not None:
        left, top, right, bottom = (int(_attribute(rect.group(1), name) or 0) for name in ("l", "t", "r", "b"))
        shape.picture_crop = (left, top, right, bottom)
    fill = _top(properties(markup), _FILLS)
    if fill is not None and fill[0] == "a:solidFill":
        paint = _paint(properties(markup)[fill[1]:fill[2]])
        shape.picture_transparent = paint is not None and paint.alpha == 0


def with_fill_line(markup: str, written: str) -> str:
    """The shape's spPr with its fill and line, and its extLst if there is a new one, as ``written``, and what else it holds kept."""
    span = dml.element_span(markup, "xdr:spPr")
    if span is None:
        return markup
    rest = properties(markup)
    start = span[1] - len(rest)
    line = dml.element_span(rest, "a:ln")
    fill = _top(rest if line is None else rest[:line[0]], _FILLS)
    if fill is not None:
        rest = rest[:fill[1]] + rest[fill[2]:]
    line = dml.element_span(rest, "a:ln")
    if line is not None:
        rest = rest[:line[0]] + rest[line[1]:]
    extensions = ""
    found = written.find("<a:extLst>")
    if found >= 0:
        written, extensions = written[:found], written[found:]
    if "<a:effectLst" in rest:
        written = written.replace("<a:effectLst/>", "")
    close = rest.rfind("</xdr:spPr>")
    return markup[:start] + written + rest[:close] + extensions + rest[close:] + markup[span[1]:]
