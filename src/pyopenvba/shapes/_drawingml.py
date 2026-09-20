"""The DrawingML that Excel, Word and PowerPoint all write.

A shape's identity, its transform, its geometry and its text are the
same markup in every host; what differs is the element that wraps them
and the part it lives in.  Everything common is here, read and written
against files the three applications saved.

The XML is handled by scanning rather than parsing into a tree: a
shape that nobody edited has to be written back exactly as it came in,
down to the attribute order and the creation-id extensions Office
leaves in it, and the cheapest way to promise that is never to take it
apart.
"""

from __future__ import annotations

import re

from pyopenvba._xml import escape, unescape
from pyopenvba.shapes._values import Shape, emu, points

_CNVPR = re.compile(r"<(?:\w+:)?cNvPr\b([^>]*?)/?>")
_ID = re.compile(r'\bid="(\d+)"')
_NAME = re.compile(r'\bname="([^"]*)"')
_OFF = re.compile(r"<a:off\b[^>]*?\bx=\"(-?\d+)\"[^>]*?\by=\"(-?\d+)\"[^>]*?/?>")
_EXT = re.compile(r"<a:ext\b[^>]*?\bcx=\"(-?\d+)\"[^>]*?\bcy=\"(-?\d+)\"[^>]*?/?>")
_PRST = re.compile(r"<a:prstGeom\b[^>]*?\bprst=\"([^\"]*)\"")
_RUN = re.compile(r"<a:t>(.*?)</a:t>|<a:br\b[^>]*?/?>|<a:p\b[^>]*?>", re.DOTALL)


def element_span(text: str, tag: str, start: int = 0) -> tuple[int, int] | None:
    """Where one element begins and ends, nesting and all.

    Returns the offsets of the whole element, ``<tag ...>...</tag>`` or
    ``<tag/>``, or None when there is no such element from ``start``.
    """
    opening = re.compile(rf"<{re.escape(tag)}(?=[\s/>])")
    found = opening.search(text, start)
    if found is None:
        return None
    head = text.find(">", found.start())
    if head < 0:
        return None
    if text[head - 1] == "/":
        return found.start(), head + 1
    depth = 1
    at = head + 1
    step = re.compile(rf"<(/?){re.escape(tag)}(?=[\s/>])")
    while depth:
        nxt = step.search(text, at)
        if nxt is None:
            return None
        closing = text.find(">", nxt.start())
        if closing < 0:
            return None
        if nxt.group(1):
            depth -= 1
        elif text[closing - 1] != "/":
            depth += 1
        at = closing + 1
    return found.start(), at


def children_spans(text: str, tags: tuple[str, ...]) -> list[tuple[str, int, int]]:
    """Every top-level child of ``text`` whose tag is one of ``tags``.

    Top-level means not inside another of the same tags, which is what
    keeps a group's members out of a sheet's own list.
    """
    out: list[tuple[str, int, int]] = []
    at = 0
    while at < len(text):
        best: tuple[str, int, int] | None = None
        for tag in tags:
            span = element_span(text, tag, at)
            if span is None:
                continue
            if best is None or span[0] < best[1]:
                best = (tag, span[0], span[1])
        if best is None:
            break
        out.append(best)
        at = best[2]
    return out


def identity(markup: str) -> tuple[int, str]:
    """The id and name a shape carries, from its first cNvPr."""
    found = _CNVPR.search(markup)
    if found is None:
        return 0, ""
    attributes = found.group(1)
    number = _ID.search(attributes)
    name = _NAME.search(attributes)
    return (int(number.group(1)) if number else 0, unescape(name.group(1)) if name else "")


def transform(markup: str) -> tuple[float, float, float, float]:
    """Where a shape sits and how big it is, in points.

    The offset and extent of the shape's own ``a:xfrm``.  A shape with
    no transform of its own -- a worksheet shape written only against
    its cell anchor -- answers zeros, and the host supplies the place.
    """
    off = _OFF.search(markup)
    ext = _EXT.search(markup)
    left = points(off.group(1)) if off else 0.0
    top = points(off.group(2)) if off else 0.0
    width = points(ext.group(1)) if ext else 0.0
    height = points(ext.group(2)) if ext else 0.0
    return left, top, width, height


def geometry(markup: str) -> str:
    """The preset geometry, such as rect, roundRect or ellipse."""
    found = _PRST.search(markup)
    return found.group(1) if found else ""


def text_of(markup: str) -> str:
    """A shape's text, with a paragraph and a break both as newlines.

    Office answers the text of a shape with the paragraphs joined by
    the host's own line ending; the caller settles which one, because
    Excel reports a break as a line feed and Word ends every paragraph
    with one.
    """
    body = element_span(markup, "xdr:txBody") or element_span(markup, "p:txBody")
    body = body or element_span(markup, "a:txBody") or element_span(markup, "wps:txbx")
    if body is None:
        return ""
    inner = markup[body[0] : body[1]]
    pieces: list[str] = []
    first = True
    for found in _RUN.finditer(inner):
        matched = found.group(0)
        if matched.startswith("<a:t>"):
            pieces.append(unescape(found.group(1) or ""))
        elif matched.startswith("<a:br"):
            pieces.append("\n")
        elif not first:
            pieces.append("\n")
        if matched.startswith("<a:p"):
            first = False
    # Excel stores a line break inside a run as CRLF and reports it
    # back as one character, so the text a caller sees is one too.
    joined = "".join(pieces)
    return joined.replace("\r\n", "\n").replace("\r", "\n")


def with_transform(markup: str, shape: Shape) -> str:
    """The same markup with the shape's place and size written into it."""
    out = _OFF.sub(f'<a:off x="{emu(shape.left)}" y="{emu(shape.top)}"/>', markup, count=1)
    return _EXT.sub(f'<a:ext cx="{emu(shape.width)}" cy="{emu(shape.height)}"/>', out, count=1)


def with_name(markup: str, name: str) -> str:
    """The same markup renamed, leaving everything else as it was."""
    found = _CNVPR.search(markup)
    if found is None:
        return markup
    attributes = found.group(1)
    if _NAME.search(attributes) is None:
        return markup
    changed = _NAME.sub(f'name="{escape(name)}"', attributes, count=1)
    return markup[: found.start(1)] + changed + markup[found.end(1) :]


def text_body(text: str, tag: str, *, size: int = 1100) -> str:
    """A txBody holding this text, one paragraph per line.

    Written the way the host writes one: Excel and PowerPoint use the
    same runs, so only the element's own prefix differs.
    """
    lines = text.split("\n") if text else [""]
    paragraphs: list[str] = []
    for line in lines:
        if line:
            run = f'<a:r><a:rPr lang="en-US" sz="{size}"/><a:t>{escape(line)}</a:t></a:r>'
        else:
            run = f'<a:endParaRPr lang="en-US" sz="{size}"/>'
        paragraphs.append(f"<a:p>{run}</a:p>")
    return (
        f"<{tag}><a:bodyPr vertOverflow=\"clip\" horzOverflow=\"clip\" rtlCol=\"0\" anchor=\"t\"/>"
        f"<a:lstStyle/>{''.join(paragraphs)}</{tag}>"
    )


def replace_text(markup: str, text: str, tag: str) -> str:
    """The same markup carrying different text."""
    span = element_span(markup, tag)
    body = text_body(text, tag)
    if span is None:
        return markup
    return markup[: span[0]] + body + markup[span[1] :]
