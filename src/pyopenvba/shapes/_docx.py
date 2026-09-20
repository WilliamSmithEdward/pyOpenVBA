"""A document's shapes.

Word keeps a shape in the text itself, in ``word/document.xml``, inside
the run that holds it.  A floating one is a ``wp:anchor`` and an
in-line one is a ``wp:inline``, both inside a ``w:drawing``, and Word
wraps each in an ``mc:AlternateContent`` whose fallback is the VML an
old reader would use.  A horizontal line and other legacy pictures are
that VML on their own, in a ``w:pict``.

Where a shape sits is not a transform: an anchored one carries
``wp:positionH`` and ``wp:positionV`` offsets, in EMU, from whatever
each is relative to -- the column and the paragraph, by default -- and
its size is ``wp:extent``.  That is what Word reports as Left and Top,
negative numbers included.

Word has no macro on a shape: there is nothing here for OnAction,
which is why the field stays empty rather than being given a default.

Every layout here was read from a document Word saved, in
``tests/fixtures/shapes``.
"""

from __future__ import annotations

import re

from pyopenvba._xml import escape, unescape
from pyopenvba.shapes import _drawingml as dml
from pyopenvba.shapes._values import Shape, emu, points

#: Where a shape can be: floating beside the text, or in it.
_ANCHOR = "wp:anchor"
_INLINE = "wp:inline"

_DOCPR = re.compile(r"<wp:docPr\b([^>]*?)/?>")
_ID = re.compile(r'\bid="(\d+)"')
_NAME = re.compile(r'\bname="([^"]*)"')
_EXTENT = re.compile(r"<wp:extent\b[^>]*?\bcx=\"(-?\d+)\"[^>]*?\bcy=\"(-?\d+)\"[^>]*?/?>")
_POSITION_H = re.compile(r"<wp:positionH\b.*?</wp:positionH>", re.DOTALL)
_POSITION_V = re.compile(r"<wp:positionV\b.*?</wp:positionV>", re.DOTALL)
_OFFSET = re.compile(r"<wp:posOffset>(-?\d+)</wp:posOffset>")
_PICT = re.compile(r"<w:pict\b[^>]*?>.*?</w:pict>", re.DOTALL)
#: The z-order Word gives an anchored shape.  It rises as shapes are
#: added, and it is the order Document.Shapes lists them in -- not
#: the order of the runs, which holds the newest first.
_RELATIVE_HEIGHT = re.compile(r'\brelativeHeight="(\d+)"')
_VML_SHAPE = re.compile(r"<v:(rect|oval|shape|line|roundrect)\b([^>]*?)(?:/>|>)")


def read_document(xml: str) -> list[Shape]:
    """Every shape in a document, in the order the text holds them.

    Anchored and in-line shapes come back in one list, each saying
    which it is, because Word keeps them in two collections but one
    order.
    """
    body = dml.element_span(xml, "w:body")
    inner = xml[body[0] : body[1]] if body else xml
    out: list[Shape] = []
    for placement, tag in (("anchored", _ANCHOR), ("inline", _INLINE)):
        at = 0
        while True:
            span = dml.element_span(inner, tag, at)
            if span is None:
                break
            at = span[1]
            shape = _shape_of(inner[span[0] : span[1]], placement)
            shape.source_at = span[0]
            out.append(shape)
    out.extend(_legacy_pictures(inner))
    # Word lists its floating shapes by z-order, which rises as each is
    # added, and its in-line ones in the order the text holds them.
    floating = sorted((one for one in out if one.placement == "anchored"), key=_z_of)
    inline = sorted((one for one in out if one.placement != "anchored"), key=lambda one: one.source_at)
    return [*floating, *inline]


def _z_of(shape: Shape) -> tuple[int, int]:
    return (shape.z_order or shape.source_at, shape.source_at)


def _shape_of(markup: str, placement: str) -> Shape:
    shape_id, name = _identity(markup)
    width, height = _size(markup)
    left, top = _position(markup) if placement == "anchored" else (0.0, 0.0)
    geometry = dml.geometry(markup)
    kind = "line" if geometry == "line" else "shape"
    if _is_picture(markup):
        kind = "picture"
    elif _is_text_box(markup):
        kind = "textBox"
    height_above = _RELATIVE_HEIGHT.search(markup)
    shape = Shape(
        name=name,
        kind=kind,
        geometry=geometry,
        left=left,
        top=top,
        width=width,
        height=height,
        text=_text_of(markup),
        shape_id=shape_id,
        placement=placement,
        source=markup,
    )
    shape.z_order = int(height_above.group(1)) if height_above else 0
    return shape


def _legacy_pictures(inner: str) -> list[Shape]:
    """The VML shapes Word writes on their own, such as a rule.

    Word reports one as an InlineShape with no name of its own, so the
    name is the one the object model gives it: its position in the
    collection, which the caller fills in.
    """
    out: list[Shape] = []
    for found in _PICT.finditer(inner):
        markup = found.group(0)
        if _inside_fallback(inner, found.start()):
            # The VML copy of a DrawingML shape, for readers that have
            # no DrawingML.  It is the same shape, not another one.
            continue
        element = _VML_SHAPE.search(markup)
        if element is None:
            continue
        style = _vml_style(element.group(2))
        rule = 'o:hr="t"' in element.group(2)
        shape = Shape(
            name="",
            kind="horizontalLine" if rule else "picture",
            geometry=element.group(1),
            # Word reports a rule's width as -1: it is as wide as the
            # text, not a measurement of its own.
            width=-1.0 if rule else _style_points(style.get("width", "0")),
            height=_style_points(style.get("height", "0")),
            placement="inline",
            source=markup,
        )
        shape.source_at = found.start()
        out.append(shape)
    return out


def _inside_fallback(inner: str, at: int) -> bool:
    """Whether this markup is the fallback half of an AlternateContent."""
    opened = inner.rfind("<mc:Fallback>", 0, at)
    if opened < 0:
        return False
    return inner.find("</mc:Fallback>", opened) > at


def _vml_style(attributes: str) -> dict[str, str]:
    """The style attribute of a VML shape, as its own pairs."""
    found = re.search(r'style="([^"]*)"', attributes)
    if found is None:
        return {}
    out: dict[str, str] = {}
    for piece in found.group(1).split(";"):
        name, _, value = piece.partition(":")
        if value:
            out[name.strip()] = value.strip()
    return out


def _style_points(value: str) -> float:
    """A VML style length, which Word writes in points."""
    try:
        return float(value.rstrip("pt"))
    except ValueError:
        return 0.0


def _identity(markup: str) -> tuple[int, str]:
    found = _DOCPR.search(markup)
    if found is None:
        return dml.identity(markup)
    number = _ID.search(found.group(1))
    name = _NAME.search(found.group(1))
    return (int(number.group(1)) if number else 0, unescape(name.group(1)) if name else "")


def _size(markup: str) -> tuple[float, float]:
    found = _EXTENT.search(markup)
    if found is None:
        return 0.0, 0.0
    return points(found.group(1)), points(found.group(2))


def _position(markup: str) -> tuple[float, float]:
    """Where an anchored shape sits, as Word reports Left and Top."""
    across = _POSITION_H.search(markup)
    down = _POSITION_V.search(markup)
    left = _offset_of(across.group(0)) if across else 0.0
    top = _offset_of(down.group(0)) if down else 0.0
    return left, top


def _offset_of(markup: str) -> float:
    found = _OFFSET.search(markup)
    return points(found.group(1)) if found else 0.0


def _text_of(markup: str) -> str:
    """A shape's text, which Word writes in its own runs.

    Inside ``w:txbxContent`` the text is WordprocessingML, not
    DrawingML: ``w:p`` paragraphs of ``w:r`` runs holding ``w:t``, with
    ``w:br`` for a line break.
    """
    body = dml.element_span(markup, "w:txbxContent")
    if body is None:
        return dml.text_of(markup)
    inner = markup[body[0] : body[1]]
    pieces: list[str] = []
    first = True
    for found in re.finditer(r"<w:t\b[^>]*>(.*?)</w:t>|<w:br\b[^>]*/?>|<w:p\b[^>]*>", inner, re.DOTALL):
        matched = found.group(0)
        if matched.startswith("<w:t"):
            pieces.append(unescape(found.group(1) or ""))
        elif matched.startswith("<w:br"):
            pieces.append("\n")
        elif not first:
            pieces.append("\n")
        if matched.startswith("<w:p"):
            first = False
    return "".join(pieces)


def _is_text_box(markup: str) -> bool:
    """A text box says so on its own cNvSpPr, as in the other hosts.

    Having a text box inside is not the test: Word puts one in every
    AutoShape that carries text, and still calls the shape an
    AutoShape.
    """
    return re.search(r'<wps:cNvSpPr\b[^>]*?\btxBox="1"', markup) is not None


def _is_picture(markup: str) -> bool:
    return "<pic:pic" in markup


def written(shapes: list[Shape], original: str) -> str:
    """A document holding these shapes.

    Each one is written back in place, because a shape in Word belongs
    to the run that holds it: moving it in the list would move it in
    the text.  A shape that was deleted takes its run with it.
    """
    out = original
    fresh: list[Shape] = []
    for shape in shapes:
        if not shape.source:
            fresh.append(shape)
            continue
        markup = _markup_for(shape)
        if markup != shape.source:
            out = out.replace(shape.source, markup, 1)
    for shape in fresh:
        out = _with_new_shape(out, shape)
    return out


def _with_new_shape(original: str, shape: Shape) -> str:
    """The document with one more shape in it.

    A shape belongs to a run, and a run to a paragraph, so a new one
    goes into the first paragraph of the body -- which is where Word
    anchors a shape a macro adds without saying where.
    """
    body = dml.element_span(original, "w:body")
    if body is None:
        return original
    inner = original[body[0] : body[1]]
    paragraph = _first_paragraph(inner)
    if paragraph is None:
        head = inner.find(">") + 1
        rebuilt = inner[:head] + f"<w:p>{new_run(shape)}</w:p>" + inner[head:]
    else:
        start, end = paragraph
        head = inner.find(">", start) + 1
        rebuilt = inner[:head] + new_run(shape) + inner[head:]
        del end
    return original[: body[0]] + rebuilt + original[body[1] :]


def _first_paragraph(inner: str) -> tuple[int, int] | None:
    span = dml.element_span(inner, "w:p")
    return span


def new_run(shape: Shape) -> str:
    """The run a new shape lives in, written the way Word writes one."""
    geometry = shape.geometry or "rect"
    text = _text_content(shape.text) if shape.text or shape.kind == "textBox" else ""
    box = f"<wps:txbx>{text}</wps:txbx>" if text else ""
    box_mark = ' txBox="1"' if shape.kind == "textBox" else ""
    return (
        "<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>"
        '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0"'
        f' relativeHeight="{251659264 + shape.z_order}" behindDoc="0" locked="0"'
        ' layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="column"><wp:posOffset>{emu(shape.left)}</wp:posOffset>'
        "</wp:positionH>"
        f'<wp:positionV relativeFrom="paragraph"><wp:posOffset>{emu(shape.top)}</wp:posOffset>'
        "</wp:positionV>"
        f'<wp:extent cx="{emu(shape.width)}" cy="{emu(shape.height)}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        f'<wp:docPr id="{shape.shape_id}" name="{escape(shape.name)}"/>'
        "<wp:cNvGraphicFramePr/>"
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        f"<wps:cNvSpPr{box_mark}/>"
        "<wps:spPr><a:xfrm>"
        '<a:off x="0" y="0"/>'
        f'<a:ext cx="{emu(shape.width)}" cy="{emu(shape.height)}"/>'
        f'</a:xfrm><a:prstGeom prst="{geometry}"><a:avLst/></a:prstGeom>'
        "</wps:spPr>"
        f"{box}"
        '<wps:bodyPr rtlCol="0" anchor="ctr"/>'
        "</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"
    )


def _markup_for(shape: Shape) -> str:
    was = _shape_of(shape.source, shape.placement or "anchored")
    markup = shape.source
    if shape.name != was.name:
        markup = _with_name(markup, shape.name)
    if (shape.width, shape.height) != (was.width, was.height):
        markup = _with_size(markup, shape)
    if shape.placement == "anchored" and (shape.left, shape.top) != (was.left, was.top):
        markup = _with_position(markup, shape)
    if shape.text != was.text:
        markup = _with_text(markup, shape.text)
    return markup


def _with_text(markup: str, text: str) -> str:
    """The same shape carrying different text.

    Word's text box holds WordprocessingML, so the paragraphs written
    here are ``w:p`` of ``w:r`` of ``w:t``, not DrawingML runs.
    """
    span = dml.element_span(markup, "w:txbxContent")
    if span is None:
        return markup
    return markup[: span[0]] + _text_content(text) + markup[span[1] :]


def _text_content(text: str) -> str:
    """A ``w:txbxContent`` holding this text, one paragraph per line."""
    paragraphs: list[str] = []
    for line in (text.split("\n") if text else [""]):
        run = f'<w:r><w:t xml:space="preserve">{escape(line)}</w:t></w:r>' if line else ""
        paragraphs.append(f"<w:p>{run}</w:p>")
    return f"<w:txbxContent>{''.join(paragraphs)}</w:txbxContent>"


def _with_name(markup: str, name: str) -> str:
    found = _DOCPR.search(markup)
    if found is None:
        return markup
    changed = _NAME.sub(f'name="{escape(name)}"', found.group(1), count=1)
    return markup[: found.start(1)] + changed + markup[found.end(1) :]


def _with_size(markup: str, shape: Shape) -> str:
    """Both places the size is written: the extent and the transform."""
    out = _EXTENT.sub(
        f'<wp:extent cx="{emu(shape.width)}" cy="{emu(shape.height)}"/>', markup, count=1
    )
    return dml.with_transform(out, Shape(left=0.0, top=0.0, width=shape.width, height=shape.height))


def _with_position(markup: str, shape: Shape) -> str:
    out = markup
    across = _POSITION_H.search(out)
    if across is not None:
        moved = _OFFSET.sub(f"<wp:posOffset>{emu(shape.left)}</wp:posOffset>", across.group(0), 1)
        out = out[: across.start()] + moved + out[across.end() :]
    down = _POSITION_V.search(out)
    if down is not None:
        moved = _OFFSET.sub(f"<wp:posOffset>{emu(shape.top)}</wp:posOffset>", down.group(0), 1)
        out = out[: down.start()] + moved + out[down.end() :]
    return out


def without(shape: Shape, original: str) -> str:
    """The document with one shape's whole run taken out."""
    at = original.find(shape.source)
    if at < 0:
        return original
    start = original.rfind("<w:r>", 0, at)
    start = original.rfind("<w:r ", 0, at) if start < 0 else start
    if start < 0:
        return original.replace(shape.source, "", 1)
    finish = original.find("</w:r>", at)
    if finish < 0:
        return original.replace(shape.source, "", 1)
    return original[:start] + original[finish + len("</w:r>") :]
