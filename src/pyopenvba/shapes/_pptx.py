"""A slide's shapes.

A slide is one part, ``ppt/slides/slideN.xml``, and its shapes are the
children of its ``p:spTree``.  There is no anchor grid as on a
worksheet: a shape's place is EMU from the slide's top-left corner, in
its own ``a:xfrm``.

The macro a click runs is not an attribute here either.  It is an
``a:hlinkClick`` inside the shape's ``p:cNvPr`` whose action is
``ppaction://macro?name=Proc``, with an empty ``r:id``, because the
relationship an ordinary hyperlink needs does not exist for a macro.

The order slides are shown in is ``p:sldIdLst`` in
``ppt/presentation.xml``, not the numbering of the slide parts, which
does not change when slides are reordered or deleted.

Every layout here was read from a presentation PowerPoint saved, in
``tests/fixtures/shapes``.
"""

from __future__ import annotations

import re

from pyopenvba._xml import escape, unescape
from pyopenvba.shapes import _drawingml as dml
from pyopenvba.shapes._values import Shape, emu

#: What each element on a slide is, as the object model names it.
_ELEMENTS: dict[str, str] = {
    "p:sp": "shape",
    "p:cxnSp": "line",
    "p:pic": "picture",
    "p:graphicFrame": "chart",
    "p:grpSp": "group",
}

_MACRO_ACTION = re.compile(r'<a:hlinkClick\b[^>]*?\baction="ppaction://macro\?name=([^"]*)"[^>]*?/?>')
_HLINK = re.compile(r"<a:hlinkClick\b[^>]*?/>")
_PLACEHOLDER = re.compile(r"<p:ph\b[^>]*?/?>")
_CNVPR_OPEN = re.compile(r"<p:cNvPr\b[^>]*?>")
_CNVPR_EMPTY = re.compile(r"<p:cNvPr\b[^>]*?/>")

#: The order the slides are shown in.
_SLIDE_ID = re.compile(r"<p:sldId\b[^>]*?\br:id=\"([^\"]+)\"[^>]*?/?>")


def slide_order(presentation_xml: str) -> list[str]:
    """The relationship ids of the slides, in the order they are shown."""
    return _SLIDE_ID.findall(presentation_xml)


def read_slide(xml: str, layout_xml: str = "") -> list[Shape]:
    """Every shape on a slide, in the order the tree holds them.

    A placeholder usually carries no transform of its own: PowerPoint
    reports where the layout puts it, so the layout is read too and a
    placeholder with no ``a:xfrm`` takes the one from the layout's
    placeholder of the same type.
    """
    tree = dml.element_span(xml, "p:spTree")
    if tree is None:
        return []
    inner = xml[tree[0] : tree[1]]
    inherited = _layout_boxes(layout_xml) if layout_xml else {}
    out: list[Shape] = []
    for tag, start, end in dml.children_spans(inner, tuple(_ELEMENTS)):
        markup = inner[start:end]
        out.append(_shape_of(markup, tag, inherited))
    return out


def _layout_boxes(layout_xml: str) -> dict[tuple[str, str], tuple[float, float, float, float]]:
    """Where the layout puts each of its placeholders."""
    tree = dml.element_span(layout_xml, "p:spTree")
    if tree is None:
        return {}
    inner = layout_xml[tree[0] : tree[1]]
    out: dict[tuple[str, str], tuple[float, float, float, float]] = {}
    for _, start, end in dml.children_spans(inner, tuple(_ELEMENTS)):
        markup = inner[start:end]
        key = _placeholder_key(markup)
        if key is None:
            continue
        box = dml.transform(markup)
        if box[2] or box[3]:
            out[key] = box
    return out


def _placeholder_key(markup: str) -> tuple[str, str] | None:
    """What a placeholder is, as the type and index it carries."""
    found = _PLACEHOLDER.search(markup)
    if found is None:
        return None
    fields = dict(re.findall(r'(\w+)="([^"]*)"', found.group(0)))
    return fields.get("type", "body"), fields.get("idx", "")


def _shape_of(
    markup: str,
    tag: str,
    inherited: dict[tuple[str, str], tuple[float, float, float, float]] | None = None,
) -> Shape:
    kind = _ELEMENTS[tag]
    shape_id, name = dml.identity(markup)
    left, top, width, height = dml.transform(markup)
    if kind == "shape" and _PLACEHOLDER.search(markup):
        kind = "placeholder"
        if not width and not height:
            key = _placeholder_key(markup)
            box = (inherited or {}).get(key or ("", ""))
            if box is not None:
                left, top, width, height = box
    elif kind == "shape" and _is_text_box(markup):
        kind = "textBox"
    macro = _MACRO_ACTION.search(markup)
    shape = Shape(
        name=name,
        kind=kind,
        geometry=dml.geometry(markup),
        left=left,
        top=top,
        width=width,
        height=height,
        text=dml.text_of(markup),
        macro=unescape(macro.group(1)) if macro else "",
        action=8 if macro else 0,
        shape_id=shape_id,
        source=markup,
    )
    if kind == "group":
        shape.children = _members(markup)
    return shape


def _is_text_box(markup: str) -> bool:
    """A text box says so on its own cNvSpPr."""
    return re.search(r"<p:cNvSpPr\b[^>]*?\btxBox=\"1\"", markup) is not None


def _members(markup: str) -> list[Shape]:
    """A group's own shapes, which sit directly inside it."""
    head = markup.find(">", markup.find("<p:grpSp")) + 1
    inner = markup[head : markup.rfind("</p:grpSp>")]
    return [
        _shape_of(inner[start:end], tag)
        for tag, start, end in dml.children_spans(inner, tuple(_ELEMENTS))
    ]


def written(shapes: list[Shape], original: str) -> str:
    """A slide part holding these shapes.

    A shape nobody touched is written back from the markup it was read
    from, so an untouched slide survives a save byte for byte.
    """
    tree = dml.element_span(original, "p:spTree")
    if tree is None:
        return original
    inner = original[tree[0] : tree[1]]
    head_end = inner.find(">") + 1
    # The tree's own properties come first and are not a shape.
    properties = dml.element_span(inner, "p:nvGrpSpPr")
    group_properties = dml.element_span(inner, "p:grpSpPr")
    keep_to = head_end
    for span in (properties, group_properties):
        if span is not None:
            keep_to = max(keep_to, span[1])
    body = "".join(_markup_for(one) for one in shapes)
    rebuilt = inner[:keep_to] + body + "</p:spTree>"
    return original[: tree[0]] + rebuilt + original[tree[1] :]


def _markup_for(shape: Shape) -> str:
    """One shape's markup: its own if nothing about it changed."""
    if not shape.source:
        return new_shape(shape)
    was = _shape_of(shape.source, _tag_of(shape.source))
    markup = shape.source
    if shape.name != was.name:
        markup = dml.with_name(markup, shape.name)
    if (shape.left, shape.top, shape.width, shape.height) != (
        was.left,
        was.top,
        was.width,
        was.height,
    ):
        markup = dml.with_transform(markup, shape)
    if shape.macro != was.macro:
        markup = with_macro(markup, shape.macro)
    if shape.text != was.text:
        markup = dml.replace_text(markup, shape.text, "p:txBody")
    return markup


def _tag_of(markup: str) -> str:
    for tag in _ELEMENTS:
        if markup.startswith("<" + tag):
            return tag
    return "p:sp"


def with_macro(markup: str, macro: str) -> str:
    """The same shape with the macro a click runs set or cleared.

    The link lives inside ``p:cNvPr`` and has to be its first child,
    which is where PowerPoint writes it.
    """
    without = _HLINK.sub("", markup, count=1) if _MACRO_ACTION.search(markup) else markup
    if not macro:
        return without
    link = f'<a:hlinkClick r:id="" action="ppaction://macro?name={escape(macro)}"/>'
    empty = _CNVPR_EMPTY.search(without)
    if empty is not None:
        head = empty.group(0)[:-2]
        return without[: empty.start()] + f"{head}>{link}</p:cNvPr>" + without[empty.end() :]
    opening = _CNVPR_OPEN.search(without)
    if opening is None:
        return without
    return without[: opening.end()] + link + without[opening.end() :]


def new_shape(shape: Shape) -> str:
    """The markup for a shape this library is adding to a slide."""
    line = shape.kind == "line"
    element = "p:cxnSp" if line else "p:sp"
    link = (
        f'<a:hlinkClick r:id="" action="ppaction://macro?name={escape(shape.macro)}"/>'
        if shape.macro
        else ""
    )
    identity = f'<p:cNvPr id="{shape.shape_id}" name="{escape(shape.name)}">{link}</p:cNvPr>'
    box = ' txBox="1"' if shape.kind == "textBox" else ""
    properties = (
        f"<p:nvCxnSpPr>{identity}<p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>"
        if line
        else f"<p:nvSpPr>{identity}<p:cNvSpPr{box}/><p:nvPr/></p:nvSpPr>"
    )
    body = "" if line else dml.text_body(shape.text, "p:txBody", size=1800)
    return (
        f"<{element}>{properties}"
        "<p:spPr><a:xfrm>"
        f'<a:off x="{emu(shape.left)}" y="{emu(shape.top)}"/>'
        f'<a:ext cx="{emu(shape.width)}" cy="{emu(shape.height)}"/>'
        f'</a:xfrm><a:prstGeom prst="{shape.geometry or "rect"}"><a:avLst/></a:prstGeom>'
        f"</p:spPr>{body}</{element}>"
    )
