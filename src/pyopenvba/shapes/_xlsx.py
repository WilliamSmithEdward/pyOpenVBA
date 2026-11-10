"""A worksheet's drawing layer.

A sheet's shapes live in ``xl/drawings/drawingN.xml``, each in an
anchor that ties it to cells: ``xdr:twoCellAnchor`` for one that moves
and sizes with them, ``oneCellAnchor`` for one that only moves, and
``absoluteAnchor`` for one pinned to the page.  Inside the anchor is
the shape -- ``xdr:sp`` for an AutoShape or a text box, ``xdr:cxnSp``
for a line, ``xdr:pic`` for a picture, ``xdr:graphicFrame`` for a chart
or a table, ``xdr:grpSp`` for a group -- and the macro a click runs is
an attribute on that element.

Excel's own form controls are the exception.  Their drawing sits in an
``mc:AlternateContent`` whose ``xdr:sp`` is hidden and whose transform
is zero; where the control actually is comes from the anchor, the
macro is in the sheet's own ``<controls>`` element, and what the
control is wired to is in the ``xl/ctrlProps`` part that names.

Every layout here was read from a workbook Excel saved, in
``tests/fixtures/shapes``.
"""

from __future__ import annotations

import math
import re
import uuid
from dataclasses import dataclass, field, replace
from html.parser import HTMLParser

from pyopenvba._xml import escape, unescape
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.shapes import _drawingml as dml
from pyopenvba.shapes._adjustments import ADJUSTMENTS, DARK_TEXT
from pyopenvba.shapes._read_look import read_picture, with_fill_line
from pyopenvba.shapes._read_text import loaded_text, read_text
from pyopenvba.shapes._values import (PRESET_GEOMETRY, ControlInfo, FillGradient, FillLine, Shape, ShapeEffects, TextLine,
                                      TextLook, child_space, emu, points, to_child)

#: The anchors a sheet's drawing is made of.
ANCHORS = ("xdr:twoCellAnchor", "xdr:oneCellAnchor", "xdr:absoluteAnchor", "mc:AlternateContent")

#: What each shape element is, as the object model names it.
_ELEMENTS: dict[str, str] = {
    "xdr:sp": "shape",
    "xdr:cxnSp": "line",
    "xdr:pic": "picture",
    "xdr:graphicFrame": "chart",
    "xdr:grpSp": "group",
}

_MACRO = re.compile(r'\bmacro="([^"]*)"')
#: The shadow an old object's Shadow = True gives a drawing (tests/fixtures/old_drawings.json, loaded_drawing_edits.json).
OLD_SHADOW = ('<a:effectLst><a:outerShdw blurRad="63500" dist="37357" dir="2700000" rotWithShape="0"><a:scrgbClr r="0" g="0" b="0"/>'
              '</a:outerShdw></a:effectLst>')
_TXBOX = re.compile(r"<xdr:cNvSpPr\b[^>]*?\btxBox=\"1\"")
_CONTROL = re.compile(r"<control\b[^>]*?>.*?</control>|<control\b[^>]*?/>", re.DOTALL)
_CONTROL_HEAD = re.compile(r"<control\b([^>]*?)/?>")
_CONTROL_PR = re.compile(r"<controlPr\b([^>]*?)/?>")
_FORM_CONTROL_PR = re.compile(r"<formControlPr\b([^>]*?)/?>")
_ATTRIBUTE = re.compile(r'([\w:.-]+)="([^"]*)"')
_FMLA_MACRO = re.compile(r"<x:FmlaMacro>.*?</x:FmlaMacro>", re.DOTALL)
_COL = re.compile(r"<xdr:col>(\d+)</xdr:col>")
_COLOFF = re.compile(r"<xdr:colOff>(-?\d+)</xdr:colOff>")
_ROW = re.compile(r"<xdr:row>(\d+)</xdr:row>")
_ROWOFF = re.compile(r"<xdr:rowOff>(-?\d+)</xdr:rowOff>")

#: What Excel's default column comes to in points: 8.43 characters is
#: 64 pixels at 96 dpi, and a point is three quarters of a pixel.
DEFAULT_COLUMN_POINTS = 48.0

#: The default row height Excel writes for Calibri 11.
DEFAULT_ROW_POINTS = 14.5

#: An empty drawing part, written when a sheet's first shape is added.
EMPTY_DRAWING = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"'
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></xdr:wsDr>'
)


@dataclass(slots=True)
class SheetGrid:
    """How far across and down a cell sits, in points.

    A shape that carries its own transform does not need this, but a
    form control does: Excel writes its transform as zero and leaves
    the anchor to say where it is.
    """

    #: Column widths in points, by zero-based column index.
    column_widths: dict[int, float] = field(default_factory=lambda: {})
    #: Row heights in points, by zero-based row index.
    row_heights: dict[int, float] = field(default_factory=lambda: {})
    default_column: float = DEFAULT_COLUMN_POINTS
    default_row: float = DEFAULT_ROW_POINTS

    def x(self, column: int, offset: int) -> float:
        before = sum(self.column_widths.get(one, self.default_column) for one in range(column))
        return before + points(offset)

    def y(self, row: int, offset: int) -> float:
        before = sum(self.row_heights.get(one, self.default_row) for one in range(row))
        return before + points(offset)

    def column_at(self, across: float) -> tuple[int, int]:
        """Which column a point falls in, and how far into it, in EMU.

        Excel will not read an offset larger than the cell it is in: it
        clamps one to the column's width and puts the shape somewhere
        else.  An anchor this library writes is worked out here so that
        never happens.
        """
        column, seen = 0, 0.0
        while True:
            width = self.column_widths.get(column, self.default_column)
            if seen + width > across:
                return column, emu(max(across - seen, 0.0))
            seen += width
            column += 1
            if column > 16383:
                return column, 0

    def row_at(self, down: float) -> tuple[int, int]:
        """Which row a point falls in, and how far into it, in EMU."""
        row, seen = 0, 0.0
        while True:
            height = self.row_heights.get(row, self.default_row)
            if seen + height > down:
                return row, emu(max(down - seen, 0.0))
            seen += height
            row += 1
            if row > 1048575:
                return row, 0


def grid_of(sheet_xml: str) -> SheetGrid:
    """The grid a sheet's own XML describes."""
    grid = SheetGrid()
    head = re.search(r"<sheetFormatPr\b([^>]*?)/?>", sheet_xml)
    if head is not None:
        fields = {key: unescape(value) for key, value in _ATTRIBUTE.findall(head.group(1))}
        grid.default_row = _number(fields.get("defaultRowHeight"), DEFAULT_ROW_POINTS)
        width = fields.get("defaultColWidth")
        if width:
            grid.default_column = characters_to_points(_number(width, 8.43))
    for found in re.finditer(r"<col\b([^>]*?)/?>", sheet_xml):
        fields = dict(_ATTRIBUTE.findall(found.group(1)))
        if "width" not in fields:
            continue
        first = int(_number(fields.get("min"), 1)) - 1
        last = int(_number(fields.get("max"), 1)) - 1
        for one in range(first, min(last, first + 16384) + 1):
            grid.column_widths[one] = characters_to_points(_number(fields["width"], 8.43))
    for found in re.finditer(r"<row\b([^>]*?)/?>", sheet_xml):
        fields = dict(_ATTRIBUTE.findall(found.group(1)))
        if "ht" in fields:
            grid.row_heights[int(_number(fields.get("r"), 1)) - 1] = _number(fields["ht"], 14.5)
    return grid


def characters_to_points(width: float) -> float:
    """A column width as Excel stores it, in points.

    Stored in characters of the standard font; Excel turns that into
    whole pixels, seven to the character plus five for the padding, and
    a point is three quarters of a pixel.
    """
    return round(width * 7 + 5) * 0.75


def _number(text: str | None, fallback: float) -> float:
    try:
        return float(text) if text is not None else fallback
    except ValueError:
        return fallback


def read_drawing(xml: str, grid: SheetGrid | None = None) -> list[Shape]:
    """Every shape a sheet's drawing part holds, in the file's order."""
    body = dml.element_span(xml, "xdr:wsDr")
    inner = xml[body[0] : body[1]] if body else xml
    out: list[Shape] = []
    for tag, start, end in dml.children_spans(inner, ANCHORS):
        markup = inner[start:end]
        shape = _from_alternate(markup, grid) if tag == "mc:AlternateContent" else _from_anchor(markup, grid)
        if shape is not None:
            out.append(shape)
    return out


def _from_anchor(markup: str, grid: SheetGrid | None) -> Shape | None:
    """The shape inside one anchor.

    The outermost of them: a group holds shapes of its own, and it is
    the group the sheet lists.
    """
    best: tuple[int, str] | None = None
    for element in _ELEMENTS:
        span = dml.element_span(markup, element)
        if span is not None and (best is None or span[0] < best[0]):
            best = (span[0], element)
    if best is None:
        return None
    span = dml.element_span(markup, best[1])
    assert span is not None
    return _shape_of(markup[span[0] : span[1]], best[1], _ELEMENTS[best[1]], markup, grid)


def _from_alternate(markup: str, grid: SheetGrid | None) -> Shape | None:
    """A form control's drawing, which Excel wraps for older readers."""
    choice = dml.element_span(markup, "mc:Choice")
    inner = markup[choice[0] : choice[1]] if choice else markup
    shape = _from_anchor(inner, grid)
    if shape is not None:
        shape.kind = "formControl"
        shape.control = ControlInfo()
        shape.source = markup
        shape.hidden = False  # a control's drawing is hidden for older readers, which Excel goes past
    return shape


def _shape_of(
    markup: str, element: str, kind: str, anchor: str, grid: SheetGrid | None
) -> Shape:
    shape_id, name = dml.identity(markup)
    left, top, width, height = _anchor_box(anchor, grid) or dml.transform(markup)
    if kind == "shape" and _TXBOX.search(markup):
        kind = "textBox"
    if kind == "chart" and 'http://schemas.openxmlformats.org/drawingml/2006/diagram' in markup:
        kind = 'smartArt'
    macro = _MACRO.search(markup)
    geometry = dml.geometry(markup)
    shape = Shape(
        name=name,
        kind=kind,
        geometry=geometry,
        left=left,
        top=top,
        width=width,
        height=height,
        text=dml.text_of(markup),
        macro=unescape(macro.group(1)) if macro else "",
        shape_id=shape_id,
        source=anchor,
        auto_shape_type=next((number for number, preset in PRESET_GEOMETRY.items() if preset == geometry), 1),
    )
    shape.loaded_box = (left, top, width, height)
    if kind == "picture":
        read_picture(markup, shape)
    if kind == "group":
        shape.children = _members(markup, grid)
        _placed_in_group(shape.children, _CHILD_SPACE.search(markup), (left, top, width, height))
    _read_settings(shape, markup, anchor)
    return shape


_CNVPR_OPEN = re.compile(r"<xdr:cNvPr\b([^>]*?)/?>")
_ROTATION = re.compile(r"<a:xfrm\b[^>]*?\brot=\"(-?[0-9]{1,10})\"")


def _read_settings(shape: Shape, markup: str, anchor: str) -> None:
    """AlternativeText, Title, Decorative, Rotation, LockAspectRatio and Locked, as the file holds them."""
    identity = _CNVPR_OPEN.search(markup)
    if identity is not None:
        for field_name, attribute in (("description", "descr"), ("title", "title")):
            found = re.search(rf'\b{attribute}="([^"]*)"', identity.group(1))
            if found is not None:
                setattr(shape, field_name, unescape(found.group(1)))
    shape.decorative = re.search(r'<adec:decorative\b[^>]*\bval="1"', markup) is not None
    head = re.match(r"<xdr:sp\b([^>]*)>", markup)
    shape.locks_text = head is None or re.search(r'\bfLocksText="(?:0|false)"', head.group(1)) is None
    # A form control's drawing is hidden for its own reasons; Visible reads only a shape's own hidden.
    link = _LINK.search(markup[: markup.find("</xdr:cNvPr>") if "</xdr:cNvPr>" in markup else 0])
    if link is not None:
        shape.link = (_attribute(link.group(1), "r:id"), _attribute(link.group(1), "tooltip"), _attribute(link.group(1), "location"))
    shape.hidden = (shape.kind != "formControl" and identity is not None
                    and re.search(r'\bhidden="(?:1|true)"', identity.group(1)) is not None)
    turned = _ROTATION.search(markup)
    shape.rotation = int(turned.group(1)) / 60000 if turned else 0.0
    transform = re.search(r"<a:xfrm\b([^>]*)>", markup)
    shape.flip_h = transform is not None and re.search(r'\bflipH="(?:1|true)"', transform.group(1)) is not None
    shape.flip_v = transform is not None and re.search(r'\bflipV="(?:1|true)"', transform.group(1)) is not None
    locks = re.search(r"<a:(?:sp|cxnSp|pic)Locks\b([^>]*)/?>", markup)
    if locks is not None:
        shape.lock_aspect = re.search(r'\bnoChangeAspect="(?:1|true)"', locks.group(1)) is not None
    client = re.search(r"<xdr:clientData\b([^>]*)/?>", anchor)
    shape.locked = client is None or re.search(r'\bfLocksWithSheet="(?:0|false)"', client.group(1)) is None
    shape.prints = client is None or re.search(r'\bfPrintsWithSheet="(?:0|false)"', client.group(1)) is None
    placement = re.search(r'<xdr:twoCellAnchor\b[^>]*\beditAs="([^"]*)"', anchor)
    shape.edit_as = placement.group(1) if placement and placement.group(1) != "twoCell" else ""
    # The other two kinds of anchor say it by what they are (tests/fixtures/drawing_anchors.json).
    if anchor.lstrip().startswith("<xdr:oneCellAnchor"):
        shape.edit_as = "oneCell"
    elif anchor.lstrip().startswith("<xdr:absoluteAnchor"):
        shape.edit_as = "absolute"


_ANCHOR_SIZE = re.compile(r'<xdr:ext\b[^>]*?\bcx="(-?\d+)"[^>]*?\bcy="(-?\d+)"')
_ANCHOR_PLACE = re.compile(r'<xdr:pos\b[^>]*?\bx="(-?\d+)"[^>]*?\by="(-?\d+)"')
_CHILD_SPACE = re.compile(r'<a:chOff\b[^>]*?\bx="(-?\d+)"[^>]*?\by="(-?\d+)"[^>]*/>\s*<a:chExt\b[^>]*?\bcx="(-?\d+)"[^>]*?\bcy="(-?\d+)"')
_LINK = re.compile(r'<a:hlinkClick\b([^>]*)>')


def _attribute(attributes: str, name: str) -> str:
    found = re.search(rf'(?:^|\s){re.escape(name)}="([^"]*)"', attributes)
    return unescape(found.group(1)) if found else ""


def _placed_in_group(members: list[Shape], space: re.Match[str] | None, box: tuple[float, float, float, float]) -> None:
    """A group's members, written in its child space, placed where the group's anchor has put that space."""
    if space is None or not int(space.group(3)) or not int(space.group(4)):
        return
    origin_x, origin_y = int(space.group(1)) / 12700, int(space.group(2)) / 12700
    across, down = box[2] / (int(space.group(3)) / 12700), box[3] / (int(space.group(4)) / 12700)
    if (origin_x, origin_y, across, down) == (box[0], box[1], 1.0, 1.0):
        return
    for member in members:
        member.left, member.top = box[0] + (member.left - origin_x) * across, box[1] + (member.top - origin_y) * down
        member.width, member.height = member.width * across, member.height * down
        _rescaled(member.children, origin_x, origin_y, across, down, box)


def _rescaled(members: list[Shape], origin_x: float, origin_y: float, across: float, down: float, box: tuple[float, float, float, float]) -> None:
    for member in members:
        member.left, member.top = box[0] + (member.left - origin_x) * across, box[1] + (member.top - origin_y) * down
        member.width, member.height = member.width * across, member.height * down
        _rescaled(member.children, origin_x, origin_y, across, down, box)


def _anchor_box(anchor: str, grid: SheetGrid | None) -> tuple[float, float, float, float] | None:
    """Where the anchor says the shape is, in points: what Excel goes by, whatever the shape's own transform says
    (tests/fixtures/drawing_anchors.json). Cells for both corners, or one cell or a place on the sheet and a size."""
    size = _ANCHOR_SIZE.search(anchor)
    width, height = (int(size.group(1)) / 12700, int(size.group(2)) / 12700) if size else (0.0, 0.0)
    place = _ANCHOR_PLACE.search(anchor)
    if place is not None:
        return int(place.group(1)) / 12700, int(place.group(2)) / 12700, width, height
    if grid is None:
        return None
    start = dml.element_span(anchor, "xdr:from")
    finish = dml.element_span(anchor, "xdr:to")
    if start is None:
        return None
    left, top = _corner(anchor[start[0] : start[1]], grid)
    if finish is None:
        return left, top, width, height
    right, bottom = _corner(anchor[finish[0] : finish[1]], grid)
    turned = _ROTATION.search(anchor)
    if turned is not None and 45 <= int(turned.group(1)) / 60000 % 180 < 135:
        # From 45 to 135 degrees, and 225 to 315, the cells hold the box turned about its centre.
        width, height = bottom - top, right - left
        return (left + right - width) / 2, (top + bottom - height) / 2, width, height
    return left, top, right - left, bottom - top


def _corner(markup: str, grid: SheetGrid) -> tuple[float, float]:
    column = _COL.search(markup)
    column_offset = _COLOFF.search(markup)
    row = _ROW.search(markup)
    row_offset = _ROWOFF.search(markup)
    return (
        grid.x(int(column.group(1)) if column else 0, int(column_offset.group(1)) if column_offset else 0),
        grid.y(int(row.group(1)) if row else 0, int(row_offset.group(1)) if row_offset else 0),
    )


def _members(markup: str, grid: SheetGrid | None) -> list[Shape]:
    """A group's own shapes, which sit directly inside its grpSp."""
    inner = markup
    span = dml.element_span(markup, "xdr:grpSp")
    if span is not None:
        head = markup.find(">", span[0]) + 1
        inner = markup[head : span[1] - len("</xdr:grpSp>")]
    out: list[Shape] = []
    for tag, start, end in dml.children_spans(inner, tuple(_ELEMENTS)):
        piece = inner[start:end]
        out.append(_shape_of(piece, tag, _ELEMENTS[tag], piece, grid))
    return out


def read_controls(sheet_xml: str, parts: dict[str, str]) -> dict[int, ControlInfo]:
    """What each form control is and does, by its shape id.

    The sheet names the control and the macro it runs; the part it
    points at holds what kind of control it is and what it is wired to.
    """
    out: dict[int, ControlInfo] = {}
    for found in _CONTROL.finditer(sheet_xml):
        markup = found.group(0)
        head = _CONTROL_HEAD.search(markup)
        if head is None:
            continue
        fields = dict(_ATTRIBUTE.findall(head.group(1)))
        try:
            shape_id = int(fields.get("shapeId", "0"))
        except ValueError:
            continue
        relationship = fields.get("r:id", "")
        control = ControlInfo(relationship=relationship)
        properties = _CONTROL_PR.search(markup)
        if properties is not None:
            control.macro_text = unescape(dict(_ATTRIBUTE.findall(properties.group(1))).get("macro", ""))
        part = parts.get(relationship, "")
        if part:
            settings = _FORM_CONTROL_PR.search(part)
            if settings is not None:
                values = {key: unescape(value) for key, value in _ATTRIBUTE.findall(settings.group(1))}
                control.kind = values.get("objectType", "control")
                control.first_button = values.get("firstButton", "0") == "1"
                control.lock_text = values.get("lockText", "0") == "1"
                control.three_d = values.get("noThreeD", "0") != "1"
                control.linked_cell = values.get("fmlaLink", "")
                control.list_range = values.get("fmlaRange", "")
                control.value = int(_number(values.get("val"), 0))
                control.selection_mode = values.get("seltype", values.get("selType", "single")).lower()
                control.selected_indices = [int(one) for one in re.findall(r"\d+", values.get("multiSel", ""))]
                control.items = [unescape(dict(_ATTRIBUTE.findall(one)).get("val", ""))
                                 for one in re.findall(r"<item\b([^>]*)/>", part)]
                if control.kind in {"Drop", "List"}:
                    control.value = int(_number(values.get("sel"), 0))
                if control.kind in {"CheckBox", "Radio"}:
                    control.value = {"Checked": 1, "Mixed": 2}.get(values.get("checked", ""), -4146)
                if control.kind in {"Spin", "Scroll"}:
                    control.minimum = int(_number(values.get("min"), 0))
                    control.maximum = int(_number(values.get("max"), 0))
                    control.increment = int(_number(values.get("inc"), 1))
                    control.page_change = int(_number(values.get("page"), 10))
                    control.drop_width = int(_number(values.get("dx"), 31))
        out[shape_id] = control
    return out


def control_macro(sheet_xml: str, shape_id: int) -> str:
    """The macro a form control runs, as the sheet holds it.

    Excel writes ``[0]!Name`` for a procedure in the workbook the
    control is in, and reports it as ``Book.xlsm!Name``.
    """
    for found in _CONTROL.finditer(sheet_xml):
        markup = found.group(0)
        head = _CONTROL_HEAD.search(markup)
        if head is None or f'shapeId="{shape_id}"' not in head.group(1):
            continue
        properties = _CONTROL_PR.search(markup)
        if properties is None:
            return ""
        return unescape(dict(_ATTRIBUTE.findall(properties.group(1))).get("macro", ""))
    return ""


def with_control_macro(sheet_xml: str, shape_id: int, macro: str) -> str:
    """The sheet with one control's macro set or cleared."""
    for found in _CONTROL.finditer(sheet_xml):
        markup = found.group(0)
        head = _CONTROL_HEAD.search(markup)
        if head is None or f'shapeId="{shape_id}"' not in head.group(1):
            continue
        properties = _CONTROL_PR.search(markup)
        if properties is None:
            return sheet_xml
        attributes = properties.group(1)
        if 'macro="' in attributes:
            changed = re.sub(r'\smacro="[^"]*"', lambda _: f' macro="{escape(macro)}"' if macro else "", attributes)
        elif macro:
            changed = attributes + f' macro="{escape(macro)}"'
        else:
            changed = attributes
        start = found.start() + properties.start(1)
        return sheet_xml[:start] + changed + sheet_xml[found.start() + properties.end(1) :]
    return sheet_xml


#: The shape type every form control is drawn from, which comes before the first control a VML part draws.
CONTROL_SHAPE_TYPE = (
    '<v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201"\r\n'
    '  path="m,l,21600r21600,l21600,xe">\r\n'
    '  <v:stroke joinstyle="miter"/>\r\n'
    '  <v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/>\r\n'
    '  <o:lock v:ext="edit" shapetype="t"/>\r\n'
    " </v:shapetype>"
)

#: The VML a sheet needs before it can hold a control at all: the id
#: map and the shape type every button is drawn from.
EMPTY_VML = (
    '<xml xmlns:v="urn:schemas-microsoft-com:vml"\r\n'
    ' xmlns:o="urn:schemas-microsoft-com:office:office"\r\n'
    ' xmlns:x="urn:schemas-microsoft-com:office:excel">\r\n'
    ' <o:shapelayout v:ext="edit">\r\n'
    '  <o:idmap v:ext="edit" data="1"/>\r\n'
    f" </o:shapelayout>{CONTROL_SHAPE_TYPE}</xml>\r\n"
)

#: What each control this can make is called in its own part, and what
#: ClientData says it is.
CONTROL_KINDS: dict[str, str] = {
    "Radio": "Radio",
    "Spin": "Spin",
    "Scroll": "Scroll",
    "GBox": "GBox",
    "Drop": "Drop",
    "List": "List",
    "Button": "Button",
    "CheckBox": "Checkbox",
    "Check Box": "Checkbox",
    "DropDown": "Drop",
    "Drop Down": "Drop",
    "ListBox": "List",
    "OptionButton": "Radio",
    "GroupBox": "GBox",
    "Label": "Label",
    "ScrollBar": "Scroll",
    "Spinner": "Spin",
}


def control_drawing(shape: Shape, grid: SheetGrid | None = None) -> str:
    """The drawing half of a form control, as Excel writes it.

    Wrapped in an ``mc:AlternateContent`` so a reader without the 2010
    drawing extensions still finds something, with the shape hidden and
    its transform empty: where the control is comes from the anchor.
    """
    grid = grid or SheetGrid()
    edit_as = f' editAs="{shape.edit_as}"' if shape.edit_as else ""
    return (
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        '<mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">'
        f"<xdr:twoCellAnchor{edit_as}>"
        f"{_corner_markup('from', grid, on_pixel(shape.left), on_pixel(shape.top))}"
        f"{_corner_markup('to', grid, on_pixel(shape.left + shape.width), on_pixel(shape.top + shape.height))}"
        '<xdr:sp macro="" textlink="">'
        "<xdr:nvSpPr>"
        f'<xdr:cNvPr id="{shape.shape_id}" name="{escape(shape.name)}" hidden="1">'
        "<a:extLst>"
        '<a:ext uri="{63B3BB69-23CF-44E3-9099-C40C66FF867C}">'
        f'<a14:compatExt spid="_x0000_s{shape.shape_id}"/>'
        "</a:ext>"
        "</a:extLst>"
        "</xdr:cNvPr>"
        "<xdr:cNvSpPr/>"
        "</xdr:nvSpPr>"
        '<xdr:spPr bwMode="auto">'
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:noFill/><a:ln w="9525"><a:miter lim="800000"/><a:headEnd/><a:tailEnd/></a:ln>'
        "</xdr:spPr>"
        "</xdr:sp>"
        '<xdr:clientData fPrintsWithSheet="0"/>'
        "</xdr:twoCellAnchor>"
        "</mc:Choice>"
        "<mc:Fallback/>"
        "</mc:AlternateContent>"
    )


def control_entry(shape: Shape, relationship: str, grid: SheetGrid | None = None) -> str:
    """The sheet's own record of a control: what it is and what it runs."""
    grid = grid or SheetGrid()
    macro = f' macro="{escape(shape.macro and f"[0]!{shape.macro}")}"' if shape.macro else ""
    return (
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        '<mc:Choice Requires="x14">'
        f'<control shapeId="{shape.shape_id}" r:id="{relationship}" name="{escape(shape.name)}">'
        f'<controlPr{"" if shape.locked else " locked=" + chr(34) + "0" + chr(34)} defaultSize="0" autoFill="0" autoPict="0"{macro}{_alt_text(shape)}>'
        f"<anchor{_anchor_flags(shape)}>"
        f"{_corner_markup('from', grid, on_pixel(shape.left), on_pixel(shape.top)).replace('xdr:from', 'from')}"
        f"{_corner_markup('to', grid, on_pixel(shape.left + shape.width), on_pixel(shape.top + shape.height)).replace('xdr:to', 'to')}"
        "</anchor>"
        "</controlPr>"
        "</control>"
        "</mc:Choice>"
        "</mc:AlternateContent>"
    )


def on_pixel(points: float) -> float:
    """A control's edge in the file: the nearest pixel, a half going up, at the 96 DPI the model emulates (tests/fixtures/new_control_places.json)."""
    return math.floor(points * 4 / 3 + 0.5) * 0.75


def _alt_text(shape: Shape) -> str:
    return f' altText="{escape(shape.description)}"' if shape.description else ""


def _anchor_flags(shape: Shape) -> str:
    """How a control's anchor in the sheet says what its Placement does."""
    return ("" if shape.edit_as == "absolute" else ' moveWithCells="1"') + (' sizeWithCells="1"' if shape.edit_as in ("", "twoCell") else "")


def control_properties(shape: Shape) -> str:
    """A control's own part: which control it is and what it is wired to."""
    control = shape.control
    kind = CONTROL_KINDS.get(control.kind if control else "Button", "Button")
    if kind == "Checkbox":
        kind = "CheckBox"
    linked = f' fmlaLink="{escape(control.linked_cell)}"' if control and control.linked_cell else ""
    listed = f' fmlaRange="{escape(control.list_range)}"' if control and control.list_range else ""
    # A drop-down names its lines only once a macro sets them (tests/fixtures/control_format_more.json).
    lines = f' dropLines="{control.drop_lines}"' if control and control.drop_lines is not None else ""
    extra = f' dropStyle="combo"{lines}' if kind == "Drop" else ""
    if kind in {"CheckBox", "Radio"} and control:
        checked = {1: "Checked", 2: "Mixed"}.get(control.value, "")
        if checked:
            extra += f' checked="{checked}"'
    if kind in {"Drop", "List"} and control:
        extra += f' sel="{control.value}" val="0"'
    # A list or drop-down has no text to lock; the rest lock theirs until LockedText is turned off.
    locked = "" if kind in {"Drop", "List"} or control and not control.lock_text else ' lockText="1"'
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"'
        f' objectType="{kind}"{linked}{listed}{extra}{locked}/>'
    )
    xml = with_control_bindings(xml, control) if control else xml
    if kind in {"CheckBox", "Radio", "Drop", "List", "GBox"} and not (control and control.three_d):
        # Shading off, as Excel writes every new one of these (tests/fixtures/new_control_settings.json), so a save read back is as new.
        xml = _ordered_attribute(xml, "formControlPr", "noThreeD", "1", _FORM_CONTROL_ORDER)
    return xml


def control_anchor(shape: Shape, grid: SheetGrid | None = None) -> str:
    """The eight numbers a control's VML places it by.

    Column, offset, row, offset for each corner, and the offsets are in
    half-points: the fixture's button sits 12 points into its column
    and says 24, and 6 points into the next and says 12.  Taken from a
    file Excel wrote, where all four numbers come out exact.
    """
    grid = grid or SheetGrid()
    left_column, left_offset = grid.column_at(shape.left)
    top_row, top_offset = grid.row_at(shape.top)
    right_column, right_offset = grid.column_at(shape.left + shape.width)
    bottom_row, bottom_offset = grid.row_at(shape.top + shape.height)
    numbers = (
        left_column,
        _half_points(left_offset),
        top_row,
        _half_points(top_offset),
        right_column,
        _half_points(right_offset),
        bottom_row,
        _half_points(bottom_offset),
    )
    return ", ".join(str(one) for one in numbers)


def _half_points(offset: int) -> int:
    """An EMU offset as the half-points a VML anchor counts in."""
    return int(round(points(offset) * 2))


def control_vml(shape: Shape, grid: SheetGrid | None = None) -> str:
    """The VML Excel actually draws a control from."""
    control = shape.control
    checked = ""
    kind = CONTROL_KINDS.get(control.kind if control else "Button", "Button")
    if control and kind in {"Checkbox", "Radio"}:
        state = {1: 1, 2: 2}.get(control.value, 0)
        checked = f"<x:Checked>{state}</x:Checked>"
        if kind == "Radio" and control.first_button:
            checked += "<x:FirstButton/>"
    if control and control.kind in {"Drop", "List"}:
        mode = {"single": "Single", "multi": "Multi", "extended": "Extend"}[control.selection_mode]
        checked = f"<x:Val>0</x:Val><x:Sel>{control.value}</x:Sel><x:SelType>{mode}</x:SelType>"
        if control.selection_mode != "single":
            checked += f"<x:MultiSel>{','.join(str(one) for one in control.selected_indices)}</x:MultiSel>"
        checked += "".join(f"<x:ListItem>{escape(one)}</x:ListItem>" for one in control.items)
    if control and kind in {"Spin", "Scroll"}:
        checked += "".join(f"<x:{tag}>{value}</x:{tag}>" for tag, value in (
            ("Val", control.value), ("Min", control.minimum),
            ("Max", control.maximum if control.maximum is not None else (30000 if kind == "Spin" else 100)),
            ("Inc", control.increment), ("Page", control.page_change), ("Dx", control.drop_width),
        ))
    macro = f"<x:FmlaMacro>[0]!{escape(shape.macro)}</x:FmlaMacro>" if shape.macro else ""
    linked = (
        f"<x:FmlaLink>{escape(control.linked_cell)}</x:FmlaLink>"
        if control and control.linked_cell
        else ""
    )
    listed = (
        f"<x:FmlaRange>{escape(control.list_range)}</x:FmlaRange>"
        if control and control.list_range
        else ""
    )
    text = (
        f"<v:textbox style='mso-direction-alt:auto' o:singleclick='f'>"
        f"<div style='text-align:center'>{'<br/>'.join(escape(line) for line in shape.text.split(chr(10)))}</div></v:textbox>"
        if shape.text
        else ""
    )
    alternative = f' alt="{escape(shape.description)}"' if shape.description else ""
    locks = ' aspectratio="t"' if shape.lock_aspect else ""
    return _new_control_settings(_vml_state((
        f'<v:shape id="{_vml_id(shape.name)}" o:spid="_x0000_s{shape.shape_id}"'
        f' type="#_x0000_t201"{alternative}'
        f" style='position:absolute;margin-left:{shape.left:g}pt;margin-top:{shape.top:g}pt;"
        f"width:{shape.width:g}pt;height:{shape.height:g}pt;z-index:1;mso-wrap-style:tight'"
        ' o:button="t" fillcolor="buttonFace [67]" o:insetmode="auto">'
        '<v:fill color2="buttonFace [67]" o:detectmouseclick="t"/>'
        f'<o:lock v:ext="edit" rotation="t"{locks}/>'
        f"{text}"
        f'<x:ClientData ObjectType="{kind}">'
        f"<x:Anchor>{control_anchor(shape, grid)}</x:Anchor>"
        f"{_vml_settings(kind, control)}"
        "<x:AutoFill>False</x:AutoFill>"
        f"{macro}{linked}{listed}{checked}"
        "<x:TextHAlign>Center</x:TextHAlign>"
        "<x:TextVAlign>Center</x:TextVAlign>"
        "</x:ClientData>"
        "</v:shape>"
    ), shape), shape)


def _new_control_settings(markup: str, shape: Shape) -> str:
    """A new control's VML with what Excel writes there for it (tests/fixtures/new_control_edits.json): a button's flags set this
    session, Display3DShading off as NoThreeD or NoThreeD2, as on every new check box, option button, list and drop-down, and a button's
    text frame setting."""
    control = shape.control
    if control is None:
        return markup
    kind = CONTROL_KINDS.get(control.kind, control.kind)
    for flag in sorted(control.button_flags):
        markup = _ordered_child(markup, flag, f"<x:{flag}/>")
    if kind in ("Drop", "List"):
        markup = _ordered_child(markup, "NoThreeD2", "" if control.three_d else "<x:NoThreeD2/>")
    elif kind in ("Checkbox", "Radio"):
        markup = _ordered_child(markup, "NoThreeD", "" if control.three_d else "<x:NoThreeD/>")
    return with_vml_text_frame(markup, shape)


def _vml_settings(kind: str, control: ControlInfo | None) -> str:
    """A control's PrintObject, Disabled, LockText and DropLines in its VML (tests/fixtures/control_format_more.json):
    a button prints only once told to, the rest unless told not to; a drop-down shows 8 lines until set."""
    printed = control.print_object if control and control.print_object is not None else kind != "Button"
    settings = "" if printed else "<x:PrintObject>False</x:PrintObject>"
    if control and control.disabled:
        settings += "<x:Disabled/>"
    if control and not control.lock_text:
        settings += "<x:LockText>False</x:LockText>"
    if kind == "Drop":
        lines = control.drop_lines if control and control.drop_lines is not None else 8
        settings += f"<x:DropLines>{lines}</x:DropLines>"
    return settings


def _vml_id(name: str) -> str:
    """A shape's name as a VML id, which cannot hold a space."""
    cleaned = re.sub(r"[^A-Za-z0-9_.-]", "", name)
    return cleaned if cleaned and not cleaned[0].isdigit() else f"_{cleaned}"


def with_vml_shape(vml: str, markup: str) -> str:
    """The VML with one more shape in it."""
    tail = vml.rfind("</xml>")
    if tail < 0:
        return vml + markup
    return vml[:tail] + markup + vml[tail:]


def with_vml_macro(vml: str, shape_id: int, macro: str) -> str:
    """The VML with one control's macro set or cleared.

    A form control keeps its macro in three places at once: the sheet's
    controlPr, the VML shape's ``x:FmlaMacro``, and nothing in the
    drawing at all.  Excel reads the VML, so a change that misses it is
    a change that does not happen.
    """
    spid = f"_x0000_s{shape_id}"
    span = _vml_shape_span(vml, spid)
    if span is None:
        return vml
    markup = vml[span[0] : span[1]]
    if _FMLA_MACRO.search(markup):
        changed = (
            _FMLA_MACRO.sub(lambda _: f"<x:FmlaMacro>{escape(macro)}</x:FmlaMacro>", markup, count=1)
            if macro
            else _FMLA_MACRO.sub("", markup, count=1)
        )
    elif macro:
        changed = markup.replace(
            "<x:ClientData",
            "<x:ClientData",
            1,
        )
        head = changed.find(">", changed.find("<x:ClientData"))
        if head < 0:
            return vml
        changed = (
            changed[: head + 1]
            + f"<x:FmlaMacro>{escape(macro)}</x:FmlaMacro>"
            + changed[head + 1 :]
        )
    else:
        changed = markup
    return vml[: span[0]] + changed + vml[span[1] :]


def _vml_shape_span(vml: str, spid: str) -> tuple[int, int] | None:
    """Where one VML shape is, found by the id Excel gave it."""
    at = 0
    while True:
        span = dml.element_span(vml, "v:shape", at)
        if span is None:
            return None
        at = span[1]
        # Excel's own file names a control by its id, and the model's by its spid.
        if f'"{spid}"' in vml[span[0] : vml.index(">", span[0])]:
            return span


class _Caption(HTMLParser):
    """Read the rendered text in a VML textbox, including line breaks."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.pieces: list[str] = []

    def handle_data(self, data: str) -> None:
        self.pieces.append(re.sub(r"\s+", " ", data))

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag == "br":
            self.pieces.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag == "div":
            self.pieces.append("\n")


def control_text(vml: str, shape_id: int) -> str:
    span = _vml_shape_span(vml, f"_x0000_s{shape_id}")
    if span is None:
        return ""
    markup = vml[span[0]:span[1]]
    box = dml.element_span(markup, "v:textbox")
    if box is None:
        return ""
    reader = _Caption()
    reader.feed(markup[box[0]:box[1]])
    reader.close()
    return "\n".join(line.strip() for line in "".join(reader.pieces).splitlines()).strip()


def read_control_settings(vml: str, shape: Shape, control: ControlInfo) -> None:
    """A control's Visible, Locked, Enabled, PrintObject and DropDownLines, which only its VML holds (tests/fixtures/old_collection_members.json)."""
    span = _vml_shape_span(vml, f"_x0000_s{shape.shape_id}")
    if span is None:
        return
    markup = vml[span[0] : span[1]]
    shape.hidden = "visibility:hidden" in markup
    alternative = re.search(r'<v:shape\b[^>]*?\balt="([^"]*)"', markup, re.DOTALL)
    shape.description = unescape(alternative.group(1)) if alternative else ""
    shape.lock_aspect = 'aspectratio="t"' in markup
    shape.locked = re.search(r"<x:Locked>\s*False", markup) is None
    control.disabled = "<x:Disabled" in markup
    control.print_object = re.search(r"<x:PrintObject>\s*False", markup) is None
    lines = re.search(r"<x:DropLines>\s*(\d+)", markup)
    control.drop_lines = int(lines.group(1)) if lines else None


def without_control(sheet_xml: str, shape_id: int) -> str:
    """Remove a sheet control and its otherwise empty compatibility wrappers."""
    def remove(match: re.Match[str]) -> str:
        head = _CONTROL_HEAD.search(match.group(0))
        fields = dict(_ATTRIBUTE.findall(head.group(1))) if head else {}
        return "" if fields.get("shapeId") == str(shape_id) else match.group(0)

    out = _CONTROL.sub(remove, sheet_xml)
    # Controls can be inside nested Choice blocks. Prune only empty
    # wrappers, keeping sibling controls and their markup unchanged.
    for _ in range(3):
        out = re.sub(r"<controls\b[^>]*>\s*</controls>", "", out)
        out = re.sub(
            r"<mc:AlternateContent\b[^>]*>\s*<mc:Choice\b[^>]*>\s*</mc:Choice>"
            r"\s*(?:<mc:Fallback\s*/>\s*)?</mc:AlternateContent>", "", out,
        )
    return out


def without_vml_control(vml: str, shape_id: int) -> str:
    span = _vml_shape_span(vml, f"_x0000_s{shape_id}")
    return vml[:span[0]] + vml[span[1]:] if span else vml


def with_control_shape(sheet_xml: str, shape: Shape, grid: SheetGrid) -> str:
    """Patch one control's name and anchor without replacing its other settings."""
    def change(match: re.Match[str]) -> str:
        markup = match.group(0)
        head = _CONTROL_HEAD.search(markup)
        if head is None or dict(_ATTRIBUTE.findall(head.group(1))).get("shapeId") != str(shape.shape_id):
            return markup
        markup = _attribute_on_tag(markup, "control", "name", shape.name)
        for tag, x, y in (
            ("from", shape.left, shape.top), ("to", shape.left + shape.width, shape.top + shape.height),
        ):
            span = dml.element_span(markup, tag)
            if span:
                corner = _corner_markup(tag, grid, x, y).replace(f"xdr:{tag}", tag)
                markup = markup[:span[0]] + corner + markup[span[1]:]
        return markup

    return _CONTROL.sub(change, sheet_xml)


def with_control_state(sheet_xml: str, shape: Shape) -> str:
    """Patch what the sheet says of a control's Locked and Placement (tests/fixtures/old_collection_edits.json)."""
    def change(found: re.Match[str]) -> str:
        markup = found.group(0)
        head = _CONTROL_HEAD.search(markup)
        if head is None or dict(_ATTRIBUTE.findall(head.group(1))).get("shapeId") != str(shape.shape_id):
            return markup
        markup = _ordered_attribute(markup, "controlPr", "locked", None if shape.locked else "0", _CONTROL_PR_ORDER)
        markup = _ordered_attribute(markup, "controlPr", "altText", shape.description or None, _CONTROL_PR_ORDER)
        markup = _ordered_attribute(markup, "anchor", "moveWithCells", None if shape.edit_as == "absolute" else "1", _ANCHOR_ORDER)
        return _ordered_attribute(markup, "anchor", "sizeWithCells", "1" if shape.edit_as in ("", "twoCell") else None, _ANCHOR_ORDER)

    return _CONTROL.sub(change, sheet_xml)


#: The schema's order of what this module patches on a control loaded from a file: x:ClientData's children, controlPr's attributes and
#: formControlPr's. Excel puts a new one before the first already there that comes after it, whatever order the rest keep
#: (tests/fixtures/loaded_control_edits.json).
_CLIENT_ORDER = ('MoveWithCells', 'SizeWithCells', 'Anchor', 'Locked', 'DefaultSize', 'PrintObject', 'Disabled', 'AutoFill', 'AutoLine', 'AutoPict',
                 'FmlaMacro', 'TextHAlign', 'TextVAlign', 'LockText', 'JustLastX', 'SecretEdit', 'Default', 'Help', 'Cancel', 'Dismiss', 'Accel',
                 'Accel2', 'Row', 'Column', 'Visible', 'RowHidden', 'ColHidden', 'VTEdit', 'MultiLine', 'VScroll', 'ValidIds', 'FmlaRange', 'WidthMin',
                 'Sel', 'NoThreeD2', 'SelType', 'MultiSel', 'LCT', 'ListItem', 'DropStyle', 'Colored', 'DropLines', 'Checked', 'FmlaLink', 'FmlaPict',
                 'NoThreeD', 'FirstButton', 'FmlaGroup', 'Val', 'Min', 'Max', 'Inc', 'Page', 'Horiz', 'Dx')
_CONTROL_PR_ORDER = ('locked', 'defaultSize', 'print', 'disabled', 'recalcAlways', 'uiObject', 'autoFill', 'autoLine', 'autoPict', 'macro', 'altText',
                     'linkedCell', 'listFillRange', 'cf', 'r:id')
_ANCHOR_ORDER = ('moveWithCells', 'sizeWithCells', 'z-order')
_FORM_CONTROL_ORDER = ('objectType', 'checked', 'colored', 'dropLines', 'dropStyle', 'dx', 'firstButton', 'fmlaGroup', 'fmlaLink', 'fmlaRange',
                       'fmlaTxbx', 'horiz', 'inc', 'justLastX', 'lockText', 'max', 'min', 'multiSel', 'noThreeD', 'noThreeD2', 'page', 'sel', 'seltype',
                       'textHAlign', 'textVAlign', 'val', 'widthMin', 'editVal', 'multiLine', 'verticalBar', 'passwordEdit')


def _ordered_attribute(markup: str, tag: str, name: str, value: str | None, order: tuple[str, ...]) -> str:
    """One attribute of ``tag`` set, or taken off for None, a new one placed where ``order`` puts it among those already there."""
    found = re.search(rf"<{re.escape(tag)}(?=[\s/>])[^>]*>", markup)
    if found is None:
        return markup
    head = found.group(0)
    attribute = re.compile(rf'\s{re.escape(name)}="[^"]*"')
    if value is None:
        head = attribute.sub("", head, count=1)
    elif attribute.search(head):
        head = attribute.sub(lambda _: f' {name}="{escape(value)}"', head, count=1)
    else:
        later = order[order.index(name) + 1:]
        place = next((match.start() for match in re.finditer(r'\s([\w:]+)="', head) if match.group(1) in later), None)
        if place is None:
            place = len(head) - (2 if head.endswith("/>") else 1)
        head = head[:place] + f' {name}="{escape(value)}"' + head[place:]
    return markup[:found.start()] + head + markup[found.end():]


def _ordered_child(markup: str, tag: str, wanted: str) -> str:
    """One of x:ClientData's children as ``wanted`` spells it, or taken off where that is empty, a new one placed in _CLIENT_ORDER."""
    # Each child on a line of its own as Excel writes them, "\r\n   " before each; none in the model's own VML.
    spacing = re.search(r"<x:ClientData\b[^>]*>(\s*)<", markup)
    separator = spacing.group(1) if spacing else ""
    found = re.search(rf"\s*<x:{tag}\b[^>]*?(?:/>|>.*?</x:{tag}>)", markup, re.DOTALL)
    if found is not None:
        if not wanted:
            return markup[:found.start()] + markup[found.end():]
        element = found.group(0).lstrip()
        start = found.end() - len(element)
        return markup if element == wanted else markup[:start] + wanted + markup[found.end():]
    if not wanted:
        return markup
    later = _CLIENT_ORDER[_CLIENT_ORDER.index(tag) + 1:]
    place = next((match.start() for match in re.finditer(r"<x:(\w+)", markup) if match.group(1) in later), None)
    if place is not None:
        return markup[:place] + wanted + separator + markup[place:]
    end = markup.find("</x:ClientData>")
    if end < 0:
        return markup
    # After the last child: where the space before the closing tag starts.
    last = len(markup[:end].rstrip())
    return markup[:last] + separator + wanted + markup[last:]


def with_control_flags(sheet_xml: str, shape: Shape) -> str:
    """What the sheet's controlPr says of a loaded control's PrintObject and Enabled."""
    control = shape.control

    def change(found: re.Match[str]) -> str:
        markup = found.group(0)
        head = _CONTROL_HEAD.search(markup)
        if control is None or head is None or dict(_ATTRIBUTE.findall(head.group(1))).get("shapeId") != str(shape.shape_id):
            return markup
        markup = _ordered_attribute(markup, "controlPr", "print", "0" if control.print_object is False else None, _CONTROL_PR_ORDER)
        return _ordered_attribute(markup, "controlPr", "disabled", "1" if control.disabled else None, _CONTROL_PR_ORDER)

    return _CONTROL.sub(change, sheet_xml)


def with_vml_flags(vml: str, shape: Shape) -> str:
    """A loaded control's PrintObject, Disabled, LockText, DropLines and NoThreeD in its VML."""
    control = shape.control
    span = _vml_shape_span(vml, f"_x0000_s{shape.shape_id}")
    if control is None or span is None:
        return vml
    markup = vml[span[0]:span[1]]
    kind = CONTROL_KINDS.get(control.kind, control.kind)
    markup = _ordered_child(markup, "PrintObject", "<x:PrintObject>False</x:PrintObject>" if control.print_object is False else "")
    markup = _ordered_child(markup, "Disabled", "<x:Disabled/>" if control.disabled else "")
    if kind in ("Checkbox", "Button", "Radio"):
        markup = _ordered_child(markup, "LockText", "" if control.lock_text else "<x:LockText>False</x:LockText>")
    if kind == "Drop":
        markup = _ordered_child(markup, "DropLines", f"<x:DropLines>{8 if control.drop_lines is None else control.drop_lines}</x:DropLines>")
    # A button's flags set this session go in; one already there stays, as Excel reads none of them back to change.
    for flag in sorted(control.button_flags):
        markup = _ordered_child(markup, flag, f"<x:{flag}/>")
    # A list and a drop-down keep Display3DShading turned off as NoThreeD2, a check box and an option button as NoThreeD.
    if kind in ("Drop", "List"):
        markup = _ordered_child(markup, "NoThreeD2", "" if control.three_d else "<x:NoThreeD2/>")
    elif kind in ("Checkbox", "Radio"):
        markup = _ordered_child(markup, "NoThreeD", "" if control.three_d else "<x:NoThreeD/>")
    return vml[:span[0]] + markup + vml[span[1]:]


def with_control_settings(properties: str, control: ControlInfo) -> str:
    """A loaded control's part with its LockText, DropLines and Display3DShading, where it keeps them: lockText is written only while on,
    dropLines only once set from its 8, noThreeD only while the shading is off."""
    kind = CONTROL_KINDS.get(control.kind, control.kind)
    if kind in ("Checkbox", "Button", "Radio"):
        properties = _ordered_attribute(properties, "formControlPr", "lockText", "1" if control.lock_text else None, _FORM_CONTROL_ORDER)
    if kind == "Drop":
        lines = None if control.drop_lines in (None, 8) else str(control.drop_lines)
        properties = _ordered_attribute(properties, "formControlPr", "dropLines", lines, _FORM_CONTROL_ORDER)
    if kind in ("Drop", "List", "Checkbox", "Radio"):
        properties = _ordered_attribute(properties, "formControlPr", "noThreeD", None if control.three_d else "1", _FORM_CONTROL_ORDER)
    return properties


def _vml_flag(markup: str, tag: str, wanted: str) -> str:
    """One setting among a VML client's, as ``wanted`` spells it, none where that is empty."""
    found = re.search(rf"\s*<x:{tag}\b[^>]*?(?:/>|>.*?</x:{tag}>)", markup, re.DOTALL)
    if found is not None and wanted:
        return markup
    if found is not None:
        return markup[:found.start()] + markup[found.end():]
    return markup.replace("</x:ClientData>", f"{wanted}</x:ClientData>") if wanted else markup


#: A button's text frame, what DrawingML writes for each value (tests/fixtures/button_text.json): HorizontalAlignment as pPr's algn and the
#: inset it drops, VerticalAlignment as bodyPr's anchor and the inset it drops (Justify and Distributed in the VML alone), Orientation
#: as bodyPr's vert and the insets it drops.
_ALIGN = {-4131: ("l", "rIns"), -4108: ("ctr", ""), -4152: ("r", "lIns"), -4130: ("just", ""), -4117: ("dist", "")}
_ANCHOR = {-4160: ("t", "bIns"), -4108: ("ctr", ""), -4107: ("b", "tIns"), -4130: (None, ""), -4117: (None, "")}
_VERT = {-4128: (None, ()), -4171: ("vert270", ()), -4170: ("vert", ()), -4166: ("wordArtVert", ("tIns", "bIns"))}
#: What the VML gets: TextHAlign and TextVAlign as named (none for the first, which is the default), the div's text-align (none where
#: it goes), and the textbox's layout flow.
_VML_H = {-4131: (None, "left"), -4108: ("Center", "center"), -4152: ("Right", "right"), -4130: ("Justify", "justify"), -4117: ("Distributed", None)}
_VML_V = {-4160: None, -4108: "Center", -4107: "Bottom", -4130: "Justify", -4117: "Distributed"}
_VML_FLOW = {-4128: "", -4171: "layout-flow:vertical;mso-layout-flow-alt:bottom-to-top;", -4170: "layout-flow:vertical;",
             -4166: "layout-flow:vertical;mso-layout-flow-alt:top-to-bottom;"}
TEXT_FRAME_VALUES = {"HorizontalAlignment": tuple(_ALIGN), "VerticalAlignment": tuple(_ANCHOR), "Orientation": tuple(_VERT)}
_BODY_ORDER = ("rot", "spcFirstLastPara", "vertOverflow", "horzOverflow", "vert", "wrap", "lIns", "tIns", "rIns", "bIns", "numCol", "spcCol",
               "rtlCol", "fromWordArt", "anchor", "anchorCtr", "forceAA", "upright", "compatLnSpc")


def control_text_frame(source: str) -> dict[str, int]:
    """A loaded control's HorizontalAlignment, VerticalAlignment and Orientation as its DrawingML holds them, what Excel reads back
    (tests/fixtures/button_text.json); a control without one reads them centred and horizontal."""
    body = re.search(r"<a:bodyPr\b[^>]*>", source)
    paragraph = re.search(r'<a:pPr\b[^>]*\balgn="(\w+)"', source)
    anchor = re.search(r'\banchor="(\w+)"', body.group(0)) if body else None
    vert = re.search(r'\bvert="(\w+)"', body.group(0)) if body else None
    return {
        "HorizontalAlignment": next((key for key, (name, _) in _ALIGN.items() if paragraph and name == paragraph.group(1)), -4108),
        "VerticalAlignment": next((key for key, (name, _) in _ANCHOR.items() if anchor and name == anchor.group(1)), -4108),
        "Orientation": next((key for key, (name, _) in _VERT.items() if vert and name == vert.group(1)), -4128),
    }


def text_frame_measured(source: str) -> bool:
    """Whether a loaded control's text is laid out as the button measured was: one paragraph centred, anchored in the middle, horizontal,
    with every inset Excel's 27432."""
    body = re.search(r"<a:bodyPr\b[^>]*>", source)
    if body is None or len(re.findall(r"<a:pPr\b", source)) != 1 or ' vert="' in body.group(0) or 'anchor="ctr"' not in body.group(0):
        return False
    insets = dict(re.findall(r'\b([lrtb]Ins)="(\d+)"', body.group(0)))
    return insets == {name: "27432" for name in ("lIns", "tIns", "rIns", "bIns")} and re.search(r'<a:pPr\b[^>]*\balgn="ctr"', source) is not None


def with_text_frame(markup: str, control: ControlInfo) -> str:
    """A loaded button's DrawingML with the text frame setting it was given this session, from the layout measured."""
    if control.text_frame is None:
        return markup
    name, value = control.text_frame
    body = re.search(r"<a:bodyPr\b[^>]*>", markup)
    if body is None:
        return markup
    head = body.group(0)
    dropped: tuple[str, ...] = ()
    if name == "HorizontalAlignment":
        algn, side = _ALIGN[value]
        markup = re.sub(r'(<a:pPr\b[^>]*\balgn=")\w+"', rf'\g<1>{algn}"', markup, count=1)
        dropped = (side,) if side else ()
    elif name == "VerticalAlignment":
        anchor, side = _ANCHOR[value]
        if anchor is not None:
            head = _ordered_attribute(head, "a:bodyPr", "anchor", anchor, _BODY_ORDER)
        dropped = (side,) if side else ()
    else:
        vert, dropped = _VERT[value]
        if vert is not None:
            head = _ordered_attribute(head, "a:bodyPr", "vert", vert, _BODY_ORDER)
    for side in dropped:
        head = _ordered_attribute(head, "a:bodyPr", side, "0", _BODY_ORDER)
    body = re.search(r"<a:bodyPr\b[^>]*>", markup)
    assert body is not None
    return markup[:body.start()] + head + markup[body.end():]


def with_vml_text_frame(vml: str, shape: Shape) -> str:
    """A loaded button's VML with the text frame setting it was given this session: TextHAlign or TextVAlign, the div's text-align and the
    textbox's layout flow, each line laid out again as Excel lays it out (tests/fixtures/button_text.json)."""
    control = shape.control
    span = _vml_shape_span(vml, f"_x0000_s{shape.shape_id}")
    if control is None or control.text_frame is None or span is None:
        return vml
    name, value = control.text_frame
    markup = vml[span[0]:span[1]]
    align: str | None = "keep"
    if name == "HorizontalAlignment":
        element, align = _VML_H[value]
        markup = _ordered_child(markup, "TextHAlign", f"<x:TextHAlign>{element}</x:TextHAlign>" if element else "")
    elif name == "VerticalAlignment":
        element = _VML_V[value]
        markup = _ordered_child(markup, "TextVAlign", f"<x:TextVAlign>{element}</x:TextVAlign>" if element else "")
        # Justified or distributed down, the div loses its text-align (measured; Excel reads the alignment across unchanged).
        align = None if value in (-4130, -4117) else "keep"
    else:
        box = re.search(r"<v:textbox\b[^>]*>", markup)
        if box is not None and _VML_FLOW[value] and "layout-flow" not in box.group(0):
            tag = re.sub(r"(\bstyle=')", rf"\g<1>{_VML_FLOW[value]}", box.group(0), count=1)
            markup = markup[:box.start()] + _vml_rewrapped(tag, "   ", 2) + markup[box.end():]
    if align != "keep":
        division = re.search(r"<div\b.*?</div>", markup, re.DOTALL)
        if division is not None:
            text = division.group(0)
            text = (re.sub(r"(<div\b[^>]*\bstyle=')text-align:\w+", rf"\g<1>text-align:{align}", text, count=1) if align else
                    re.sub(r"<div\s+style='text-align:\w+'>", "<div>", text, count=1))
            markup = markup[:division.start()] + _vml_rewrapped(text, "   ", 3) + markup[division.end():]
    return vml[:span[0]] + markup + vml[span[1]:]


def vml_wrapped(pieces: list[str], indent: str = "  ", column: int = 0) -> str:
    """A line of VML as Excel's writer lays it out: it breaks, ``indent`` in, before a piece that starts with a space, an attribute or
    a word of text, once the line has reached column 67, and before a style declaration once it has reached 70, the space dropped;
    ``column`` is how far in the line already is where the first piece starts. Fitted to every v:shape and v:textbox tag and div line in
    tests/fixtures/loaded_control_edits.json, button_text.json, notes.json and the workbooks they open, 1,293 lines, no other pair of
    limits fitting them all; the words alone allow any limit from 40 to 83."""
    lines = [pieces[0]]
    for piece in pieces[1:]:
        if column * (len(lines) == 1) + len(lines[-1]) >= (67 if piece.startswith(" ") else 70):
            lines.append(indent + piece.lstrip(" "))
        else:
            lines[-1] += piece
    return "\r\n".join(lines)


_VML_ATTRIBUTE = re.compile(r"""\s[\w:]+=(?:"[^"]*"|'[^']*')""")


def _vml_rewrapped(line: str, indent: str = "  ", column: int = 0) -> str:
    """A line of VML Excel wrote, after an edit, laid out again as Excel lays one out: it may break before each attribute of its tags
    and after each declaration of a style."""
    flat = re.sub(rf";\r\n{indent}", ";", re.sub(rf"(?<!;)\r\n{indent}", " ", line))
    # Text breaks between words, each word taking the space before it as an attribute does.
    cuts = {0, *(text.start(1) + space.start() for text in re.finditer(r">([^<]+)<", flat) for space in re.finditer(r" ", text.group(1)))}
    for tag in re.finditer(r"<[^>]*>", flat):
        for attribute in _VML_ATTRIBUTE.finditer(tag.group(0)):
            start = tag.start() + attribute.start()
            cuts.add(start)
            if attribute.group(0).startswith(" style="):
                cuts.update(start + semicolon.end() for semicolon in re.finditer(";", attribute.group(0)[:-2]))
    ordered = sorted(cuts)
    return vml_wrapped([flat[one:two] for one, two in zip(ordered, [*ordered[1:], len(flat)])], indent, column)


def _vml_state(markup: str, shape: Shape, *, loaded: bool = False) -> str:
    """A control's Visible, Locked and Placement in its VML, changed only where they differ: on a loaded control each goes where Excel
    puts it among those its file holds (_CLIENT_ORDER), on the model's own at the end."""
    if shape.hidden != ("visibility:hidden" in markup):
        if not shape.hidden:
            markup = re.sub(r";?visibility:hidden", "", markup, count=1)
        elif "mso-wrap-style" in markup:
            markup = markup.replace("mso-wrap-style", "visibility:hidden;mso-wrap-style", 1)
        else:
            markup = re.sub(r"(\bstyle=['\"][^'\"]*)", r"\1;visibility:hidden", markup, count=1)
    markup = _attribute_on_tag(markup, "v:shape", "alt", shape.description or None)
    flag = _ordered_child if loaded else _vml_flag
    markup = flag(markup, "Locked", "" if shape.locked else "<x:Locked>False</x:Locked>")
    markup = flag(markup, "MoveWithCells", "<x:MoveWithCells/>" if shape.edit_as == "absolute" else "")
    return flag(markup, "SizeWithCells", "" if shape.edit_as in ("", "twoCell") else "<x:SizeWithCells/>")


def _attribute_on_tag(markup: str, tag: str, name: str, value: str | None) -> str:
    found = re.search(rf"<{re.escape(tag)}(?=[\s/>])[^>]*>", markup)
    if found is None:
        return markup
    head = found.group(0)
    attribute = re.compile(rf"\b{re.escape(name)}=([\"'])(.*?)\1", re.DOTALL)
    replacement = f'{name}="{escape(value)}"' if value is not None else ""
    if value is None:
        head = attribute.sub("", head, count=1)
    elif attribute.search(head):
        head = attribute.sub(lambda _: replacement, head, count=1)
    else:
        at = len(head) - (2 if head.endswith("/>") else 1)
        head = head[:at] + " " + replacement + head[at:]
    return markup[:found.start()] + head + markup[found.end():]


def with_vml_control(vml: str, shape: Shape, before: Shape, grid: SheetGrid) -> str:
    """Patch a control's visible shape, preserving its style and other properties."""
    span = _vml_shape_span(vml, f"_x0000_s{shape.shape_id}")
    if span is None:
        return vml
    markup = vml[span[0]:span[1]]
    if shape.name != before.name:
        # What it was found by goes on, as Excel's does (tests/fixtures/old_collection_edits.json).
        markup = _attribute_on_tag(markup, "v:shape", "o:spid", f"_x0000_s{shape.shape_id}")
        markup = _attribute_on_tag(markup, "v:shape", "id", _vml_id(shape.name))
    markup = _vml_state(markup, shape, loaded=True)
    if (shape.left, shape.top, shape.width, shape.height) != (before.left, before.top, before.width, before.height):
        style = re.search(r"\bstyle=([\"'])(.*?)\1", markup, re.DOTALL)
        if style:
            body = style.group(2)
            for key, value in (("margin-left", shape.left), ("margin-top", shape.top),
                               ("width", shape.width), ("height", shape.height)):
                pattern = re.compile(rf"(?<![\w-]){key}\s*:[^;]+")
                entry = f"{key}:{value:g}pt"
                body = pattern.sub(lambda _: entry, body) if pattern.search(body) else body + ";" + entry
            markup = markup[:style.start(2)] + body + markup[style.end(2):]
        anchor = dml.element_span(markup, "x:Anchor")
        if anchor:
            markup = (markup[:anchor[0]] + f"<x:Anchor>{control_anchor(shape, grid)}</x:Anchor>"
                      + markup[anchor[1]:])
    if shape.text != before.text:
        body = "<div>" + "<br/>".join(escape(line) for line in shape.text.split("\n")) + "</div>"
        box = dml.element_span(markup, "v:textbox")
        words = [] if box is None else list(re.finditer(r">([^<]*[^<\s][^<]*)<", markup[box[0]:box[1]]))
        if box and len(words) == 1 and "\n" not in shape.text:
            # One line in one run of words: Excel changes the words alone, in the div and font it has
            # (tests/fixtures/loaded_control_edits.json).
            start, end = box[0] + words[0].start(1), box[0] + words[0].end(1)
            markup = markup[:start] + escape(shape.text) + markup[end:]
        elif box:
            start = markup.find(">", box[0]) + 1
            markup = markup[:start] + body + markup[box[1] - len("</v:textbox>"):]
        else:
            start = markup.find("<x:ClientData")
            if start >= 0:
                markup = markup[:start] + f"<v:textbox>{body}</v:textbox>" + markup[start:]
    head, was = re.match(r"<v:shape\b[^>]*>", markup), re.match(r"<v:shape\b[^>]*>", vml[span[0]:span[1]])
    if head is not None and was is not None and head.group(0) != was.group(0) and "\r\n" in was.group(0):
        # An edited opening tag Excel wrote is laid out again as Excel's writer lays one out.
        leading = vml[:span[0]]
        markup = _vml_rewrapped(head.group(0), column=len(leading) - leading.rfind("\n") - 1) + markup[head.end():]
    return vml[:span[0]] + markup + vml[span[1]:]


def with_control_bindings(properties: str, control: ControlInfo) -> str:
    """Patch modeled bindings and values, preserving other control settings."""
    for name, value in (("fmlaLink", control.linked_cell), ("fmlaRange", control.list_range)):
        properties = _attribute_on_tag(properties, "formControlPr", name, value)
    kind = CONTROL_KINDS.get(control.kind, control.kind)
    if kind == "Radio":
        properties = _attribute_on_tag(properties, "formControlPr", "firstButton", "1" if control.first_button else None)
    if kind in {"Checkbox", "Radio"}:
        state = {1: "Checked", 2: "Mixed"}.get(control.value, "")
        properties = _attribute_on_tag(properties, "formControlPr", "checked", state or None)
    if kind in {"Spin", "Scroll"}:
        for attribute, value in (
            ("val", control.value), ("min", control.minimum),
            ("max", control.maximum if control.maximum is not None else (30000 if kind == "Spin" else 100)),
            ("inc", control.increment), ("page", control.page_change), ("dx", control.drop_width),
        ):
            properties = _attribute_on_tag(properties, "formControlPr", attribute, str(value))
    if control.kind in {"Drop", "List"}:
        properties = _attribute_on_tag(properties, "formControlPr", "sel", str(control.value))
        properties = _attribute_on_tag(properties, "formControlPr", "seltype", control.selection_mode)
        properties = _attribute_on_tag(properties, "formControlPr", "multiSel",
                                       ",".join(str(one) for one in control.selected_indices))
        def attribute_text(value: str) -> str:
            return escape(value).replace("\r", "&#13;").replace("\n", "&#10;").replace("\t", "&#9;")

        item_list = "<itemLst>" + "".join(f'<item val="{attribute_text(one)}"/>' for one in control.items) + "</itemLst>"
        if not control.items:
            item_list = ""
        span = dml.element_span(properties, "itemLst")
        if span:
            properties = properties[:span[0]] + item_list + properties[span[1]:]
        elif item_list:
            head = _FORM_CONTROL_PR.search(properties)
            if head and head.group(0).endswith("/>"):
                properties = (properties[:head.end() - 2] + ">" + item_list
                              + "</formControlPr>" + properties[head.end():])
            else:
                properties = properties.replace("</formControlPr>", item_list + "</formControlPr>")
    return properties


def with_vml_bindings(vml: str, shape_id: int, control: ControlInfo) -> str:
    """Keep the legacy representation of the same bindings in sync."""
    span = _vml_shape_span(vml, f"_x0000_s{shape_id}")
    if span is None:
        return vml
    markup = vml[span[0]:span[1]]
    settings = [("FmlaLink", control.linked_cell), ("FmlaRange", control.list_range)]
    kind = CONTROL_KINDS.get(control.kind, control.kind)
    if kind == "Radio":
        markup = re.sub(r"<x:FirstButton\b[^>]*(?:/>|>.*?</x:FirstButton>)", "", markup, flags=re.DOTALL)
        if control.first_button:
            markup = markup.replace("</x:ClientData>", "<x:FirstButton/></x:ClientData>")
    if kind in {"Checkbox", "Radio"}:
        settings.append(("Checked", str({1: 1, 2: 2}.get(control.value, 0))))
    if kind in {"Spin", "Scroll"}:
        settings.extend((tag, str(value)) for tag, value in (
            ("Val", control.value), ("Min", control.minimum),
            ("Max", control.maximum if control.maximum is not None else (30000 if kind == "Spin" else 100)),
            ("Inc", control.increment), ("Page", control.page_change), ("Dx", control.drop_width),
        ))
    if control.kind in {"Drop", "List"}:
        settings.append(("Sel", str(control.value)))
        settings.append(("SelType", {"single": "Single", "multi": "Multi", "extended": "Extend"}[control.selection_mode]))
        settings.append(("MultiSel", ",".join(str(one) for one in control.selected_indices)))
        markup = re.sub(r"<x:ListItem>.*?</x:ListItem>", "", markup, flags=re.DOTALL)
        items = "".join(f"<x:ListItem>{escape(one)}</x:ListItem>" for one in control.items)
        markup = markup.replace("</x:ClientData>", items + "</x:ClientData>")
    for tag, value in settings:
        pattern = re.compile(rf"<x:{tag}(?:\s[^>]*)?/>|<x:{tag}>.*?</x:{tag}>", re.DOTALL)
        replacement = f"<x:{tag}>{escape(value)}</x:{tag}>" if value else ""
        if pattern.search(markup):
            markup = pattern.sub(lambda _: replacement, markup)
        elif replacement:
            markup = markup.replace("</x:ClientData>", replacement + "</x:ClientData>")
    return vml[:span[0]] + markup + vml[span[1]:]


def new_anchor(shape: Shape, grid: SheetGrid | None = None) -> str:
    """The markup for a shape this library is adding.

    A twoCellAnchor holding the absolute transform as well as the
    cells.  The cells are what Excel reads: it clamps an offset to the
    cell that holds it, so the corner each one lands in is worked out
    from the sheet's own columns and rows.
    """
    grid = grid or SheetGrid()
    left, top, width, height = _turned_box(shape)
    opening = f'<xdr:twoCellAnchor editAs="{shape.edit_as}">' if shape.edit_as else "<xdr:twoCellAnchor>"
    return (
        f"{opening}"
        f"{_corner_markup('from', grid, left, top)}"
        f"{_corner_markup('to', grid, left + width, top + height)}"
        f"{_new_element(shape)}{_client_data(shape)}</xdr:twoCellAnchor>"
    )


def shape_default_markup(shape: Shape) -> str:
    """The theme's spDef for SetShapesDefaultProperties on a new AutoShape (tests/fixtures/shape_columns.json)."""
    look = _fill_line(shape)
    return (f"<a:spDef><a:spPr>{look}</a:spPr>" if look else "<a:spDef><a:spPr/>") + (
        '<a:bodyPr vertOverflow="clip" horzOverflow="clip" rtlCol="0" anchor="t"/><a:lstStyle><a:defPPr algn="l">'
        '<a:defRPr sz="1100"/></a:defPPr></a:lstStyle><a:style><a:lnRef idx="2"><a:schemeClr val="accent1">'
        '<a:shade val="15000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>'
        '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/>'
        "</a:fontRef></a:style></a:spDef>")


_PICTURE_FILLS = {"clear": '<a:solidFill><a:schemeClr val="accent1"><a:alpha val="0"/></a:schemeClr></a:solidFill>',
                  "solid": '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>'}


def _picture_effects(shape: Shape) -> str:
    """What a picture's blip holds of PictureFormat's settings, in Excel's order (tests/fixtures/loaded_shape_sets.json)."""
    effects = ""
    if shape.picture_clear is not None:
        effects += (f'<a:clrChange><a:clrFrom><a:srgbClr val="{shape.picture_clear}"/></a:clrFrom><a:clrTo>'
                    f'<a:srgbClr val="{shape.picture_clear}"><a:alpha val="0"/></a:srgbClr></a:clrTo></a:clrChange>')
    effects += {2: "<a:grayscl/>", 3: '<a:biLevel thresh="50000"/>'}.get(shape.picture_colour_type, "")
    if shape.picture_lum is not None:
        bright, contrast = shape.picture_lum
        effects += "<a:lum" + (f' bright="{bright}"' if bright else "") + (f' contrast="{contrast}"' if contrast else "") + "/>"
    return effects


def _with_loaded_picture(markup: str, shape: Shape, was: Shape) -> str:
    """A loaded picture's markup with what PictureFormat changed. Effects the model does not read stop it being changed."""
    if (shape.picture_lum, shape.picture_colour_type, shape.picture_clear) != (was.picture_lum, was.picture_colour_type, was.picture_clear):
        found = re.search(r"(<a:blip\b[^>]*?)(/>|>(.*?)</a:blip>)", markup, re.DOTALL)
        if found is not None:
            inside = found.group(3) or ""
            at = inside.find("<a:extLst")
            rest = "" if at < 0 else inside[at:]
            kept = _picture_effects(was)
            if (inside if at < 0 else inside[:at]) != kept:
                raise VBAUnsupportedError("Saving PictureFormat over effects the model does not read is not implemented")
            effects = _picture_effects(shape) + rest
            blip = f"{found.group(1)}>{effects}</a:blip>" if effects else f"{found.group(1)}/>"
            markup = markup[:found.start()] + blip + markup[found.end():]
    if shape.picture_fill is not None:
        markup = with_fill_line(markup, _PICTURE_FILLS[shape.picture_fill])
    return markup


def _with_loaded_links(markup: str, shape: Shape) -> str:
    """A loaded connector's markup with the connections it holds now, and any flips a connection gave it."""
    span = re.search(r"<xdr:cNvCxnSpPr(?:/>|>(.*?)</xdr:cNvCxnSpPr>)", markup, re.DOTALL)
    if span is None:
        return markup
    kept = re.sub(r"<a:(?:stCxn|endCxn)\b[^>]*/>", "", span.group(1) or "")
    kept += "" if shape.begin_link is None else f'<a:stCxn id="{shape.begin_link.target}" idx="{shape.begin_link.site - 1}"/>'
    kept += "" if shape.end_link is None else f'<a:endCxn id="{shape.end_link.target}" idx="{shape.end_link.site - 1}"/>'
    written = f"<xdr:cNvCxnSpPr>{kept}</xdr:cNvCxnSpPr>" if kept else "<xdr:cNvCxnSpPr/>"
    return markup[:span.start()] + written + markup[span.end():]


def _new_text_body(shape: Shape) -> str:
    """A shape's txBody as the model writes it."""
    body = dml.text_body(shape.text, "xdr:txBody")
    if shape.kind in ("shape", "freeform"):
        # Excel starts an AutoShape's paragraph left-aligned and keeps it so once text is set
        # (tests/fixtures/text_frame.json).
        body = body.replace("<a:p>", '<a:p><a:pPr algn="l"/>', 1)
    if shape.kind == "textBox":
        # A text box's text runs across, as Excel writes it.
        body = body.replace('horzOverflow="clip" rtlCol=', 'horzOverflow="clip" vert="horz" rtlCol=', 1)
    if shape.kind == "shape" and shape.geometry in DARK_TEXT:
        body = body.replace('<a:endParaRPr lang="en-US" sz="1100"/>', '<a:endParaRPr lang="en-US" sz="1100"><a:solidFill>'
                            '<a:schemeClr val="tx1"/></a:solidFill></a:endParaRPr>')
    if shape.text_look is not None:
        body = _laid_out(body, shape.text_look, _marked_paragraphs(shape, shape.text_look))
    if shape.painted:
        body = body.replace(
            '<a:pPr algn="l"/><a:endParaRPr lang="en-US" sz="1100"/>',
            '<a:pPr marL="0" indent="0" algn="l"/><a:endParaRPr lang="en-US" sz="1100"><a:solidFill><a:schemeClr val="lt1"/>'
            '</a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:endParaRPr>')
    if shape.text_paragraphs is not None:
        own = iter(shape.text_paragraphs)
        body = re.sub(r"<a:p>(?:<a:pPr\b[^>]*/>)?", lambda _: f"<a:p>{next(own, '')}", body)
    return body


def read_shape_text(shape: Shape) -> None:
    """Fill in a loaded shape's text settings from its markup, the first time a macro reaches them."""
    if shape.text_read or not shape.source:
        return
    found = read_text(shape.source)
    if shape.text == dml.text_of(shape.source):
        # What VBA reaches tells a line break from a paragraph's end, where the shape's plain text does not.
        shape.text = loaded_text(shape.source)
    shape.text_look, shape.text_marks, shape.text_end, shape.text_paragraphs = found.look, found.marks, found.end, found.paragraphs
    shape.text_read = True
    span = dml.element_span(shape.source, "xdr:txBody")
    shape.text_rewritable = span is not None and _new_text_body(shape) == shape.source[span[0]:span[1]]


def _with_loaded_text(markup: str, shape: Shape, was: Shape) -> str:
    """A loaded shape's markup with what a macro changed of its text: written out if the model writes this text just as the file
    does; otherwise only new words, in the first run's settings, with the frame and the first paragraph's kept."""
    span = dml.element_span(markup, "xdr:txBody")
    if span is not None and shape.kind == "formControl":
        # A control's caption: only the words change, in the first run's settings (tests/fixtures/loaded_control_edits.json).
        return markup if shape.text == was.text else markup[:span[0]] + _new_words(markup[span[0]:span[1]], shape.text) + markup[span[1]:]
    if span is None or shape.kind not in ("shape", "textBox"):
        return dml.replace_text(markup, shape.text, "xdr:txBody") if shape.text != was.text else markup
    if shape.text_read:
        before = read_text(shape.source)
        same = (before.look, before.marks, before.end, before.paragraphs) == (
            shape.text_look, shape.text_marks, shape.text_end, shape.text_paragraphs)
        if same and shape.text == loaded_text(shape.source):
            return markup
        if shape.text_rewritable:
            return markup[:span[0]] + _new_text_body(shape) + markup[span[1]:]
    elif shape.text == was.text:
        return markup
    return markup[:span[0]] + _new_words(markup[span[0]:span[1]], shape.text) + markup[span[1]:]


def _new_words(body: str, text: str) -> str:
    start = body.find("<a:p>") if "<a:p>" in body else body.find("<a:p ")
    head = body if start < 0 else body[:start]
    run = re.search(r"<a:rPr\b[^>]*?(?:/>|>.*?</a:rPr>)", body, re.DOTALL)
    end = re.search(r"<a:endParaRPr\b[^>]*?(?:/>|>.*?</a:endParaRPr>)", body, re.DOTALL)
    settings = re.search(r"<a:pPr\b[^>]*?(?:/>|>.*?</a:pPr>)", body, re.DOTALL)
    paragraph = "" if settings is None else settings.group(0)
    props = run.group(0) if run is not None else "" if end is None else end.group(0).replace("endParaRPr", "rPr")
    lines: list[str] = []
    for line in text.split("\n"):
        inner = f"<a:r>{props}<a:t>{dml.escaped_text(line)}</a:t></a:r>" if line else (
            "" if end is None and not props else (end.group(0) if end is not None else props.replace("rPr", "endParaRPr")))
        lines.append(f"<a:p>{paragraph}{inner}</a:p>")
    return head + "".join(lines) + "</xdr:txBody>"


def _new_picture(shape: Shape) -> str:
    """A picture AddPicture made, its image in the part the relationship names (tests/fixtures/shape_add_picture.json)."""
    locks = '<a:picLocks noChangeAspect="1"/>' if shape.lock_aspect else "<a:picLocks/>"
    effects = _picture_effects(shape)
    macro = f' macro="{escape(shape.macro)}"' if shape.macro else ""
    blip = (f'r:link="{shape.image_relationship}"/>' if shape.image_link else
            f'r:embed="{shape.image_relationship}">{effects}<a:extLst><a:ext uri="{{28A0092B-C50C-407E-A947-70E740481C1C}}">'
            f'<a14:useLocalDpi {_A14} val="0"/></a:ext></a:extLst></a:blip>')
    return (
        f"<xdr:pic{macro}><xdr:nvPicPr>{_identity(shape)}<xdr:cNvPicPr>{locks}</xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill>"
        f'<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" {blip}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
        f'<xdr:spPr><a:xfrm{_xfrm_attributes(shape)}><a:off x="{emu(shape.left)}" y="{emu(shape.top)}"/>'
        f'<a:ext cx="{emu(shape.width)}" cy="{emu(shape.height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f"{_PICTURE_FILLS.get(shape.picture_fill or '', '')}</xdr:spPr></xdr:pic>")


def _new_element(shape: Shape) -> str:
    """A new shape's own element, which an anchor or a group holds."""
    if shape.kind == "group":
        # The members sit in the group's child space, which a move or a resize of the group leaves alone
        # (tests/fixtures/shape_group.json, tests/fixtures/shape_group_items.json).
        across, down, wide, tall = child_space(shape)
        members = ""
        for one in shape.children:
            left, top, width, height = to_child(shape, (one.left, one.top, one.width, one.height))
            members += _new_element(replace(one, left=left, top=top, width=width, height=height))
        return (
            f"<xdr:grpSp><xdr:nvGrpSpPr>{_identity(shape)}<xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>"
            f'<xdr:grpSpPr><a:xfrm><a:off x="{emu(shape.left)}" y="{emu(shape.top)}"/>'
            f'<a:ext cx="{emu(shape.width)}" cy="{emu(shape.height)}"/><a:chOff x="{emu(across)}" y="{emu(down)}"/>'
            f'<a:chExt cx="{emu(wide)}" cy="{emu(tall)}"/></a:xfrm></xdr:grpSpPr>{members}</xdr:grpSp>'
        )
    if shape.kind == "picture":
        return _new_picture(shape)
    line = shape.kind == "line"
    body = "" if line else _new_text_body(shape)
    macro = f' macro="{escape(shape.macro)}"' if shape.macro else ' macro=""'
    element = "xdr:cxnSp" if line else "xdr:sp"
    box = ' txBox="1"' if shape.kind == "textBox" else ""
    textlink = "" if line else ' textlink=""' + ("" if shape.locks_text else ' fLocksText="0"')
    identity = _identity(shape)
    locks = "" if shape.lock_aspect is None else (
        '<a:spLocks noChangeAspect="1"/>' if shape.lock_aspect else "<a:spLocks/>")
    properties = (
        f"<xdr:nvCxnSpPr>{identity}{_connections(shape)}</xdr:nvCxnSpPr>"
        if line
        else f"<xdr:nvSpPr>{identity}<xdr:cNvSpPr{box}>{locks}</xdr:cNvSpPr></xdr:nvSpPr>".replace(
            f"<xdr:cNvSpPr{box}></xdr:cNvSpPr>", f"<xdr:cNvSpPr{box}/>")
    )
    turn = _xfrm_attributes(shape)
    guides = ""
    if shape.adjustments is not None:
        guides = "".join(f'<a:gd name="{name}" fmla="val {value}"/>' for (name, _, _), value in zip(
            ADJUSTMENTS[shape.geometry], shape.adjustments, strict=True))
    geometry = (f'<a:prstGeom prst="{shape.geometry or "rect"}"><a:avLst>{guides}</a:avLst></a:prstGeom>'
                if guides else f'<a:prstGeom prst="{shape.geometry or "rect"}"><a:avLst/></a:prstGeom>')
    style = _TEXT_BOX_STYLE if shape.kind == "textBox" else _STYLE
    if shape.kind == "line" or shape.geometry == "arc":
        # A line has its own style, as an open freeform does (tests/fixtures/shape_fill_line.json), and so does an arc
        # (tests/fixtures/old_drawings.json).
        style = _OPEN_STYLE
    if shape.kind == "freeform":
        geometry = _free_geometry(shape)
        # An open path is drawn as a line, a closed one filled as an AutoShape is (tests/fixtures/shape_polyline.json).
        if not any(command == "close" for command, _ in shape.path):
            style = _OPEN_STYLE
    return (
        f"<{element}{macro}{textlink}>"
        f"{properties}"
        f"<xdr:spPr><a:xfrm{turn}>"
        f'<a:off x="{emu(shape.left)}" y="{emu(shape.top)}"/>'
        f'<a:ext cx="{emu(shape.width)}" cy="{emu(shape.height)}"/>'
        f"</a:xfrm>{geometry}"
        f"{_fill_line(shape)}</xdr:spPr>"
        f"{style}{body}"
        f"</{element}>"
    )


def _fill_line(shape: Shape) -> str:
    """The fill and line a shape's spPr holds: a text box's own, and only what Shape.Fill and Shape.Line set on top of
    the style, as Excel writes them (tests/fixtures/shape_fill_line.json)."""
    look = shape.fill_line
    text_box = shape.kind == "textBox"
    if look is None:
        return _TEXT_BOX_FILL if text_box else ""
    fill_colour = _colour(look.fill_rgb, look.fill_scheme or look.style_fill or ("lt1" if text_box else "accent1"), 0, look.fill_alpha,
                          look.fill_mods, look.fill_index)
    fill_paint = f"<a:solidFill>{fill_colour}</a:solidFill>"
    if look.fill_pattern is not None:
        back = _colour(look.fill_back_rgb or (None if look.fill_back_scheme else "FFFFFF"), look.fill_back_scheme or "",
                       0, look.fill_alpha, look.fill_back_mods)
        fill_paint = (f'<a:pattFill prst="{look.fill_pattern}"><a:fgClr>{fill_colour}</a:fgClr><a:bgClr>{back}</a:bgClr>'
                      "</a:pattFill>")
    if look.fill_gradient is not None:
        fill_paint = _gradient(look.fill_gradient)
    fill = ""
    if look.fill_hidden or (look.fill_bare and shape.kind != "picture"):
        fill = "<a:noFill/>"
    elif look.fill_bare:
        pass  # A picture with no fill of its own stays without one (tests/fixtures/shape_settings_kept.json).
    elif look.fill_pattern is not None or look.fill_gradient is not None:
        fill = fill_paint
    elif (look.fill_rgb is not None or look.fill_alpha is not None or look.fill_solid or text_box
          or look.fill_scheme is not None or look.fill_mods or look.fill_index is not None):
        fill = fill_paint
    elif look.no_fill:
        fill = "<a:noFill/>"
    scheme, shade = look.style_line or {"textBox": ("lt1", 50000), "line": ("accent1", 0)}.get(shape.kind, ("accent1", 15000))
    if look.line_scheme is not None or look.line_plain:
        shade = 0
    own_line = (look.line_rgb is not None or look.line_scheme is not None or bool(look.line_mods) or look.line_plain
                or look.line_index is not None)
    paint = _line_paint(look, look.line_scheme or scheme, shade, own_line)
    line = ""
    if text_box or look.line_full or look.line_ends:
        width = (look.frozen_width if look.line_hidden else look.line_width) or look.style_width or (9525 if text_box else 19050)
        dash = (look.frozen_dash if look.line_hidden else look.line_dash) or "solid"
        head, tail = (look.frozen_head, look.frozen_tail) if look.line_hidden else (look.line_head, look.line_tail)
        line = _stroke("a:ln", "", width, "<a:noFill/>" if look.line_hidden else paint, dash, text_box, look.line_ends,
                       _arrow_ends(head, tail, look.line_ends), look.line_align)
    elif (own_line or look.line_width is not None or look.line_dash is not None or look.line_pattern is not None
          or look.line_open or look.line_align is not None or any(look.line_head) or any(look.line_tail)):
        attributes = ("" if look.line_width is None else f' w="{look.line_width}"') + (
            "" if look.line_cap is None else f' cap="{look.line_cap}"') + (
            "" if look.line_align is None else f' algn="{look.line_align}"')
        inside = (paint if own_line or look.line_pattern is not None or look.line_open else "") + (
            "" if look.line_dash is None else f'<a:prstDash val="{look.line_dash}"/>') + _arrow_ends(
            look.line_head, look.line_tail, False)
        line = f"<a:ln{attributes}>{inside}</a:ln>" if inside else f"<a:ln{attributes}/>"
    extensions = ""
    if look.fill_hidden:
        extensions += (f'<a:ext uri="{{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}}"><a14:hiddenFill {_A14}>'
                       f"{fill_paint}</a14:hiddenFill></a:ext>")
    if look.line_hidden and not look.line_lost:
        width = look.line_width or look.style_width or (9525 if text_box else 19050)
        extensions += ('<a:ext uri="{91240B29-F687-4F45-9708-019B960494DF}">'
                       + _stroke("a14:hiddenLine", f" {_A14}", width, paint, look.line_dash or "solid", text_box, False,
                                 _arrow_ends(look.line_head, look.line_tail, False), None)
                       + "</a:ext>")
    effect_list, hidden = _effect_lists(look.extra)
    extensions += hidden
    return (fill + line + (effect_list or ("<a:effectLst/>" if look.effects else ""))
            + (f"<a:extLst>{extensions}</a:extLst>" if extensions else ""))


def _effect_colour(rgb: str | None, scheme: str | None, alpha: int | None) -> str:
    inner = "" if alpha is None else f'<a:alpha val="{alpha}"/>'
    if scheme:
        return f'<a:schemeClr val="{scheme}">{inner}</a:schemeClr>' if inner else f'<a:schemeClr val="{scheme}"/>'
    if rgb:
        return f'<a:srgbClr val="{rgb}">{inner}</a:srgbClr>' if inner else f'<a:srgbClr val="{rgb}"/>'
    black = 'r="0" g="0" b="0"'
    return f"<a:scrgbClr {black}>{inner}</a:scrgbClr>" if inner else f"<a:scrgbClr {black}/>"


def _shadow_markup(effects: ShapeEffects) -> str:
    size = "" if effects.shadow_size is None else f' sx="{effects.shadow_size}" sy="{effects.shadow_size}"'
    distance = f' dist="{effects.shadow_distance}"' if effects.shadow_distance else ""
    direction = f' dir="{effects.shadow_direction}"' if effects.shadow_direction else ""
    return (f'<a:outerShdw blurRad="{effects.shadow_blur}"{distance}{direction}{size} rotWithShape="0">'
            f"{_effect_colour(effects.shadow_rgb, None, effects.shadow_alpha)}</a:outerShdw>")


def _effect_lists(effects: ShapeEffects | None) -> tuple[str, str]:
    """The shape's effectLst, and the extensions for a shadow changed while hidden or obscured, each as the file has
    them (tests/fixtures/shape_effects.json)."""
    if effects is None:
        return "", ""
    glow = ""
    if effects.glow:
        radius = "" if effects.glow_radius is None else f' rad="{effects.glow_radius}"'
        glow = (f"<a:glow{radius}>{_effect_colour(effects.glow_rgb, effects.glow_scheme, effects.glow_alpha)}"
                "</a:glow>")
    shadow = _shadow_markup(effects) if effects.shadow_on else ""
    others = ""
    if effects.reflection:
        attributes = "".join(f' {name}="{value}"' for name, value in (
            ("blurRad", effects.reflection_blur), ("stA", effects.reflection_start), ("endA", effects.reflection_end),
            ("endPos", effects.reflection_place), ("dist", effects.reflection_distance)) if value is not None)
        others += f'<a:reflection{attributes} dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/>'
    if effects.soft_edge:
        others += f'<a:softEdge rad="{effects.soft_edge}"/>'
    inside = glow + (shadow if effects.shadow_visible else "") + others
    hidden = ""
    if effects.shadow_on and not effects.shadow_visible:
        # The hidden list has the rest of them too.
        hidden += ('<a:ext uri="{AF507438-7753-43E0-B8FC-AC1667EBCBE1}">'
                   f"<a14:hiddenEffects {_A14}><a:effectLst>{glow}{shadow}{others}</a:effectLst>"
                   "</a14:hiddenEffects></a:ext>")
    if effects.shadow_obscured:
        hidden += f'<a:ext uri="{{53640926-AAD7-44D8-BBD7-CCE9431645EC}}"><a14:shadowObscured {_A14} val="1"/></a:ext>'
    written = effects.list_only or effects.shadow_on or effects.glow or effects.reflection or effects.soft_edge
    touched = written or effects.reflection_type or effects.touched
    return (f"<a:effectLst>{inside}</a:effectLst>" if inside else "<a:effectLst/>") if touched else "", hidden


_A14 = 'xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"'


def _colour(rgb: str | None, scheme: str, shade: int, alpha: int | None,
            mods: tuple[tuple[str, int], ...] = (), index: int | None = None) -> str:
    """A colour of the file's: its own, else a theme colour at its shade, then what Brightness and TintAndShade set,
    either at its alpha. A SchemeColor is written as its RGB with the palette index beside it."""
    inside = (f'<a:shade val="{shade}"/>' if shade and rgb is None else "") + "".join(
        f'<a:{name} val="{value}"/>' for name, value in mods) + (
        "" if alpha is None else f'<a:alpha val="{alpha}"/>')
    tag, value = ("a:srgbClr", rgb) if rgb is not None else ("a:schemeClr", scheme)
    legacy = "" if index is None else (
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="a14" '
        f'{_A14} a14:legacySpreadsheetColorIndex="{index}"')
    return f'<{tag} val="{value}"{legacy}>{inside}</{tag}>' if inside else f'<{tag} val="{value}"{legacy}/>'


#: What a corner gradient's path and tile hold for each variant.
_CORNERS = {1: ('r="100000" b="100000"', 'l="-100000" t="-100000"'), 2: ('l="100000" b="100000"', 't="-100000" r="-100000"'),
            3: ('t="100000" r="100000"', 'l="-100000" b="-100000"'), 4: ('l="100000" t="100000"', 'r="-100000" b="-100000"')}


def _gradient(gradient: FillGradient) -> str:
    """A gradient fill as Excel writes it (tests/fixtures/shape_fill_gradients.json)."""
    stops = "".join(f'<a:gs pos="{stop.position}">{_colour(stop.rgb, stop.scheme or "", 0, stop.alpha, stop.mods)}</a:gs>'
                    for stop in gradient.stops)
    if gradient.path is None:
        shape = f'<a:lin ang="{gradient.angle}" scaled="1"/><a:tileRect/>'
    elif gradient.corner is None:
        shape = '<a:path path="rect"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path><a:tileRect/>'
    else:
        to, tile = _CORNERS[gradient.corner]
        shape = f'<a:path path="rect"><a:fillToRect {to}/></a:path><a:tileRect {tile}/>'
    return (f'<a:gradFill flip="none" rotWithShape="{1 if gradient.rotate else 0}"><a:gsLst>{stops}</a:gsLst>{shape}'
            "</a:gradFill>")


def _line_paint(look: FillLine, scheme: str, shade: int, own: bool) -> str:
    """What a line is painted with: a pattern with the colours set on it, an empty solidFill where a back colour had no
    pattern to take it, or its colour."""
    if look.line_pattern is not None:
        if look.line_alpha is not None:
            # Its transparency writes out the default black on white.
            alpha = f'<a:alpha val="{look.line_alpha}"/>'
            return (f'<a:pattFill prst="{look.line_pattern}"><a:fgClr><a:prstClr val="black">{alpha}</a:prstClr></a:fgClr>'
                    f'<a:bgClr><a:prstClr val="white">{alpha}</a:prstClr></a:bgClr></a:pattFill>')
        fore = "" if not own else f"<a:fgClr>{_colour(look.line_rgb, scheme, 0, None, look.line_mods)}</a:fgClr>"
        back = ""
        if look.line_back_rgb is not None or look.line_back_scheme is not None:
            back = f'<a:bgClr>{_colour(look.line_back_rgb, look.line_back_scheme or "", 0, None)}</a:bgClr>'
        inside = fore + back
        return (f'<a:pattFill prst="{look.line_pattern}">{inside}</a:pattFill>' if inside
                else f'<a:pattFill prst="{look.line_pattern}"/>')
    if look.line_open and not own:
        return "<a:solidFill/>"
    colour = _colour(look.line_rgb, scheme, shade, look.line_alpha, look.line_mods, look.line_index)
    return f"<a:solidFill>{colour}</a:solidFill>"


def _arrow_ends(head: tuple[str | None, str | None, str | None], tail: tuple[str | None, str | None, str | None],
                full: bool) -> str:
    """A line's headEnd and tailEnd: what was set, or all of each once the line is written out with its ends."""
    written = ""
    for tag, (kind, wide, long) in (("a:headEnd", head), ("a:tailEnd", tail)):
        if full:
            kind, wide, long = kind or "none", wide or "med", long or "med"
        attributes = "".join(f' {name}="{value}"' for name, value in (("type", kind), ("w", wide), ("len", long))
                             if value is not None)
        written += f"<{tag}{attributes}/>" if attributes else ""
    return written


def _stroke(tag: str, namespace: str, width: int, paint: str, dash: str, text_box: bool, full: bool, ends: str,
            align: str | None) -> str:
    """A line written out in full: as a text box has it, with no join until its ends are written and a dash only if one is set
    (tests/fixtures/shape_settings_kept.json), or as the style gives any other shape's."""
    if text_box and not full:
        aligned = "" if align is None else f' algn="{align}"'
        dashed = "" if dash == "solid" else f'<a:prstDash val="{dash}"/>'
        return f'<{tag}{namespace} w="{width}" cmpd="sng"{aligned}>{paint}{dashed}</{tag}>'
    join = "<a:round/>" if text_box else '<a:miter lim="800000"/>'
    return (f'<{tag}{namespace} w="{width}" cap="flat" cmpd="sng" algn="ctr">{paint}<a:prstDash val="{dash}"/>{join}'
            f"{ends}</{tag}>")


def _free_geometry(shape: Shape) -> str:
    """A freeform's custGeom: its path as wide and tall as the shape, in EMU (tests/fixtures/shape_polyline.json)."""
    steps = "".join(
        "<a:close/>" if command == "close" else
        f"<a:{command}>" + "".join(f'<a:pt x="{x}" y="{y}"/>' for x, y in points) + f"</a:{command}>"
        for command, points in shape.path)
    return ('<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="0" b="0"/><a:pathLst>'
            f'<a:path w="{emu(shape.width)}" h="{emu(shape.height)}">{steps}</a:path></a:pathLst></a:custGeom>')


def _connections(shape: Shape) -> str:
    """A connector's cNvCxnSpPr, naming the shape and site each connected end holds (tests/fixtures/shape_connectors.json)."""
    inside = '<a:cxnSpLocks noChangeAspect="1"/>' if shape.lock_aspect else "<a:cxnSpLocks/>" if shape.link_lost or shape.lock_aspect is False else ""
    inside += "" if shape.begin_link is None else f'<a:stCxn id="{shape.begin_link.target}" idx="{shape.begin_link.site - 1}"/>'
    inside += "" if shape.end_link is None else f'<a:endCxn id="{shape.end_link.target}" idx="{shape.end_link.site - 1}"/>'
    return f"<xdr:cNvCxnSpPr>{inside}</xdr:cNvCxnSpPr>" if inside else "<xdr:cNvCxnSpPr/>"


def _text_line(line: TextLine) -> str:
    """The a:ln of a shape's letters (tests/fixtures/shape_font2_paint.json)."""
    if line.hidden:
        return "<a:ln><a:noFill/></a:ln>"
    if line.rgb is None and line.scheme is None:
        inside = "".join(f'<a:{name} val="{value}"/>' for name, value in line.mods) + ("" if line.alpha is None else f'<a:alpha val="{line.alpha}"/>')
        ink = f'<a:prstClr val="black">{inside}</a:prstClr>' if inside else '<a:prstClr val="black"/>'
    else:
        ink = _colour(line.rgb, line.scheme or "", 0, line.alpha, line.mods)
    if line.full:
        return (f'<a:ln w="{9525 if line.width is None else line.width}" cap="flat" cmpd="{line.compound or "sng"}" algn="ctr"><a:solidFill>{ink}</a:solidFill>'
                f'<a:prstDash val="{line.dash or "solid"}"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:ln>')
    attributes = ("" if line.width is None else f' w="{line.width}"') + ("" if line.compound is None else f' cmpd="{line.compound}"')
    dash = "" if line.dash is None else f'<a:prstDash val="{line.dash}"/>'
    return f"<a:ln{attributes}><a:solidFill>{ink}</a:solidFill>{dash}</a:ln>"


def _font_markup(look: TextLook) -> tuple[str, str]:
    """What Characters.Font writes on a run or a paragraph's end: its attributes and its children."""
    attributes = f' sz="{round(look.font_size * 100)}"'
    if look.font_bold is not None:
        attributes += f' b="{1 if look.font_bold else 0}"'
    if look.font_italic is not None:
        attributes += f' i="{1 if look.font_italic else 0}"'
    if look.font_underline is not None:
        attributes += f' u="{look.font_underline}"'
    if look.font_strike is not None:
        attributes += f' strike="{look.font_strike}"'
    if look.font_kern is not None:
        attributes += f' kern="{look.font_kern}"'
    if look.font_cap is not None:
        attributes += f' cap="{look.font_cap}"'
    if look.font_spacing is not None:
        attributes += f' spc="{look.font_spacing}"'
    if look.font_normalize:
        attributes += ' normalizeH="1"'
    if look.font_baseline is not None:
        attributes += f' baseline="{look.font_baseline}"'
    children = ""
    if look.font_line is not None:
        children += _text_line(look.font_line)
    mods = (("tint", look.font_tint),) if look.font_tint is not None else look.font_mods
    if look.font_no_fill:
        children += "<a:noFill/>"
    elif look.font_rgb is not None:
        children += f'<a:solidFill>{_colour(look.font_rgb, "", 0, look.font_alpha, mods, look.font_index)}</a:solidFill>'
    elif look.font_theme is not None or mods or look.font_alpha is not None or look.font_solid:
        children += f'<a:solidFill>{_colour(None, look.font_theme or look.font_scheme, 0, look.font_alpha, mods)}</a:solidFill>'
    if look.font_shadow:
        children += ('<a:effectLst><a:outerShdw blurRad="63500" dist="37357" dir="2700000" rotWithShape="0">'
                     '<a:scrgbClr r="0" g="0" b="0"/></a:outerShdw></a:effectLst>')
    if look.font_effects is not None:
        children += look.font_effects
    children += (look.font_highlight or "") + (look.font_underline_line or "") + (look.font_underline_fill or "")
    children += (look.font_latin or "") + (look.font_ea or "") + (look.font_cs or "") + (look.font_sym or "")
    return attributes, children


def _run_markup(look: TextLook, mark: dict[str, object]) -> tuple[str, str]:
    return _font_markup(replace(look, **mark))  # pyright: ignore[reportArgumentType]


def text_run_spans(shape: Shape) -> list[tuple[int, int]]:
    """Where each run of a new shape's text starts, and its length: characters alike in what they write run together,
    across a paragraph's end too (tests/fixtures/shape_runs.json)."""
    spans: list[tuple[int, int]] = []
    last: tuple[str, str] | None = None
    marks: list[dict[str, object]] = shape.text_marks or [{} for _ in shape.text]
    for at, mark in enumerate(marks):
        markup = _run_markup(shape.text_look or TextLook(), mark)
        if spans and markup == last:
            spans[-1] = (spans[-1][0], spans[-1][1] + 1)
        else:
            spans.append((at, 1))
        last = markup
    return spans


def _props(tag: str, markup: tuple[str, str]) -> str:
    return f'<{tag} lang="en-US"{markup[0]}>{markup[1]}</{tag}>' if markup[1] else f'<{tag} lang="en-US"{markup[0]}/>'


def _marked_paragraphs(shape: Shape, look: TextLook) -> str | None:
    """A text's paragraphs with their runs, when some of it has settings of its own (tests/fixtures/shape_runs.json).

    A paragraph's end is written only where it holds settings its last run does not write, and an empty paragraph always."""
    if not shape.text_marks and shape.text_end is None:
        return None
    marks: list[dict[str, object]] = shape.text_marks or [{} for _ in shape.text]
    paragraphs: list[str] = []
    at = 0
    lines = shape.text.split("\n")
    for index, line in enumerate(lines):
        runs = ""
        written: tuple[str, str] | None = None
        pending = ""
        for offset, character in enumerate(line):
            markup = _run_markup(look, marks[at + offset])
            if markup != written and pending:
                runs += f'<a:r>{_props("a:rPr", written or markup)}<a:t>{dml.escaped_text(pending)}</a:t></a:r>'
                pending = ""
            written = markup
            pending += character
        if pending:
            runs += f'<a:r>{_props("a:rPr", written or _font_markup(look))}<a:t>{dml.escaped_text(pending)}</a:t></a:r>'
        end = marks[at + len(line)] if index < len(lines) - 1 else shape.text_end
        if not line or (end is not None and (end or index == len(lines) - 1) and _run_markup(look, end) != written):
            runs += _props("a:endParaRPr", _run_markup(look, end or {}))
        paragraphs.append(f"<a:p>{runs}</a:p>")
        at += len(line) + 1
    return "".join(paragraphs)


def _autofit(which: str) -> str:
    return "<a:spAutoFit/>" if which == "sp" else f"<a:{which}Autofit/>"


def _laid_out(body: str, look: TextLook, paragraphs: str | None = None) -> str:
    """A new shape's txBody with the layout TextFrame gave it: bodyPr's attributes in the schema's order, and each
    paragraph's algn and rtl (tests/fixtures/text_frame.json)."""
    insets = "".join(f' {name}="{emu(value)}"' for name, value in zip(("lIns", "tIns", "rIns", "bIns"), look.margins)
                     if value is not None)
    overflow = f'vertOverflow="{look.vertical_overflow}" horzOverflow="{look.horizontal_overflow}"'
    if look.full:
        # Every attribute, each margin spelled out, as Excel writes an AutoShape whose overflow has changed.
        insets = "".join(f' {name}="{emu(value if value is not None else default)}"' for name, value, default in zip(
            ("lIns", "tIns", "rIns", "bIns"), look.margins, (7.2, 3.6, 7.2, 3.6)))
        upright = ' upright="1"' if look.upright else ""
        head = (f'<a:bodyPr rot="0" spcFirstLastPara="0" {overflow} vert="{look.vert or "horz"}" wrap="{look.wrap or "square"}"'
                f'{insets} numCol="{look.columns}" spcCol="{emu(look.column_spacing)}" rtlCol="0" fromWordArt="0" anchor="{look.anchor}" anchorCtr="{1 if look.anchor_center else 0}"'
                f' forceAA="0"{upright} compatLnSpc="1"><a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp>'
                f"{_autofit(look.autofit or 'no')}</a:bodyPr>")
    else:
        vert = f' vert="{look.vert}"' if look.vert else ""
        wrap = f' wrap="{look.wrap}"' if look.wrap else ""
        centred = "" if look.anchor_center is None else f' anchorCtr="{1 if look.anchor_center else 0}"'
        upright = ' upright="1"' if look.upright else ""
        columns = ("" if look.columns == 1 else f' numCol="{look.columns}"') + (
            f' spcCol="{emu(look.column_spacing)}"' if look.column_spacing else "")
        attributes = f'{overflow}{vert}{wrap}{insets}{columns} rtlCol="0" anchor="{look.anchor}"{centred}{upright}'
        # TextFrame2's AutoSize is bodyPr's one child (tests/fixtures/text_frame2.json).
        head = (f"<a:bodyPr {attributes}>{_autofit(look.autofit)}</a:bodyPr>" if look.autofit
                else f"<a:bodyPr {attributes}/>")
    body = re.sub(r"<a:bodyPr\b[^>]*/>", lambda _: head, body, count=1)
    if look.deleted:
        body = body.replace('<a:endParaRPr lang="en-US" sz="1100"/>', '<a:endParaRPr lang="en-US"/>')
    attributes, children = _font_markup(look)
    if attributes != ' sz="1100"' or children:
        for tag in ("a:rPr", "a:endParaRPr"):
            body = body.replace(f'<{tag} lang="en-US" sz="1100"/>', f'<{tag} lang="en-US"{attributes}>{children}</{tag}>'
                                if children else f'<{tag} lang="en-US"{attributes}/>')
    if paragraphs is not None:
        body = re.sub(r"<a:p>.*</a:p>", lambda _: paragraphs, body, flags=re.DOTALL)
    settings = (f' algn="{look.align}"' if look.align else "") + (
        "" if look.rtl is None else f' rtl="{1 if look.rtl else 0}"')
    paragraph = f"<a:p><a:pPr{settings}/>" if settings else "<a:p>"
    return re.sub(r"<a:p>(?:<a:pPr\b[^>]*/>)?", lambda _: paragraph, body)


def _identity(shape: Shape) -> str:
    """A new shape's cNvPr: its id, name, alternative text and title, and the creation id Excel gives every shape. An elbow or
    curve that keeps Excel's own name reads as "Elbow Connector 2" but is saved as "Connector: Elbow 2"
    (tests/fixtures/shape_connectors.json)."""
    described = f' descr="{escape(shape.description)}"' if shape.description else ""
    titled = f' title="{escape(shape.title)}"' if shape.title else ""
    creation = "{" + str(uuid.uuid4()).upper() + "}"
    extensions = (
        '<a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">'
        f'<a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{creation}"/>'
        "</a:ext>"
    )
    if shape.decorative:
        extensions += (
            '<a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}">'
            '<adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/>'
            "</a:ext>"
        )
    hidden = ' hidden="1"' if shape.hidden else ""
    name = shape.name
    if shape.kind == "freeform":
        # A freeform Excel named reads "Freeform 1" but is saved as "Freeform: Shape 1" (tests/fixtures/shape_polyline.json).
        name = re.sub(r"^Freeform (\d+)$", r"Freeform: Shape \1", name)
    if shape.kind == "line":
        name = re.sub(r"^(Elbow|Curved) Connector (\d+)$", r"Connector: \1 \2", name)
    if shape.kind == "shape":
        # So with a rounded rectangle, read as "Rounded Rectangle 1" (tests/fixtures/old_drawings.json).
        name = re.sub(r"^Rounded Rectangle (\d+)$", r"Rectangle: Rounded Corners \1", name)
    return (f'<xdr:cNvPr id="{shape.shape_id}" name="{escape(name)}"{described}{titled}{hidden}>'
            f"<a:extLst>{extensions}</a:extLst></xdr:cNvPr>")


def _corner_markup(which: str, grid: SheetGrid, across: float, down: float) -> str:
    """One corner of an anchor, as the cell it lands in."""
    column, column_offset = grid.column_at(across)
    row, row_offset = grid.row_at(down)
    return (
        f"<xdr:{which}><xdr:col>{column}</xdr:col><xdr:colOff>{column_offset}</xdr:colOff>"
        f"<xdr:row>{row}</xdr:row><xdr:rowOff>{row_offset}</xdr:rowOff></xdr:{which}>"
    )


#: The style Excel gives a new AutoShape, taken from one it wrote.
_STYLE = (
    "<xdr:style>"
    '<a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef>'
    '<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>'
    '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>'
    '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>'
    "</xdr:style>"
)
#: The style Excel gives a line, and an open freeform, drawn as one (tests/fixtures/shape_polyline.json).
_OPEN_STYLE = (
    "<xdr:style>"
    '<a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>'
    '<a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>'
    '<a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef>'
    '<a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef>'
    "</xdr:style>"
)
#: The fill, line and style Excel gives a new text box, taken from one it wrote (tests/fixtures/shape_duplicate.json).
_TEXT_BOX_FILL = ('<a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:ln w="9525" cmpd="sng"><a:solidFill>'
                  '<a:schemeClr val="lt1"><a:shade val="50000"/></a:schemeClr></a:solidFill></a:ln>')
_TEXT_BOX_STYLE = (
    "<xdr:style>"
    '<a:lnRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef>'
    '<a:fillRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef>'
    '<a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef>'
    '<a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef>'
    "</xdr:style>"
)


def written(shapes: list[Shape], original: str, grid: SheetGrid | None = None) -> str:
    """A drawing part holding these shapes.

    A shape that came from the file and was not changed is written back
    from the markup it was read from, so an untouched drawing survives
    a save byte for byte.
    """
    body = dml.element_span(original, "xdr:wsDr")
    if body is None:
        original = EMPTY_DRAWING
        body = dml.element_span(original, "xdr:wsDr")
        assert body is not None
    head = original.find(">", body[0]) + 1
    tail = original.rfind("</xdr:wsDr>")
    return original[:head] + "".join(_markup_for(one, grid) for one in shapes) + original[tail:]


def _markup_for(shape: Shape, grid: SheetGrid | None = None) -> str:
    """One shape's markup: its own if nothing about it changed."""
    if not shape.source:
        if shape.kind == "formControl":
            return control_drawing(shape, grid)
        return new_anchor(shape, grid)
    was = _shape_of_source(shape)
    markup = shape.source
    if was is None:
        return markup
    if shape.name != was.name:
        markup = dml.with_name(markup, shape.name)
    if (shape.left, shape.top, shape.width, shape.height) != (
        was.left,
        was.top,
        was.width,
        was.height,
    ):
        markup = _moved(markup, shape, was, grid)
    if shape.kind != "formControl":
        if shape.macro != was.macro:
            element = re.search(r"<(xdr:(?:sp|cxnSp|pic|graphicFrame|grpSp))(?=[\s/>])", markup)
            if element:
                markup = _attribute_on_tag(markup, element.group(1), "macro", shape.macro)
        if shape.locks_text != was.locks_text:
            markup = _attribute_on_tag(markup, "xdr:sp", "fLocksText", None if shape.locks_text else "0")
        if shape.shadow_added:
            # In the schema's place: after the fill and line, before any 3-D settings.
            properties = re.search(r"<xdr:spPr>.*?(?=<a:scene3d\b|<a:sp3d\b|<a:extLst\b|</xdr:spPr>)", markup, re.DOTALL)
            if properties is not None:
                markup = markup[:properties.end()] + OLD_SHADOW + markup[properties.end():]
    # Excel-authored controls can also carry a hidden DrawingML
    # caption. Excel prefers it to VML when both exist.
    markup = _with_loaded_text(markup, shape, was)
    if shape.kind == "formControl" and shape.control is not None:
        markup = with_text_frame(markup, shape.control)
    if shape.kind == "picture":
        markup = _with_loaded_picture(markup, shape, was)
    if shape.kind == "line" and shape.links_read:
        markup = _with_loaded_links(markup, shape)
    return _restyled(markup, shape, was, grid)


def shape_as_loaded(shape: Shape) -> Shape | None:
    """What a shape was when the file was opened, if it came from one."""
    return _shape_of_source(shape) if shape.source else None


def _shape_of_source(shape: Shape) -> Shape | None:
    """What the markup this shape came from says about it."""
    read = read_drawing(f"<xdr:wsDr>{shape.source}</xdr:wsDr>")
    if not read:
        return None
    if shape.loaded_box is not None:
        # Where the load put it, by its anchor; the markup alone, with no grid, has only its transform.
        read[0].left, read[0].top, read[0].width, read[0].height = shape.loaded_box
    return read[0]


def _moved(markup: str, shape: Shape, was: Shape, grid: SheetGrid | None = None) -> str:
    """The same markup with the shape somewhere else.

    The anchor moves with the transform: Excel reads the anchor when it
    opens the file, so leaving it behind would put the shape back where
    it was.  Both corners are worked out from the grid, because an
    offset that runs past its own cell is clamped rather than carried.
    """
    out = dml.with_transform(markup, shape)
    span = dml.element_span(out, "xdr:from")
    finish = dml.element_span(out, "xdr:to")
    if span is None:
        return out
    grid = grid or SheetGrid()
    left, top, width, height = _turned_box(shape)
    start = _corner_markup("from", grid, left, top)
    if finish is None:
        # A one-cell anchor has only this corner, and its size apart (tests/fixtures/tool_drawings.json).
        if (shape.width, shape.height) != (was.width, was.height):
            raise VBAUnsupportedError("Resizing a shape on a one-cell anchor is not implemented")
        return out[: span[0]] + start + out[span[1] :]
    end = _corner_markup("to", grid, left + width, top + height)
    return out[: span[0]] + start + out[span[1] : finish[0]] + end + out[finish[1] :]


def _turned_box(shape: Shape) -> tuple[float, float, float, float]:
    """The box a shape is anchored by: its own, or, turned nearer upright than flat (45 to 135 degrees, or 225 to
    315), the same box turned on its centre (tests/fixtures/shape_properties.json)."""
    left, top, width, height = shape.left, shape.top, shape.width, shape.height
    if 45 <= shape.rotation % 360 < 135 or 225 <= shape.rotation % 360 < 315:
        return left + (width - height) / 2, top + (height - width) / 2, height, width
    return left, top, width, height


_ADEC_URI = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}"
_ADEC = ('<a:ext uri="' + _ADEC_URI + '"><adec:decorative '
         'xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="{value}"/></a:ext>')


def _with_lock(markup: str, shape: Shape) -> str:
    """A loaded shape's markup with its aspect locked or not, in the locks element of its kind, first in its properties
    (tests/fixtures/shape_settings_kept.json). What else the element locks stays."""
    tag, holder = {"line": ("a:cxnSpLocks", "xdr:cNvCxnSpPr"), "picture": ("a:picLocks", "xdr:cNvPicPr")}.get(shape.kind, ("a:spLocks", "xdr:cNvSpPr"))
    found = re.search(rf"<{tag}\b([^>]*?)/>", markup)
    kept = "" if found is None else re.sub(r'\s*noChangeAspect="[^"]*"', "", found.group(1)).rstrip()
    locks = f'<{tag} noChangeAspect="1"{kept}/>' if shape.lock_aspect else f"<{tag}{kept}/>"
    if found is not None:
        return markup[:found.start()] + locks + markup[found.end():]
    opened = re.search(rf"<{holder}([^>]*?)(/?)>", markup)
    if opened is None:
        return markup
    if opened.group(2):
        return markup[:opened.start()] + f"<{holder}{opened.group(1)}>{locks}</{holder}>" + markup[opened.end():]
    return markup[:opened.end()] + locks + markup[opened.end():]


def _restyled(markup: str, shape: Shape, was: Shape, grid: SheetGrid | None) -> str:
    """A loaded shape's markup with the settings a macro changed patched in, as Excel patches them
    (tests/fixtures/loaded_shape/cases.json): descr and title on cNvPr, the decorative extension (val 0 once
    turned off, as Excel leaves it), xfrm's rot and the anchor it turns, spLocks, and clientData's fLocksWithSheet."""
    if (shape.description, shape.title, shape.hidden) != (was.description, was.title, was.hidden):
        found = _CNVPR_OPEN.search(markup)
        if found is not None:
            attributes = re.sub(r'\s(?:descr|title|hidden)="[^"]*"', "", found.group(1))
            name = re.search(r'\bname="[^"]*"', attributes)
            at = name.end() if name is not None else len(attributes)
            added = (f' descr="{escape(shape.description)}"' if shape.description else "") + (
                f' title="{escape(shape.title)}"' if shape.title else "") + (' hidden="1"' if shape.hidden else "")
            attributes = attributes[:at] + added + attributes[at:]
            markup = markup[: found.start(1)] + attributes + markup[found.end(1):]
    if shape.decorative != was.decorative:
        value = "1" if shape.decorative else "0"
        if re.search(r"<adec:decorative\b", markup):
            markup = re.sub(r'(<adec:decorative\b[^>]*\bval=")[^"]*"', rf'\g<1>{value}"', markup, count=1)
        elif "</a:extLst></xdr:cNvPr>" in markup:
            markup = markup.replace("</a:extLst></xdr:cNvPr>", _ADEC.replace("{value}", value)
                                    + "</a:extLst></xdr:cNvPr>", 1)
    if shape.fill_line is not None:
        markup = with_fill_line(markup, _fill_line(shape))
    # A control's VML says False where its drawing has no spLocks to say anything, and None; either way the aspect is free.
    free = shape.kind == "formControl" and not shape.lock_aspect and not was.lock_aspect
    if shape.lock_aspect != was.lock_aspect and not free:
        markup = _with_lock(markup, shape)
    # A control keeps its Locked and PrintObject in the sheet and its VML; Excel leaves its drawing's clientData as it is.
    if (shape.locked, shape.prints) != (was.locked, was.prints) and shape.kind != "formControl":
        markup = re.sub(r"<xdr:clientData\b[^>]*/>", lambda _: _client_data(shape), markup, count=1)
    if shape.edit_as != was.edit_as:
        # tests/fixtures/chart_object_more.json
        tag = f'<xdr:twoCellAnchor editAs="{shape.edit_as}">' if shape.edit_as else "<xdr:twoCellAnchor>"
        markup = re.sub(r"<xdr:twoCellAnchor\b[^>]*>", lambda _: tag, markup, count=1)
    if (shape.rotation, shape.flip_h, shape.flip_v) != (was.rotation, was.flip_h, was.flip_v):
        found = re.search(r"<a:xfrm\b[^>]*>", markup)
        if found is not None:
            tag = re.sub(r'\s(?:rot|flipH|flipV)="[^"]*"', "", found.group())
            tag = tag.replace("<a:xfrm", "<a:xfrm" + _xfrm_attributes(shape), 1)
            markup = markup[: found.start()] + tag + markup[found.end():]
        if _turned_box(shape) != _turned_box(was):
            markup = _moved(markup, shape, was, grid)
    return markup


def _client_data(shape: Shape) -> str:
    """The anchor's clientData: fLocksWithSheet and fPrintsWithSheet, each only where it is off."""
    return ("<xdr:clientData" + ("" if shape.locked else ' fLocksWithSheet="0"')
            + ("" if shape.prints else ' fPrintsWithSheet="0"') + "/>")


def _xfrm_attributes(shape: Shape) -> str:
    """The xfrm's rot, flipH and flipV, in that order, each only where it is set (tests/fixtures/shape_flip_order.json)."""
    return ((f' rot="{round(shape.rotation * 60000)}"' if shape.rotation else "")
            + (' flipH="1"' if shape.flip_h else "") + (' flipV="1"' if shape.flip_v else ""))


def geometry_for(mso_type: int) -> str:
    """The preset geometry an AddShape asked for."""
    return PRESET_GEOMETRY.get(mso_type, "rect")
