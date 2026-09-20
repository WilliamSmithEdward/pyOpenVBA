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

import re
from dataclasses import dataclass, field

from pyopenvba._xml import escape, unescape
from pyopenvba.shapes import _drawingml as dml
from pyopenvba.shapes._values import PRESET_GEOMETRY, ControlInfo, Shape, emu, points

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
            if seen + width > across or width <= 0:
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
            if seen + height > down or height <= 0:
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
        fields = dict(_ATTRIBUTE.findall(head.group(1)))
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
    return shape


def _shape_of(
    markup: str, element: str, kind: str, anchor: str, grid: SheetGrid | None
) -> Shape:
    shape_id, name = dml.identity(markup)
    left, top, width, height = dml.transform(markup)
    if not width and not height:
        left, top, width, height = _anchor_box(anchor, grid) or (left, top, width, height)
    if kind == "shape" and _TXBOX.search(markup):
        kind = "textBox"
    macro = _MACRO.search(markup)
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
        shape_id=shape_id,
        source=anchor,
    )
    if kind == "group":
        shape.children = _members(markup, grid)
    return shape


def _anchor_box(anchor: str, grid: SheetGrid | None) -> tuple[float, float, float, float] | None:
    """Where the anchor says the shape is, in points.

    Used for a shape whose own transform is empty, which is how Excel
    writes a form control.
    """
    if grid is None:
        return None
    start = dml.element_span(anchor, "xdr:from")
    finish = dml.element_span(anchor, "xdr:to")
    if start is None:
        return None
    left, top = _corner(anchor[start[0] : start[1]], grid)
    if finish is None:
        return left, top, 0.0, 0.0
    right, bottom = _corner(anchor[finish[0] : finish[1]], grid)
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
            control.macro_text = dict(_ATTRIBUTE.findall(properties.group(1))).get("macro", "")
        part = parts.get(relationship, "")
        if part:
            settings = _FORM_CONTROL_PR.search(part)
            if settings is not None:
                values = dict(_ATTRIBUTE.findall(settings.group(1)))
                control.kind = values.get("objectType", "control")
                control.linked_cell = values.get("fmlaLink", "")
                control.list_range = values.get("fmlaRange", "")
                control.value = int(_number(values.get("val"), 0))
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
        return dict(_ATTRIBUTE.findall(properties.group(1))).get("macro", "")
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
            changed = re.sub(r'\smacro="[^"]*"', f' macro="{escape(macro)}"' if macro else "", attributes)
        elif macro:
            changed = attributes + f' macro="{escape(macro)}"'
        else:
            changed = attributes
        start = found.start() + properties.start(1)
        return sheet_xml[:start] + changed + sheet_xml[found.start() + properties.end(1) :]
    return sheet_xml


#: What a form control's shape id starts at.  Excel numbers controls
#: from 1025, apart from the drawing shapes, which start at 2.
FIRST_CONTROL_ID = 1025

#: The VML a sheet needs before it can hold a control at all: the id
#: map and the shape type every button is drawn from.
EMPTY_VML = (
    '<xml xmlns:v="urn:schemas-microsoft-com:vml"\r\n'
    ' xmlns:o="urn:schemas-microsoft-com:office:office"\r\n'
    ' xmlns:x="urn:schemas-microsoft-com:office:excel">\r\n'
    ' <o:shapelayout v:ext="edit">\r\n'
    '  <o:idmap v:ext="edit" data="1"/>\r\n'
    ' </o:shapelayout><v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201"\r\n'
    '  path="m,l,21600r21600,l21600,xe">\r\n'
    '  <v:stroke joinstyle="miter"/>\r\n'
    '  <v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/>\r\n'
    '  <o:lock v:ext="edit" shapetype="t"/>\r\n'
    " </v:shapetype></xml>\r\n"
)

#: What each control this can make is called in its own part, and what
#: ClientData says it is.
CONTROL_KINDS: dict[str, str] = {
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
    return (
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        '<mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">'
        '<xdr:twoCellAnchor editAs="oneCell">'
        f"{_corner_markup('from', grid, shape.left, shape.top)}"
        f"{_corner_markup('to', grid, shape.left + shape.width, shape.top + shape.height)}"
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
        f'<controlPr defaultSize="0" autoFill="0" autoPict="0"{macro}>'
        '<anchor moveWithCells="1">'
        f"{_corner_markup('from', grid, shape.left, shape.top).replace('xdr:from', 'from')}"
        f"{_corner_markup('to', grid, shape.left + shape.width, shape.top + shape.height).replace('xdr:to', 'to')}"
        "</anchor>"
        "</controlPr>"
        "</control>"
        "</mc:Choice>"
        "</mc:AlternateContent>"
    )


def control_properties(shape: Shape) -> str:
    """A control's own part: which control it is and what it is wired to."""
    control = shape.control
    kind = CONTROL_KINDS.get(control.kind if control else "Button", "Button")
    linked = f' fmlaLink="{escape(control.linked_cell)}"' if control and control.linked_cell else ""
    listed = f' fmlaRange="{escape(control.list_range)}"' if control and control.list_range else ""
    extra = ' dropStyle="combo" dropLines="8"' if kind == "Drop" else ""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"'
        f' objectType="{kind}"{linked}{listed}{extra} lockText="1"/>'
    )


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
    kind = CONTROL_KINDS.get(control.kind if control else "Button", "Button")
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
        f"<div style='text-align:center'>{escape(shape.text)}</div></v:textbox>"
        if shape.text
        else ""
    )
    return (
        f'<v:shape id="{_vml_id(shape.name)}" o:spid="_x0000_s{shape.shape_id}"'
        ' type="#_x0000_t201"'
        f" style='position:absolute;margin-left:{shape.left:g}pt;margin-top:{shape.top:g}pt;"
        f"width:{shape.width:g}pt;height:{shape.height:g}pt;z-index:1;mso-wrap-style:tight'"
        ' o:button="t" fillcolor="buttonFace [67]" o:insetmode="auto">'
        '<v:fill color2="buttonFace [67]" o:detectmouseclick="t"/>'
        '<o:lock v:ext="edit" rotation="t"/>'
        f"{text}"
        f'<x:ClientData ObjectType="{kind}">'
        f"<x:Anchor>{control_anchor(shape, grid)}</x:Anchor>"
        "<x:PrintObject>False</x:PrintObject>"
        "<x:AutoFill>False</x:AutoFill>"
        f"{macro}{linked}{listed}"
        "<x:TextHAlign>Center</x:TextHAlign>"
        "<x:TextVAlign>Center</x:TextVAlign>"
        "</x:ClientData>"
        "</v:shape>"
    )


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
            _FMLA_MACRO.sub(f"<x:FmlaMacro>{escape(macro)}</x:FmlaMacro>", markup, count=1)
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
        if f'o:spid="{spid}"' in vml[span[0] : span[1]]:
            return span


def new_anchor(shape: Shape, grid: SheetGrid | None = None) -> str:
    """The markup for a shape this library is adding.

    A twoCellAnchor holding the absolute transform as well as the
    cells.  The cells are what Excel reads: it clamps an offset to the
    cell that holds it, so the corner each one lands in is worked out
    from the sheet's own columns and rows.
    """
    grid = grid or SheetGrid()
    line = shape.kind == "line"
    body = "" if line else dml.text_body(shape.text, "xdr:txBody")
    macro = f' macro="{escape(shape.macro)}"' if shape.macro else ' macro=""'
    element = "xdr:cxnSp" if line else "xdr:sp"
    box = ' txBox="1"' if shape.kind == "textBox" else ""
    textlink = "" if line else ' textlink=""'
    identity = f'<xdr:cNvPr id="{shape.shape_id}" name="{escape(shape.name)}"/>'
    properties = (
        f"<xdr:nvCxnSpPr>{identity}<xdr:cNvCxnSpPr/></xdr:nvCxnSpPr>"
        if line
        else f"<xdr:nvSpPr>{identity}<xdr:cNvSpPr{box}/></xdr:nvSpPr>"
    )
    inner = (
        f"<{element}{macro}{textlink}>"
        f"{properties}"
        "<xdr:spPr><a:xfrm>"
        f'<a:off x="{emu(shape.left)}" y="{emu(shape.top)}"/>'
        f'<a:ext cx="{emu(shape.width)}" cy="{emu(shape.height)}"/>'
        f'</a:xfrm><a:prstGeom prst="{shape.geometry or "rect"}"><a:avLst/></a:prstGeom></xdr:spPr>'
        f"{'' if shape.kind == 'textBox' else _STYLE}{body}"
        f"</{element}>"
    )
    return (
        "<xdr:twoCellAnchor>"
        f"{_corner_markup('from', grid, shape.left, shape.top)}"
        f"{_corner_markup('to', grid, shape.left + shape.width, shape.top + shape.height)}"
        f"{inner}<xdr:clientData/></xdr:twoCellAnchor>"
    )


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
            markup = _MACRO.sub(f'macro="{escape(shape.macro)}"', markup, count=1)
        if shape.text != was.text:
            markup = dml.replace_text(markup, shape.text, "xdr:txBody")
    return markup


def _shape_of_source(shape: Shape) -> Shape | None:
    """What the markup this shape came from says about it."""
    read = read_drawing(f"<xdr:wsDr>{shape.source}</xdr:wsDr>")
    return read[0] if read else None


def _moved(markup: str, shape: Shape, was: Shape, grid: SheetGrid | None = None) -> str:
    """The same markup with the shape somewhere else.

    The anchor moves with the transform: Excel reads the anchor when it
    opens the file, so leaving it behind would put the shape back where
    it was.  Both corners are worked out from the grid, because an
    offset that runs past its own cell is clamped rather than carried.
    """
    del was
    out = dml.with_transform(markup, shape)
    span = dml.element_span(out, "xdr:from")
    finish = dml.element_span(out, "xdr:to")
    if span is None or finish is None:
        return out
    grid = grid or SheetGrid()
    start = _corner_markup("from", grid, shape.left, shape.top)
    end = _corner_markup("to", grid, shape.left + shape.width, shape.top + shape.height)
    return out[: span[0]] + start + out[span[1] : finish[0]] + end + out[finish[1] :]


def geometry_for(mso_type: int) -> str:
    """The preset geometry an AddShape asked for."""
    return PRESET_GEOMETRY.get(mso_type, "rect")
