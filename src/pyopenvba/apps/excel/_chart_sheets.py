"""A chart sheet read from a file: a tab of its own holding one chart, which a macro reaches through Sheets and Charts and
Worksheets passes over (tests/fixtures/chart_sheets/chart_sheets.json). The chart is the file's, read as an embedded
chart's is; the part stays as it came, so a save keeps everything the model does not read. A workbook with one refuses
adding, moving, copying, deleting or renaming tabs for now: where they land among chart sheets is not measured."""
from __future__ import annotations

import posixpath
import re
import uuid
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._charts import CHART_TYPES, SERIES_KINDS, Chart, shown_type
from pyopenvba._xml import escape
from pyopenvba.apps.excel._charts_file import REL_CHART, read_chart, rels_part, xml_tree
from pyopenvba.apps.excel._creator import ItemDefault
from pyopenvba.apps.excel._model import ExcelObject, Worksheet
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import EMPTY, ERR_SUBSCRIPT_OUT_OF_RANGE, MISSING, NOTHING, VBAInt, error, to_text
from pyopenvba.shapes._values import Shape

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts import ChartState
    from pyopenvba.apps.excel._model import Workbook
    from pyopenvba.powerquery._opc import OpcFile

_REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
_SHAPE_ID = 1


def _related(package: OpcFile, part: str, kind: str) -> str:
    """The part ``part`` relates to by ``kind``."""
    rels = rels_part(part)
    if package.has(rels):
        for entry in xml_tree(package.read(rels)):
            target = entry.attrib.get("Target", "")
            if entry.attrib.get("Type") == kind and target:
                return target[1:] if target.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(part), target))
    return ""


#: A new chart sheet's part and drawing, as Excel writes them (tests/fixtures/chart_location.json). The anchor's size and the zoom
#: are the window's in Excel, which works them out again on opening; these are what the measuring window gave.
def _new_id() -> str:
    return "{" + str(uuid.uuid4()).upper() + "}"


_NEW_SHEET = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
              '<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
              'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
              'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr xr3" '
              'xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" '
              'xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="@uid@"><sheetPr/>'
              '<sheetViews><sheetView zoomScale="95" workbookViewId="0" zoomToFit="1"/></sheetViews>'
              '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><drawing r:id="rId1"/></chartsheet>')
_NEW_DRAWING = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
                'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:absoluteAnchor><xdr:pos x="0" y="0"/>'
                '<xdr:ext cx="8662737" cy="6286500"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="{name}"><a:extLst>'
                '<a:ext uri="@creation-uri@"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" '
                'id="@creation@"/></a:ext></a:extLst></xdr:cNvPr>'
                '<xdr:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></xdr:cNvGraphicFramePr></xdr:nvGraphicFramePr>'
                '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic>'
                '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
                '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{chart}"/>'
                '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>')
_REL_CHART_SHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"
_CT_CHART_SHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"
_CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"
#: What a sheet's name cannot hold.
_ILLEGAL = re.compile(r"[\\/?*\[\]:]")


def new_name(book: Workbook, given: object) -> str:
    """A new chart sheet's name: as given, or the next Chart number free, which counts on through the workbook in any case
    (tests/fixtures/chart_location.json). An empty or overlong name is error 5, and one taken or that a sheet cannot hold 1004."""
    taken = {tab.name.casefold() for tab in book.tabs()}
    book.chart_sheets_named += 1
    if given is MISSING:
        number = book.chart_sheets_named
        while f"chart{number}" in taken:
            number += 1
        book.chart_sheets_named = number
        return f"Chart{number}"
    book.chart_sheets_named -= 1
    wanted = to_text(given)
    if not wanted or len(wanted) > 31:
        raise error(5, "Invalid procedure call or argument")
    if wanted.casefold() in taken or _ILLEGAL.search(wanted):
        raise error(1004, "Location method of Chart class failed")
    book.chart_sheets_named += 1
    return wanted


class ChartSheet(Chart):
    """One chart sheet. Chart's members read its chart, through a worksheet no tab shows that carries the chart's state."""

    vba_type_name = "Chart"

    def __init__(self, book: Workbook, name: str, part: str, package: OpcFile, sheet_id: int, visible: int) -> None:
        numbered = book.next_sheet_id
        carrier = Worksheet(book, name)
        book.next_sheet_id = numbered
        drawing = _related(package, part, _REL_DRAWING)
        chart = _related(package, drawing, REL_CHART) if drawing else ""
        if not chart or not package.has(chart):
            raise VBAUnsupportedError("opening a workbook with a chart sheet that holds no chart is not implemented")
        carrier.chart_states[_SHAPE_ID] = read_chart(package.read(chart).decode("utf-8"), chart)
        super().__init__(carrier, Shape(name=name, kind="chart", shape_id=_SHAPE_ID))
        self.book, self.name, self.part_name, self.sheet_id, self.visible = book, name, part, sheet_id, visible
        self.xml = package.read(part).decode("utf-8", errors="replace")
        found = re.search(r'\bcodeName="([^"]*)"', self._element("sheetPr"))
        self.code_name = found.group(1) if found else ""
        #: What _io's save asks of a tab.
        self.page_print_names_owned: set[str] = set()
        carrier.chart_sheet = self

    @classmethod
    def created(cls, book: Workbook, name: str, state: ChartState, shape: Shape, place: int) -> ChartSheet:
        """A new chart sheet at ``place`` among the tabs, holding a chart with this state and shape, made active. Its parts are
        written at the next save (see write_parts). Charts.Add can make one the same way."""
        numbered = book.next_sheet_id
        carrier = Worksheet(book, name)
        book.next_sheet_id = numbered + 1
        state.part = ""
        state.dirty = True
        shape.shape_id, shape.source = _SHAPE_ID, ""
        carrier.chart_states[_SHAPE_ID] = state
        carrier.shapes_.append(shape)
        sheet = cls.__new__(cls)
        Chart.__init__(sheet, carrier, shape)
        sheet.book, sheet.name, sheet.part_name, sheet.sheet_id, sheet.visible = book, name, "", numbered, -1
        sheet.xml, sheet.code_name = _NEW_SHEET.replace("@uid@", _new_id()), ""
        sheet.page_print_names_owned = set()
        carrier.chart_sheet = sheet
        if not book.chart_sheets:
            book.file_tabs = list(book.sheets_)
        book.chart_sheets.insert(sum(isinstance(tab, ChartSheet) for tab in book.file_tabs[:place]), sheet)
        book.file_tabs.insert(place, sheet)
        book.saved = False
        sheet.Activate()
        return sheet

    def write_parts(self, package: OpcFile) -> str:
        """The parts of a chart sheet made here, and its relationship from the workbook."""
        from pyopenvba.apps.excel._charts_file import write as write_charts
        from pyopenvba.powerquery._sheets import add_content_type, add_relationship

        number = 1
        while package.has(f"xl/chartsheets/sheet{number}.xml"):
            number += 1
        self.part_name = f"xl/chartsheets/sheet{number}.xml"
        # Drawings number in tab order, and the worksheets' before this one are yet to be written
        # (tests/fixtures/chart_location.json).
        before = self.book.file_tabs[:self.book.file_tabs.index(self)]
        waiting = sum(isinstance(tab, Worksheet) and tab.drawing_dirty and not tab.drawing_part and bool(tab.shapes_) for tab in before)
        drawing = 1
        while waiting or package.has(f"xl/drawings/drawing{drawing}.xml"):
            waiting -= not package.has(f"xl/drawings/drawing{drawing}.xml")
            drawing += 1
        part = f"xl/drawings/drawing{drawing}.xml"
        package.write(part, _NEW_DRAWING.encode("utf-8"))
        add_content_type(package, part, _CT_DRAWING)
        write_charts(self.sheet, package, part)
        identifier = re.search(r'Id="([^"]+)"', package.read(rels_part(part)).decode("utf-8"))
        assert identifier is not None
        drawing_xml = _NEW_DRAWING.format(name=escape(self.shape.name), chart=identifier.group(1))
        drawing_xml = drawing_xml.replace("@creation-uri@", "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}").replace("@creation@", _new_id())
        package.write(part, drawing_xml.encode("utf-8"))
        package.write(self.part_name, self.saved_xml().encode("utf-8"))
        add_content_type(package, self.part_name, _CT_CHART_SHEET)
        add_relationship(package, rels_part(self.part_name), _REL_DRAWING, f"../drawings/drawing{drawing}.xml")
        return add_relationship(package, "xl/_rels/workbook.xml.rels", _REL_CHART_SHEET, f"chartsheets/sheet{number}.xml")

    def write_changes(self, package: OpcFile) -> None:
        """A chart made here changed since it was last written."""
        from pyopenvba.apps.excel._charts_file import markup

        state = self.state
        if state.dirty and not state.original and state.part:
            package.write(state.part, markup(self.sheet, _SHAPE_ID).encode("utf-8"))
            state.dirty = False

    @method
    def Location(self, Where: object, Name: object = MISSING) -> object:  # noqa: N803 - native named arguments
        raise VBAUnsupportedError("Moving a chart off a chart sheet is not implemented")

    def held_gone(self) -> bool:
        return not any(sheet is self for sheet in self.book.chart_sheets)

    def admit_read(self) -> None:
        if not any(sheet is self for sheet in self.book.chart_sheets) or not any(book is self.book for book in self.book.application.workbooks_.books):
            raise VBAUnsupportedError("Accessing a detached chart sheet is not implemented")

    def _tab(self, step: int) -> object:
        tabs = self.book.tabs()
        index = tabs.index(self) + step
        return tabs[index] if 0 <= index < len(tabs) else NOTHING

    def _element(self, element: str) -> str:
        found = re.search(rf"<{element}\b[^>]*>", self.xml)
        return found.group(0) if found else ""

    @member
    def Name(self) -> object:
        self.admit_read()
        return self.name

    @setter("Name")
    def _set_name(self, value: object) -> None:
        raise VBAUnsupportedError("Renaming a chart sheet is not implemented")

    @member
    def Parent(self) -> object:
        return self.book

    @member
    def Index(self) -> object:
        self.admit_read()
        return VBAInt(self.book.tabs().index(self) + 1, "Long")

    @member
    def CodeName(self) -> object:
        self.admit_read()
        return self.code_name

    @member
    def Type(self) -> object:
        """The chart's kind as the old gallery constants number it, as a series' Type is."""
        self.admit_read()
        return VBAInt(SERIES_KINDS[CHART_TYPES[shown_type(self.state)]], "Long")

    @member
    def Visible(self) -> object:
        self.admit_read()
        return VBAInt(self.visible, "Long")

    @setter("Visible")
    def _set_visible(self, value: object) -> None:
        raise VBAUnsupportedError("Hiding or showing a chart sheet is not implemented")

    @member
    def Next(self) -> object:
        self.admit_read()
        return self._tab(1)

    @member
    def Previous(self) -> object:
        self.admit_read()
        return self._tab(-1)

    @member
    def Tab(self) -> object:
        self.admit_read()
        return _ChartSheetTab(self)

    @member
    def ProtectContents(self) -> object:
        self.admit_read()
        return bool(re.search(r'\bcontent="1"', self._element("sheetProtection")))

    @method
    def Activate(self) -> object:
        self.admit_read()
        self.book.active_chart_sheet = self
        self.book.view.first_sheet = min(self.book.view.first_sheet, self.book.tabs().index(self))
        self.book.selected_sheets = []
        self.book.application.activate_book(self.book)
        return EMPTY

    @method
    def Select(self, Replace: object = MISSING) -> object:  # noqa: N803 - native named argument
        """A hidden chart sheet cannot be selected, though it can be activated (tests/fixtures/chart_sheets/hidden.json)."""
        if self.visible != -1:
            raise error(1004, "Select method of Chart class failed")
        if Replace is not MISSING:
            raise VBAUnsupportedError("Selecting a chart sheet with Replace is not implemented")
        return self.Activate()

    @method
    def Delete(self) -> object:
        raise VBAUnsupportedError("Deleting a chart sheet is not implemented")

    @method
    def Copy(self, Before: object = MISSING, After: object = MISSING) -> object:  # noqa: N803 - native named arguments
        raise VBAUnsupportedError("Copying a chart sheet is not implemented")

    @method
    def Move(self, Before: object = MISSING, After: object = MISSING) -> object:  # noqa: N803 - native named arguments
        raise VBAUnsupportedError("Moving a chart sheet is not implemented")

    @method
    def Protect(self, Password: object = MISSING, DrawingObjects: object = MISSING, Contents: object = MISSING,  # noqa: N803 - native named arguments
                Scenarios: object = MISSING, UserInterfaceOnly: object = MISSING) -> object:
        raise VBAUnsupportedError("Protecting a chart sheet is not implemented")

    @member
    def PageSetup(self) -> object:
        raise VBAUnsupportedError("PageSetup of a chart sheet is not implemented")

    def saved_xml(self) -> str:
        """The part, its view selected as the tab is: the one thing a macro can change in it."""
        active = self.book.active_chart_sheet is self
        view = re.search(r"<sheetView\b[^>]*>", self.xml)
        if view is None or (' tabSelected="1"' in view.group(0)) == active:
            return self.xml
        tag = view.group(0).replace(' tabSelected="1"', "")
        if active:
            tag = tag.replace("<sheetView", '<sheetView tabSelected="1"', 1)
        return self.xml[:view.start()] + tag + self.xml[view.end():]


class _ChartSheetTab(ExcelObject):
    """A chart sheet's tab as the file has it; False is what Excel reads for a tab with no colour."""

    vba_type_name = "Tab"

    def __init__(self, sheet: ChartSheet) -> None:
        self.sheet = sheet

    def _colour(self) -> int | None:
        found = re.search(r'<tabColor\b[^>]*\brgb="[0-9A-Fa-f]{2}([0-9A-Fa-f]{6})"', self.sheet.xml)
        if found is None:
            if "<tabColor" in self.sheet.xml:
                raise VBAUnsupportedError("A chart sheet's tab colour other than an RGB one is not implemented")
            return None
        red, green, blue = (int(found.group(1)[at:at + 2], 16) for at in (0, 2, 4))
        return red + green * 256 + blue * 65536

    @member
    def Color(self) -> object:
        colour = self._colour()
        return False if colour is None else VBAInt(colour, "Long")

    @setter("Color")
    def _set_color(self, value: object) -> None:
        raise VBAUnsupportedError("Colouring a chart sheet's tab is not implemented")


class ChartSheets(VBACollection, ItemDefault, ExcelObject):
    """Workbook.Charts where the workbook has chart sheets."""

    vba_type_name = "Sheets"

    def __init__(self, book: Workbook) -> None:
        self.book = book

    def vba_items(self) -> list[object]:
        return list(self.book.chart_sheets)

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            for sheet in self.book.chart_sheets:
                if sheet.name.casefold() == index.casefold():
                    return sheet
            raise error(ERR_SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range")
        return super().vba_lookup(index, items)

    @method
    def Add(self, Before: object = MISSING, After: object = MISSING, Count: object = MISSING,  # noqa: N803 - native named arguments
            Type: object = MISSING) -> object:
        raise VBAUnsupportedError("Charts.Add is not implemented")
