"""A chart sheet read from a file: a tab of its own holding one chart, which a macro reaches through Sheets and Charts and
Worksheets passes over (tests/fixtures/chart_sheets/chart_sheets.json). The chart is the file's, read as an embedded
chart's is; the part stays as it came, so a save keeps everything the model does not read. A workbook with one refuses
adding, moving, copying, deleting or renaming tabs for now: where they land among chart sheets is not measured."""
from __future__ import annotations

import posixpath
import re
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._charts import CHART_TYPES, SERIES_KINDS, Chart, shown_type
from pyopenvba.apps.excel._charts_file import REL_CHART, read_chart, rels_part, xml_tree
from pyopenvba.apps.excel._creator import ItemDefault
from pyopenvba.apps.excel._model import ExcelObject, Worksheet
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import EMPTY, ERR_SUBSCRIPT_OUT_OF_RANGE, MISSING, NOTHING, VBAInt, error
from pyopenvba.shapes._values import Shape

if TYPE_CHECKING:
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
