"""An Excel instance in memory, with a VBA interpreter attached.

    >>> from pyopenvba.apps.excel import ExcelApplication
    >>> app = ExcelApplication()
    >>> book = app.add_workbook()
    >>> app.add_module('''
    ... Sub Fill()
    ...     Dim r As Long
    ...     For r = 1 To 3
    ...         Cells(r, 1).Value = r * 10
    ...     Next r
    ...     Range("B1").Value = "total"
    ...     Range("B2").Value = Application.WorksheetFunction.Sum(Range("A1:A3"))
    ... End Sub
    ... ''', name="Module1")
    >>> _ = app.run("Fill")
    >>> app.sheet(1).value("A1"), app.sheet(1).value("B2")
    (10, 300.0)

State comes from nothing, as above, or from a real workbook through
:meth:`ExcelApplication.open`, and goes back out through
:meth:`ExcelApplication.save`.  :meth:`ExcelApplication.describe` prints
what is in there.
"""

from __future__ import annotations

import datetime as _dt
from pathlib import Path
from typing import Any

from pyopenvba.apps.excel._bridge import ExcelBridge
from pyopenvba.apps.excel._model import (
    Application,
    Cell,
    Range,
    Workbook,
    Worksheet,
)
from pyopenvba.interpreter._runtime import Interpreter, ModuleRuntime
from pyopenvba.interpreter._values import EMPTY, MISSING, to_text
from pyopenvba.shapes import Shape
from pyopenvba.apps.excel._named_api import NamedRange, NamedRangeAPI
from pyopenvba.apps.excel._model import Names

__all__ = [
    "Application",
    "Cell",
    "ExcelApplication",
    "Range",
    "SheetView",
    "Workbook",
    "Worksheet",
    "NamedRange",
]


class SheetView(NamedRangeAPI):
    """A worksheet, from Python rather than from VBA."""

    def __init__(self, sheet: Worksheet) -> None:
        self.sheet = sheet

    @property
    def _named_collection(self) -> Names:
        return Names(self.sheet.book, self.sheet)

    @property
    def name(self) -> str:
        return self.sheet.name

    def value(self, reference: str) -> object:
        """What one cell or block holds, as a Python value."""
        target = self.sheet.vba_get("Range", [reference])
        assert isinstance(target, Range)
        return _plain(target.vba_get("Value"))

    def set_value(self, reference: str, value: object) -> None:
        target = self.sheet.vba_get("Range", [reference])
        assert isinstance(target, Range)
        target.vba_set("Value", value)

    def copy_range(self, reference: str, destination: SheetView, target: str, *, name_conflict: str = "reuse") -> None:
        """Copy cells; cross-book name conflicts can reuse, rename, or error."""
        source = self.sheet.vba_get("Range", [reference])
        assert isinstance(source, Range)
        source.copy_to(destination.sheet.vba_get("Range", [target]), name_conflict=name_conflict)

    def copy(self, *, before: SheetView | None = None, after: SheetView | None = None) -> SheetView:
        """Copy this sheet; no destination creates and activates a new workbook."""
        self.sheet.Copy(before.sheet if before else MISSING, after.sheet if after else MISSING)
        book = self.sheet.book.application.active_book
        assert book is not None and book.active_sheet is not None
        return SheetView(book.active_sheet)

    def move(self, *, before: SheetView | None = None, after: SheetView | None = None) -> SheetView:
        """Move this sheet; omit both anchors to move into a new workbook."""
        self.sheet.Move(before.sheet if before else MISSING, after.sheet if after else MISSING)
        book = self.sheet.book.application.active_book
        assert book is not None and book.active_sheet is not None
        return self if book.active_sheet is self.sheet else SheetView(book.active_sheet)

    def formula(self, reference: str) -> str:
        target = self.sheet.vba_get("Range", [reference])
        assert isinstance(target, Range)
        return to_text(target.vba_get("Formula"))

    def shapes(self) -> list[Shape]:
        """Shape snapshots in drawing order, including control details.

        Use ``update_shape`` to edit. Mutating a snapshot does not change
        the workbook, including its nested controls and group members.
        """
        from pyopenvba.apps.excel._shape_api import snapshot
        from pyopenvba.apps.excel._controls import refresh

        for item in self.sheet.shapes_:
            refresh(self.sheet, item)
        return [snapshot(item) for item in self.sheet.shapes_]

    def shape(self, name: str) -> Shape:
        """A snapshot by case-insensitive name; raise KeyError if absent."""
        from pyopenvba.apps.excel._shape_api import lookup, snapshot
        from pyopenvba.apps.excel._controls import refresh

        item = lookup(self.sheet, name)
        refresh(self.sheet, item)
        return snapshot(item)

    def add_shape(
        self, shape_type: int = 1, *, name: str | None = None,
        left: float = 0, top: float = 0, width: float = 100, height: float = 50,
        text: str = "", macro: str = "",
    ) -> Shape:
        """Add a measured msoAutoShapeType (1 is a rectangle), in points."""
        from pyopenvba.apps.excel._shape_api import add

        return add(self.sheet, "shape", shape_type, name=name, left=left, top=top,
                   width=width, height=height, text=text, macro=macro)

    def add_textbox(
        self, *, name: str | None = None, left: float = 0, top: float = 0,
        width: float = 100, height: float = 50, text: str = "", macro: str = "",
    ) -> Shape:
        """Add a horizontal text box, with coordinates and size in points."""
        from pyopenvba.apps.excel._shape_api import add

        return add(self.sheet, "textBox", 1, name=name, left=left, top=top,
                   width=width, height=height, text=text, macro=macro)

    def add_button(
        self, *, name: str | None = None, left: float = 0, top: float = 0,
        width: float = 100, height: float = 30, text: str = "", macro: str = "",
    ) -> Shape:
        """Add a Forms button and its drawing, worksheet, properties and VML records."""
        from pyopenvba.apps.excel._shape_api import add

        return add(self.sheet, "formControl", 0, name=name, left=left, top=top,
                   width=width, height=height, text=text, macro=macro)

    def add_form_control(
        self, control_type: int, *, name: str | None = None,
        left: float = 0, top: float = 0, width: float = 100, height: float = 30,
        text: str = "", macro: str = "",
    ) -> Shape:
        """Add a Forms control using Excel's xlFormControl number.

        Supported types: button (0), checkbox (1), dropdown (2), group
        box (4), label (5), list box (6), option button (7), scroll bar
        (8), spinner (9). Coordinates and dimensions are in points.
        """
        from pyopenvba.apps.excel._shape_api import add

        return add(self.sheet, "formControl", control_type, name=name, left=left, top=top,
                   width=width, height=height, text=text, macro=macro)

    def update_shape(
        self, name: str, *, new_name: str | None = None,
        left: float | None = None, top: float | None = None,
        width: float | None = None, height: float | None = None,
        text: str | None = None, macro: str | None = None,
    ) -> Shape:
        """Edit a shape. None leaves a field alone; an empty string clears text or macro.

        Positions and sizes are finite, nonnegative points. Names must
        be nonempty and unique on the sheet. Validation precedes changes.
        """
        from pyopenvba.apps.excel._shape_api import update

        return update(self.sheet, name, new_name=new_name, left=left, top=top,
                      width=width, height=height, text=text, macro=macro)

    def update_control(
        self, name: str, *, linked_cell: str | None = None, list_range: str | None = None,
    ) -> Shape:
        """Edit saved A1 or named control bindings; an empty string disconnects it.

        None preserves the current setting. References must address this
        workbook. This changes the binding without writing a cell value.
        Supported control values and cell edits subsequently synchronize through
        ``set_control_value`` or VBA's ``ControlFormat.Value``.
        """
        from pyopenvba.apps.excel._shape_api import update_control

        return update_control(self.sheet, name, linked_cell=linked_cell, list_range=list_range)

    def set_control_value(self, name: str, value: int) -> Shape:
        """Set a checkbox/radio state, list index, or spinner/scroll-bar value.

        Checkboxes accept off (-4146 or 0), on (1), or mixed (2).
        Radios accept off or on; selecting one clears its group peers
        and writes the selected one-based index to the group's linked cell.
        Dropdowns/list boxes accept zero or a one-based index within the
        A1 source range or inline list. Changed single-selection values
        update the linked cell. For a multi/extended list, this replaces
        its selected indexes with one index (or clears them with zero)
        without writing the linked cell.
        """
        from pyopenvba.apps.excel._controls import set_value
        from pyopenvba.apps.excel._shape_api import lookup, snapshot

        item = lookup(self.sheet, name)
        set_value(self.sheet, item, value)
        return snapshot(item)

    def remove_shape(self, name: str) -> None:
        """Remove a shape by name, including a form control's supporting records."""
        from pyopenvba.apps.excel._shape_api import remove

        remove(self.sheet, name)

    def control_items(self, name: str) -> list[str]:
        """Return detached list items from an inline list or its A1 source."""
        from pyopenvba.apps.excel._controls import items
        from pyopenvba.apps.excel._shape_api import lookup

        return items(self.sheet, lookup(self.sheet, name))

    def add_control_item(self, name: str, text: str, index: int = 0) -> None:
        """Insert at a one-based index; zero or beyond the end appends.

        A range-backed list is disconnected and replaced by this one item,
        matching Excel. Its source cells are not edited.
        """
        self._edit_control_items(name, "add", index, text)

    def set_control_items(self, name: str, entries: list[str]) -> None:
        """Replace a list with validated strings, disconnecting an A1 source.

        Resets selection. Converting a range-backed control can write zero
        to its linked cell; source cells remain untouched. An empty list
        clears the control. Validation occurs before any change.
        """
        from pyopenvba.apps.excel._controls import replace_items
        from pyopenvba.apps.excel._shape_api import lookup

        replace_items(self.sheet, lookup(self.sheet, name), entries)

    def update_control_item(self, name: str, index: int, text: str) -> None:
        """Replace at a positive index, or append beyond the end, as Excel does.

        A range-backed list is disconnected and replaced by this one item.
        """
        self._edit_control_items(name, "set", index, text)

    def remove_control_item(self, name: str, index: int, count: int = 1) -> None:
        """Remove inline items without writing the linked cell; reject range sources."""
        self._edit_control_items(name, "remove", index, count=count)

    def clear_control_items(self, name: str) -> None:
        """Clear items and selection; disconnect a range source and reset its linked index."""
        self._edit_control_items(name, "clear")

    def _edit_control_items(self, name: str, operation: str, index: int = 0,
                            text: str = "", count: int = 1) -> None:
        from pyopenvba.apps.excel._controls import edit_items
        from pyopenvba.apps.excel._shape_api import lookup

        edit_items(self.sheet, lookup(self.sheet, name), operation, index, text, count)

    def set_control_selection_mode(self, name: str, mode: int) -> None:
        """Set list-box mode: -4142 (or 1) single, -4154 (or 2) multi, 3 extended."""
        from pyopenvba.apps.excel._controls import set_mode
        from pyopenvba.apps.excel._shape_api import lookup

        set_mode(self.sheet, lookup(self.sheet, name), mode)

    def set_control_selection(self, name: str, indices: list[int]) -> None:
        """Replace a multi/extended list-box selection with one-based indexes."""
        from pyopenvba.apps.excel._controls import set_selection
        from pyopenvba.apps.excel._shape_api import lookup

        set_selection(self.sheet, lookup(self.sheet, name), indices)

    def rows(self) -> list[list[object]]:
        """Everything on the sheet, row by row, as Python values."""
        from pyopenvba.apps.excel._calc import as_python

        bounds = self.sheet.used_bounds()
        if bounds is None:
            return []
        top, left, bottom, right = bounds
        out: list[list[object]] = []
        for row in range(top, bottom + 1):
            line: list[object] = []
            for column in range(left, right + 1):
                cell = self.sheet.cell(row, column)
                if cell is None:
                    line.append(None)
                    continue
                # A number is a Date or a Currency only through its cell's format, and only where one holds it.
                number = isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool)
                line.append(_plain(as_python(cell.value, cell) if number else cell.value))
            out.append(line)
        return out

    def describe(self) -> str:
        return self.sheet.describe()

    def __repr__(self) -> str:
        return f"<SheetView {self.sheet.name!r}>"


class ExcelApplication(NamedRangeAPI):
    """Excel, in memory: workbooks, sheets, cells, and the macros that move them."""

    def __init__(self) -> None:
        self.application = Application()
        self.interpreter = Interpreter(ExcelBridge(self.application))
        self.application.interpreter = self.interpreter

    # --- state in -------------------------------------------------------------------

    @classmethod
    def open(cls, path: str | Path, *, with_vba: bool = True) -> ExcelApplication:
        """Load a workbook, its sheets, its names, its queries and its macros."""
        from pyopenvba.apps.excel._io import load_workbook

        app = cls()
        book = load_workbook(app.application, Path(path))
        app.application.workbooks_.books.append(book)
        app.application.activate_book(book)
        if with_vba:
            app.load_vba(path)
        return app

    def load_vba(self, path: str | Path) -> list[str]:
        """Add every module of the file's VBA project to the interpreter.

        A workbook with no project is ordinary rather than an error, so
        this comes back with an empty list rather than raising.
        """
        from pyopenvba.excel import ExcelFile
        from pyopenvba.exceptions import PyOpenVBAError
        from pyopenvba.vba import VBAModuleKind

        target = Path(path)
        if target.suffix.lower() not in (".xlsm", ".xlsb", ".xlam", ".xls"):
            return []
        names: list[str] = []
        try:
            with ExcelFile(target) as host:
                project = host.vba_project()
                for module in project.modules:
                    kind = "standard" if module.kind is VBAModuleKind.standard else "class"
                    self.interpreter.add_module(module.source, name=module.name, kind=kind)
                    names.append(module.name)
        except PyOpenVBAError:
            return names
        return names

    def add_workbook(self) -> Workbook:
        """A new empty workbook with one sheet, as Excel's New does."""
        book = self.application.workbooks_.vba_get("Add")
        assert isinstance(book, Workbook)
        return book

    def open_workbook(self, path: str | Path) -> Workbook:
        """Open another workbook without importing its VBA modules."""
        book = self.application.workbooks_.Open(Filename=str(path))
        assert isinstance(book, Workbook)
        return book

    def workbooks(self) -> list[Workbook]:
        return list(self.application.workbooks_.books)

    def activate_workbook(self, book: Workbook) -> None:
        if book not in self.application.workbooks_.books:
            raise ValueError("Workbook is not open in this application")
        book.Activate()

    def add_module(self, source: str, *, name: str = "", kind: str = "standard") -> ModuleRuntime:
        """Parse VBA and add it to the project this instance runs."""
        return self.interpreter.add_module(source, name=name, kind=kind)

    # --- running --------------------------------------------------------------------

    def run(self, macro: str, *args: object) -> object:
        """Run a macro by name, as Application.Run would."""
        return _plain(self.interpreter.run(macro, list(args)))

    def evaluate(self, expression: str) -> object:
        """Evaluate one VBA expression against the current state."""
        from pyopenvba.interpreter._parse import parse_expression

        parsed = parse_expression(expression)
        self.interpreter.initialise()
        frame = _bare_frame(self.interpreter)
        return _plain(self.interpreter.evaluate(parsed, frame))

    def refresh_query(self, name: str) -> list[list[object]]:
        """Evaluate one query and give back its rows, headers first.

        The rows also land on the sheet the query loads to, when it has
        one, which is what the workbook keeps and what a save writes.
        """
        from pyopenvba.apps.excel._refresh import refresh

        table = refresh(self.workbook, name)
        return [list(table.columns), *[list(row) for row in table.rows]]

    def refresh_all(self) -> None:
        """Evaluate every query in the workbook, as RefreshAll does."""
        self.workbook.vba_get("RefreshAll")

    def freeze_clock(self, when: _dt.datetime) -> None:
        """Pin Now, Date and Time, so two runs can be compared."""
        self.interpreter.freeze(when)

    def answer_dialogs(self, *answers: object) -> None:
        """Queue what the next MsgBox and InputBox calls come back with."""
        self.interpreter.answers.extend(answers)

    @property
    def console(self) -> list[str]:
        """Everything Debug.Print wrote."""
        return self.interpreter.console

    @property
    def dialogs(self) -> list[tuple[str, str, str, int]]:
        """Every MsgBox and InputBox the run put up."""
        return self.interpreter.dialogs

    # --- looking at the state ---------------------------------------------------------

    @property
    def workbook(self) -> Workbook:
        book = self.application.active_book
        if book is None:
            raise LookupError("no workbook is open")
        return book

    @property
    def _named_collection(self) -> Names:
        return self.workbook.names_

    def sheet(self, which: int | str = 1, *, workbook: Workbook | None = None) -> SheetView:
        """One sheet by position (from 1) or by name."""
        book = workbook or self.workbook
        if book not in self.application.workbooks_.books:
            raise ValueError("Workbook is not open in this application")
        if isinstance(which, int):
            if not 1 <= which <= len(book.sheets_):
                raise LookupError(f"this workbook has {len(book.sheets_)} sheets")
            return SheetView(book.sheets_[which - 1])
        return SheetView(book.sheet_named(which))

    def sheets(self) -> list[SheetView]:
        return [SheetView(sheet) for sheet in self.workbook.sheets_]

    def describe(self) -> str:
        """The whole instance as a developer-readable block of text."""
        lines = [self.application.describe()]
        if self.interpreter.modules:
            names = ", ".join(runtime.name for runtime in self.interpreter.modules.values())
            lines.append(f"  VBA: {names}")
        if self.console:
            lines.append("  Debug.Print:")
            lines.extend(f"    {line}" for line in self.console)
        if self.dialogs:
            lines.append("  Dialogs:")
            for kind, prompt, title, _style in self.dialogs:
                lines.append(f"    {kind}: {prompt}" + (f" [{title}]" if title else ""))
        return "\n".join(lines)

    # --- state out ---------------------------------------------------------------------

    def save(self, path: str | Path | None = None, *, workbook: Workbook | None = None) -> Path:
        """Write the workbook out, keeping every part this does not model."""
        from pyopenvba.apps.excel._io import save_workbook

        book = workbook or self.workbook
        if book not in self.application.workbooks_.books:
            raise ValueError("Workbook is not open in this application")
        target = Path(path) if path is not None else Path(book.path) / book.name
        save_workbook(book, target)
        book.path, book.name = str(target.resolve().parent), target.name
        book.saved = True
        return target

    def __repr__(self) -> str:
        return f"<ExcelApplication {len(self.application.workbooks_.books)} workbook(s)>"


def _bare_frame(interpreter: Interpreter) -> Any:
    from pyopenvba.interpreter import _ast as A
    from pyopenvba.interpreter._runtime import Frame, ModuleRuntime as _ModuleRuntime

    if interpreter.modules:
        module = next(iter(interpreter.modules.values()))
    else:
        module = _ModuleRuntime(A.Module(name="<immediate>"), interpreter)
    return Frame(procedure=A.Procedure(name="<immediate>"), module=module)


def _plain(value: object) -> Any:
    """A VBA value as the plain Python one a caller expects."""
    from pyopenvba.interpreter._values import VBAArray, VBADate, VBAInt, VBASingle

    if value is EMPTY or value is MISSING:
        return None
    if isinstance(value, VBAInt):
        return int(value)
    if isinstance(value, VBASingle):
        return float(value)
    if isinstance(value, VBADate):
        return value.to_datetime()
    if isinstance(value, VBAArray):
        if value.dimensions == 2:
            rows = value.bounds[0]
            columns = value.bounds[1]
            return [
                [_plain(value.get([row, column])) for column in range(columns[0], columns[1] + 1)]
                for row in range(rows[0], rows[1] + 1)
            ]
        return [_plain(item) for item in value.elements()]
    return value
