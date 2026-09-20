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

__all__ = [
    "Application",
    "Cell",
    "ExcelApplication",
    "Range",
    "SheetView",
    "Workbook",
    "Worksheet",
]


class SheetView:
    """A worksheet, from Python rather than from VBA."""

    def __init__(self, sheet: Worksheet) -> None:
        self.sheet = sheet

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

    def formula(self, reference: str) -> str:
        target = self.sheet.vba_get("Range", [reference])
        assert isinstance(target, Range)
        return to_text(target.vba_get("Formula"))

    def rows(self) -> list[list[object]]:
        """Everything on the sheet, row by row, as Python values."""
        bounds = self.sheet.used_bounds()
        if bounds is None:
            return []
        top, left, bottom, right = bounds
        out: list[list[object]] = []
        for row in range(top, bottom + 1):
            line: list[object] = []
            for column in range(left, right + 1):
                cell = self.sheet.cell(row, column)
                line.append(None if cell is None else _plain(cell.value))
            out.append(line)
        return out

    def describe(self) -> str:
        return self.sheet.describe()

    def __repr__(self) -> str:
        return f"<SheetView {self.sheet.name!r}>"


class ExcelApplication:
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

    def sheet(self, which: int | str = 1) -> SheetView:
        """One sheet by position (from 1) or by name."""
        book = self.workbook
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

    def save(self, path: str | Path | None = None) -> Path:
        """Write the workbook out, keeping every part this does not model."""
        from pyopenvba.apps.excel._io import save_workbook

        book = self.workbook
        target = Path(path) if path is not None else Path(book.path) / book.name
        save_workbook(book, target)
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
