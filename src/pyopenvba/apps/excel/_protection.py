"""Sheet protection: Protect, Unprotect, and the edits a protected sheet refuses a macro.

Measured from VBA (scripts/measure_protection.py, tests/fixtures/protection.json).

Protect sets Contents, DrawingObjects and Scenarios, True unless given,
and the Allow options, False unless given. A second Protect on a
protected sheet does nothing when it asks for what the sheet already has,
UserInterfaceOnly kept unless given, and checks nothing then, not even a
password that differs; otherwise it needs the sheet's password and takes
the new one. A password is matched case and all, and a sheet with none
opens to any. Unprotect or a changing Protect with no password on a sheet
that has one asks for it in a dialog, which a macro here cannot answer.

While a sheet is protected with Contents, and not for the user interface
alone, a macro's edits meet it:

- A value or formula reaches every unlocked cell it is written to, then
  error 1004 says the sheet is protected if any locked cell was among
  them; clearing contents goes the same way, and Clear clears contents
  only.
- Formatting a cell, locked or not, needs AllowFormattingCells, a column's
  width or hiding AllowFormattingColumns, a row's AllowFormattingRows;
  Locked, FormulaHidden and MergeCells stay refused. The error names the
  property, except a colour or pattern's, which is Excel's generic one.
- Inserting whole rows or columns needs the Allow option; deleting them
  needs it and no locked cell in them; inserting or deleting cells is
  refused.
- Merging, copying or cutting onto locked cells, filling, AutoFilter,
  RemoveDuplicates and sorting a range fail; Replace answers True and
  changes nothing; the Sort object sorts anyway.

Protect empties the clipboard, as Excel's does.
"""

from __future__ import annotations

from collections.abc import Generator
from contextlib import contextmanager
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area
from pyopenvba.apps.excel._model import ExcelObject
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import member
from pyopenvba.interpreter._values import MISSING, error, to_bool, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Worksheet

#: Error 1004's text for a change a protected sheet refuses.
PROTECTED = ("The cell or chart you're trying to change is on a protected sheet. To make a change, unprotect the "
             "sheet. You might be requested to enter a password.")
WRONG_PASSWORD = ("The password you supplied is not correct. Verify that the CAPS LOCK key is off and be sure to use "
                  "the correct capitalization.")
APPLICATION_DEFINED = "Application-defined or object-defined error"
FILTERING = ("You cannot use this command on a protected sheet. To use this command, you must first unprotect the "
             "sheet (Review tab, Protect group, Unprotect Sheet button). You may be prompted for a password.")

#: Protect's Allow options, in its order.
ALLOWS = ("AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows", "AllowInsertingColumns",
          "AllowInsertingRows", "AllowInsertingHyperlinks", "AllowDeletingColumns", "AllowDeletingRows",
          "AllowSorting", "AllowFiltering", "AllowUsingPivotTables")
#: Format properties whose refusal is Excel's generic error rather than one naming the property.
_GENERIC = frozenset({"Color", "ColorIndex", "ThemeColor", "TintAndShade", "Pattern", "PatternColor",
                      "PatternColorIndex", "PatternThemeColor", "PatternTintAndShade", "Gradient", "InvertIfNegative"})


@dataclass(frozen=True, slots=True)
class SheetProtection:
    """What Protect set on a sheet."""

    password: str = ""
    contents: bool = True
    drawing_objects: bool = True
    scenarios: bool = True
    user_interface_only: bool = False
    allows: frozenset[str] = field(default_factory=lambda: frozenset[str]())

    def options(self) -> tuple[bool, bool, bool, bool, frozenset[str]]:
        return self.contents, self.drawing_objects, self.scenarios, self.user_interface_only, self.allows

    def opens(self, password: str) -> bool:
        """Whether a password opens the sheet: its own, matched case and all, or any when it has none."""
        return not self.password or password == self.password


class Protection(ExcelObject):
    """Worksheet.Protection: the Allow options the sheet's last Protect gave, which Unprotect leaves as they are."""

    vba_type_name = "Protection"

    def __init__(self, sheet: Worksheet) -> None:
        self.sheet = sheet

    def _allows(self, name: str) -> bool:
        # What the last Protect gave, kept through Unprotect.
        return name in self.sheet.protection_allows

    @member
    def AllowFormattingCells(self) -> object:
        return self._allows("AllowFormattingCells")

    @member
    def AllowFormattingColumns(self) -> object:
        return self._allows("AllowFormattingColumns")

    @member
    def AllowFormattingRows(self) -> object:
        return self._allows("AllowFormattingRows")

    @member
    def AllowInsertingColumns(self) -> object:
        return self._allows("AllowInsertingColumns")

    @member
    def AllowInsertingRows(self) -> object:
        return self._allows("AllowInsertingRows")

    @member
    def AllowInsertingHyperlinks(self) -> object:
        return self._allows("AllowInsertingHyperlinks")

    @member
    def AllowDeletingColumns(self) -> object:
        return self._allows("AllowDeletingColumns")

    @member
    def AllowDeletingRows(self) -> object:
        return self._allows("AllowDeletingRows")

    @member
    def AllowSorting(self) -> object:
        return self._allows("AllowSorting")

    @member
    def AllowFiltering(self) -> object:
        return self._allows("AllowFiltering")

    @member
    def AllowUsingPivotTables(self) -> object:
        return self._allows("AllowUsingPivotTables")


def protect(sheet: Worksheet, password: object, options: dict[str, object]) -> None:
    """Worksheet.Protect, with ``options`` Protect's other arguments by name, MISSING where left out."""
    current = sheet.protection

    def flag(name: str, default: bool) -> bool:
        given = options.get(name, MISSING)
        return default if given is MISSING else to_bool(given)

    supplied = "" if password is MISSING else to_text(password)
    wanted = SheetProtection(
        password=supplied, contents=flag("Contents", True), drawing_objects=flag("DrawingObjects", True),
        scenarios=flag("Scenarios", True),
        user_interface_only=flag("UserInterfaceOnly", current.user_interface_only if current else False),
        allows=frozenset(name for name in ALLOWS if flag(name, False)))
    if current is not None:
        if wanted.options() == current.options():
            return
        if current.password and password is MISSING:
            raise VBAUnsupportedError("Protect that changes a sheet protected with a password, given none, asks for "
                                      "the password in a dialog; that is not implemented")
        if not current.opens(supplied):
            raise error(1004, WRONG_PASSWORD)
    sheet.protection = wanted
    sheet.protection_allows = wanted.allows
    sheet.book.application.clipboard = None


def unprotect(sheet: Worksheet, password: object) -> None:
    current = sheet.protection
    if current is None:
        return
    if current.password and password is MISSING:
        raise VBAUnsupportedError("Unprotect of a sheet protected with a password, given none, asks for the "
                                  "password in a dialog; that is not implemented")
    if not current.opens("" if password is MISSING else to_text(password)):
        raise error(1004, WRONG_PASSWORD)
    sheet.protection = None


def enforced(sheet: Worksheet) -> SheetProtection | None:
    """The protection a macro's edit to cells meets: none unless Contents is on and not for the user alone."""
    found = sheet.protection
    if found is None or not found.contents or found.user_interface_only:
        return None
    return found


def drawing_enforced(sheet: Worksheet) -> bool:
    """Whether a macro meets the protection of the sheet's shapes: DrawingObjects on, and not for the user alone."""
    found = sheet.protection
    return found is not None and found.drawing_objects and not found.user_interface_only


def guard_drawing(sheet: Worksheet, what: str) -> None:
    """Report a change to a drawing-protected sheet's shapes that was not measured."""
    if drawing_enforced(sheet):
        raise VBAUnsupportedError(f"{what} on a sheet whose drawing objects are protected is not implemented")


def refuse(sheet: Worksheet, message: str, allow: str = "") -> None:
    """Error 1004 with ``message`` when the sheet refuses the edit, unless the Allow option named lets it through."""
    found = enforced(sheet)
    if found is not None and not (allow and allow in found.allows):
        raise error(1004, message)


#: Range properties a protected sheet refuses, and the Allow option that lets each through; none for those it
#: always refuses.
_RANGE_FORMATS: dict[str, str] = {
    **dict.fromkeys(("NumberFormat", "HorizontalAlignment", "VerticalAlignment", "WrapText", "ShrinkToFit",
                     "IndentLevel", "AddIndent", "Orientation", "ReadingOrder"), "AllowFormattingCells"),
    **dict.fromkeys(("Locked", "FormulaHidden", "MergeCells"), ""),
    **dict.fromkeys(("ColumnWidth", "UseStandardWidth"), "AllowFormattingColumns"),
    **dict.fromkeys(("RowHeight", "UseStandardHeight"), "AllowFormattingRows"),
}
#: Range properties whose writes the write gate meets, or that change no cell.
_RANGE_FREE = frozenset({"Value", "Value2", "Formula", "FormulaR1C1", "Name"})


def check_range_set(target: Range, member: str) -> None:
    """Refuse setting a Range property where a protected sheet refuses it."""
    sheet = target.sheet
    if enforced(sheet) is None or member in _RANGE_FREE:
        return
    if member == "Hidden":
        rows = target.whole == "rows" or (not target.whole and all(
            area.left == 1 and area.right == MAX_COLUMNS for area in target.areas))
        refuse(sheet, refused_format("Range", member), "AllowFormattingRows" if rows else "AllowFormattingColumns")
        return
    if member not in _RANGE_FORMATS:
        raise VBAUnsupportedError(f"setting Range.{member} on a protected sheet is not implemented")
    refuse(sheet, refused_format("Range", member), _RANGE_FORMATS[member])


def whole_lines(target: Range) -> str:
    """"rows" or "columns" when a range is whole rows or whole columns, else ""."""
    if target.whole:
        return target.whole
    if all(area.left == 1 and area.right == MAX_COLUMNS for area in target.areas):
        return "rows"
    if all(area.top == 1 and area.bottom == MAX_ROWS for area in target.areas):
        return "columns"
    return ""


def check_insert(target: Range) -> None:
    """Inserting whole rows or columns needs its Allow option; inserting cells is refused."""
    sheet = target.sheet
    if enforced(sheet) is None:
        return
    message = "Insert method of Range class failed"
    lines = whole_lines(target)
    if not lines:
        raise error(1004, message)
    refuse(sheet, message, "AllowInsertingRows" if lines == "rows" else "AllowInsertingColumns")


def check_delete(target: Range) -> None:
    """Deleting whole rows or columns needs its Allow option and no locked cell among them; deleting cells is refused."""
    sheet = target.sheet
    if enforced(sheet) is None:
        return
    message = "Delete method of Range class failed"
    lines = whole_lines(target)
    if not lines:
        raise error(1004, message)
    refuse(sheet, message, "AllowDeletingRows" if lines == "rows" else "AllowDeletingColumns")
    if any_locked(sheet, target.areas):
        raise error(1004, PROTECTED)


def check_sort(target: Range) -> None:
    """Range.Sort needs AllowSorting, and then unlocked cells; the Sort object is not stopped at all."""
    sheet = target.sheet
    found = enforced(sheet)
    if found is None:
        return
    if "AllowSorting" not in found.allows:
        raise error(1004, "Sort method of Range class failed")
    whole(sheet, target.areas, "Sort")


def check_format_set(sheet: Worksheet, class_name: str, member: str) -> None:
    """Refuse setting a Font, Interior, Border or Borders property of a protected sheet's cells."""
    refuse(sheet, refused_format(class_name, member), "AllowFormattingCells")


def refused_format(class_name: str, member: str) -> str:
    """What Excel says when a protected sheet refuses a format property."""
    if member in _GENERIC:
        return APPLICATION_DEFINED
    return f"Unable to set the {member} property of the {class_name} class"


def locked(sheet: Worksheet, row: int, column: int) -> bool:
    return sheet.style_at(row, column).protection.locked


def _lock_states(sheet: Worksheet, areas: list[Area]) -> set[bool]:
    """Whether the areas hold locked positions, unlocked ones or both, without walking positions nobody touched."""
    from pyopenvba.apps.excel import _row_formats

    found: set[bool] = set()
    for area in areas:
        bottom, right = min(area.bottom, MAX_ROWS), min(area.right, MAX_COLUMNS)
        if (bottom - area.top + 1) * (right - area.left + 1) <= 4096:
            found.update(locked(sheet, row, column) for row in range(area.top, bottom + 1)
                         for column in range(area.left, right + 1))
        else:
            found.update(style.protection.locked for style in _row_formats.distinct_styles(sheet, area))
    return found


def any_locked(sheet: Worksheet, areas: list[Area]) -> bool:
    return True in _lock_states(sheet, areas)


class Gate:
    """A macro's write in progress on a protected sheet, and whether it has passed over a locked cell."""

    def __init__(self, sheet: Worksheet) -> None:
        self.sheet = sheet
        self.refused = False

    def admits(self, row: int, column: int) -> bool:
        if locked(self.sheet, row, column):
            self.refused = True
            return False
        return True


@contextmanager
def writing(sheet: Worksheet) -> Generator[None]:
    """A write of values or formulas: every unlocked cell takes it, and a locked one among them refuses it after."""
    if enforced(sheet) is None or sheet.write_gate is not None:
        yield
        return
    gate = Gate(sheet)
    sheet.write_gate = gate
    try:
        yield
    finally:
        sheet.write_gate = None
    if gate.refused:
        raise error(1004, PROTECTED)


def whole(sheet: Worksheet, areas: list[Area], what: str) -> None:
    """An edit Excel was measured making only when none or every cell is locked: all locked refuses it."""
    if enforced(sheet) is None:
        return
    states = _lock_states(sheet, areas)
    if True not in states:
        return
    if False not in states:
        raise error(1004, PROTECTED)
    raise VBAUnsupportedError(f"{what} over locked and unlocked cells of a protected sheet is not implemented")
