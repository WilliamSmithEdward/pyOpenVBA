"""Workbook settings a macro reads and sets, measured in live Excel (tests/fixtures/workbook_settings.json,
scripts/measure_workbook_settings.py):

- DefaultTableStyle and DefaultPivotTableStyle are TableStyle objects, set by one or by a built-in style's name in
  any case; a name Excel does not know, "" or a number leaves no default, which reads as the text "None". A save
  writes them as styles.xml's tableStyles defaultTableStyle and defaultPivotStyle, and leaves a missing one out. A
  table ListObjects.Add makes takes the default table style.
- DefaultSlicerStyle and DefaultTimelineStyle are SlicerStyleLight1 and TimeSlicerStyleLight1; a name Excel does
  not know is error 450. Setting a known one, which a save writes outside the parts measured, reports itself
  unsupported.
- ShowPivotTableFieldList, ShowPivotChartActiveFields and InactiveListBorderVisible are workbookPr's
  hidePivotFieldList, showPivotChartFilter and showBorderUnselectedTables, the first turned round;
  text is error 13.
- CreateBackup and HasPassword are False and error 1004 to set; CaseSensitive, UseWholeCellCriteria and
  UseWildcards are False, True and True and error 450 to set; Final is False.
- UpdateRemoteReferences and ChartDataPointTrack are True and AccuracyVersion 0; text is error 1004 to the first
  and 13 to the second. Setting them, which a save writes outside the parts measured, reports itself unsupported.
"""
from __future__ import annotations

import re
from typing import TYPE_CHECKING, Final

from pyopenvba.apps.excel._creator import ItemDefault
from pyopenvba.apps.excel._model import ExcelObject
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import NULL, VBAInt, error, to_bool, to_integer

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Workbook


def _names(prefix: str, groups: tuple[tuple[str, int], ...]) -> tuple[str, ...]:
    return tuple(f"{prefix}{group}{number}" for group, count in groups for number in range(1, count + 1))


TABLE_STYLES: Final = _names("TableStyle", (("Light", 21), ("Medium", 28), ("Dark", 11)))
PIVOT_STYLES: Final = _names("PivotStyle", (("Light", 28), ("Medium", 28), ("Dark", 28)))
SLICER_STYLES: Final = _names("SlicerStyle", (("Light", 6), ("Other", 2), ("Dark", 6)))
#: A timeline has no Other styles (tests/fixtures/table_styles.json).
TIMELINE_STYLES: Final = _names("TimeSlicerStyle", (("Light", 6), ("Dark", 6)))
#: Every built-in style in the order TableStyles lists them, 170 in all.
BUILT_IN_STYLES: Final = TABLE_STYLES + PIVOT_STYLES + SLICER_STYLES + TIMELINE_STYLES
_BUILT_IN: Final = {name.lower(): name for name in BUILT_IN_STYLES}
#: Each family's prefix, the gallery member that shows it, and its name in English Excel's own words.
_FAMILIES: Final = (("TimeSlicerStyle", "ShowAsAvailableTimelineStyle", "Timeline Style"),
                    ("TableStyle", "ShowAsAvailableTableStyle", "Table Style"),
                    ("PivotStyle", "ShowAsAvailablePivotTableStyle", "Pivot Style"),
                    ("SlicerStyle", "ShowAsAvailableSlicerStyle", "Slicer Style"))


def _custom_styles(book: Workbook) -> bool:
    """Whether the workbook's stylesheet holds a table, slicer or timeline style of its own."""
    return re.search(r"<(?:[\w.-]+:)?(?:tableStyle|slicerStyle|timelineStyle)\b", book.stylesheet.xml) is not None


class TableStyles(VBACollection, ItemDefault, ExcelObject):
    """The workbook's table styles: the 170 built in, in Excel's order (tests/fixtures/table_styles.json)."""

    vba_type_name = "TableStyles"

    def __init__(self, book: Workbook) -> None:
        self.book = book

    def vba_items(self) -> list[object]:
        if _custom_styles(self.book):
            raise VBAUnsupportedError("TableStyles of a workbook with styles of its own is not implemented")
        return [TableStyle(self.book, name) for name in BUILT_IN_STYLES]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            found = _BUILT_IN.get(index.lower())
            if found is None and _custom_styles(self.book):
                raise VBAUnsupportedError("TableStyles of a style of the workbook's own is not implemented")
            if found is None:
                raise error(9, "Subscript out of range")
            return TableStyle(self.book, found)
        place = int(to_integer(index, "Long"))
        if not 1 <= place <= len(items):
            raise error(9, "Subscript out of range")
        return items[place - 1]

    @member
    def Parent(self) -> object:
        return self.book

    @member
    def Application(self) -> object:
        return self.book.application


class TableStyle(ExcelObject):
    """A built-in table, pivot, slicer or timeline style, by its name."""

    vba_type_name = "TableStyle"

    def __init__(self, book: Workbook, name: str) -> None:
        self.book = book
        self.name = name

    def family(self) -> tuple[str, str, str]:
        return next(family for family in _FAMILIES if self.name.startswith(family[0]))

    @method(name="_Default")
    def hidden_default(self) -> object:
        return self.name

    @member(default=True)
    def Name(self) -> object:
        return self.name

    @member
    def NameLocal(self) -> object:
        """The name in English Excel's own words: TableStyleMedium2 reads Table Style Medium 2."""
        prefix, _, words = self.family()
        found = re.fullmatch(r"([A-Za-z]+?)(\d+)", self.name[len(prefix):])
        assert found is not None
        return f"{words} {found[1]} {found[2]}"

    @member
    def BuiltIn(self) -> object:
        return True

    @member
    def Parent(self) -> object:
        return TableStyles(self.book)

    @member
    def Application(self) -> object:
        return self.book.application

    def shown_in(self, gallery: str) -> object:
        return self.family()[1] == gallery

    def set_shown(self, gallery: str, value: object) -> None:
        # A built-in style keeps the gallery it shows in: changing it is error 5.
        if to_bool(value) == self.shown_in(gallery):
            raise VBAUnsupportedError(f"TableStyle.{gallery} set to what it reads is not implemented")
        raise error(5, "Invalid procedure call or argument")

    @member
    def ShowAsAvailableTableStyle(self) -> object:
        return self.shown_in("ShowAsAvailableTableStyle")

    @setter("ShowAsAvailableTableStyle")
    def _show_table(self, value: object) -> None:
        self.set_shown("ShowAsAvailableTableStyle", value)

    @member
    def ShowAsAvailablePivotTableStyle(self) -> object:
        return self.shown_in("ShowAsAvailablePivotTableStyle")

    @setter("ShowAsAvailablePivotTableStyle")
    def _show_pivot(self, value: object) -> None:
        self.set_shown("ShowAsAvailablePivotTableStyle", value)

    @member
    def ShowAsAvailableSlicerStyle(self) -> object:
        return self.shown_in("ShowAsAvailableSlicerStyle")

    @setter("ShowAsAvailableSlicerStyle")
    def _show_slicer(self, value: object) -> None:
        self.set_shown("ShowAsAvailableSlicerStyle", value)

    @member
    def ShowAsAvailableTimelineStyle(self) -> object:
        return self.shown_in("ShowAsAvailableTimelineStyle")

    @setter("ShowAsAvailableTimelineStyle")
    def _show_timeline(self, value: object) -> None:
        self.set_shown("ShowAsAvailableTimelineStyle", value)

    @method
    def Delete(self) -> object:
        # A built-in style cannot be deleted (tests/fixtures/table_styles.json).
        raise error(1004, "Delete method of TableStyle class failed")


def table_style(book: Workbook, name: object) -> TableStyle:
    """Workbook.TableStyles(name): a built-in style, by its name in any case or by number."""
    found = TableStyles(book).vba_get("Item", [name])
    assert isinstance(found, TableStyle)
    return found


def _style_name(value: object) -> str | None:
    """The built-in style a value names, None for anything else."""
    if isinstance(value, TableStyle):
        return value.name
    if isinstance(value, str):
        return _BUILT_IN.get(value.lower())
    return None


def default_style(book: Workbook, kind: str) -> object:
    """DefaultTableStyle ("table") or DefaultPivotTableStyle ("pivot")."""
    name = book.stylesheet.table_style_default(kind)
    return "None" if name is None else TableStyle(book, name)


def set_default_style(book: Workbook, kind: str, value: object) -> None:
    name = _style_name(value)
    allowed = TABLE_STYLES if kind == "table" else PIVOT_STYLES
    book.stylesheet.set_table_style_default(kind, name if name in allowed else None)
    book.saved = False


def set_slicer_default(value: object, styles: tuple[str, ...], name: str) -> None:
    if _style_name(value) not in styles:
        raise error(450, "Wrong number of arguments or invalid property assignment")
    raise VBAUnsupportedError(f"Workbook.{name} set to another style is not implemented")


#: The workbookPr flags: the setting, the attribute, whether the setting is the attribute turned round, and the
#: setting a workbook has where the file leaves the attribute out.
FLAGS: Final = {
    "ShowPivotTableFieldList": ("hidePivotFieldList", True, True),
    "ShowPivotChartActiveFields": ("showPivotChartFilter", False, False),
    "InactiveListBorderVisible": ("showBorderUnselectedTables", False, True),
}


def flag(book: Workbook, name: str) -> object:
    attribute, inverted, default = FLAGS[name]
    found = book.workbook_pr.get(attribute)
    if found is None:
        return default
    on = found in ("1", "true")
    return not on if inverted else on


def set_flag(book: Workbook, name: str, value: object) -> None:
    if value is NULL:
        raise VBAUnsupportedError(f"Workbook.{name} set to Null is not implemented")
    wanted = to_bool(value)
    if wanted == flag(book, name):
        return
    attribute, inverted, _default = FLAGS[name]
    book.workbook_pr[attribute] = "1" if wanted != inverted else "0"
    book.workbook_pr_changed.add(attribute)
    book.saved = False


def refused(name: str, number: int) -> None:
    """A setting Excel will not take: error 1004 or 450."""
    if number == 1004:
        raise error(1004, f"Unable to set the {name} property of the Workbook class")
    raise error(450, "Wrong number of arguments or invalid property assignment")


def unmeasured_set(name: str, value: object, text_error: int) -> None:
    """A setting whose place in a saved file was not measured: text is the error Excel gives, anything else
    reports itself unsupported."""
    if isinstance(value, str):
        try:
            to_bool(value)
        except VBARuntimeError:
            if text_error == 1004:
                refused(name, 1004)
            raise
    raise VBAUnsupportedError(f"Workbook.{name} set is not implemented")


ACCURACY_VERSION: Final = VBAInt(0, "Long")

#: Workbook properties as a new workbook, and one saved and opened again, read them
#: (tests/fixtures/workbook_properties.json); those a macro may set it sets for the session, nothing of them
#: reaching the file.
SESSION: Final[dict[str, object]] = {
    "AutoSaveOn": False, "AutoUpdateFrequency": 0, "ChangeHistoryDuration": 0, "CompatibilityVersion": 2,
    "ConflictResolution": 1, "ConnectionsDisabled": False, "DoNotPromptForConvert": True, "EncryptionProvider": "",
    "Excel8CompatibilityMode": False, "HighlightChangesOnScreen": False, "IsAddin": False, "IsInplace": False,
    "KeepChangeHistory": True, "ListChangesOnNewSheet": False, "MultiUserEditing": False,
    "PersonalViewListSettings": True, "PersonalViewPrintSettings": True, "RevisionNumber": 0,
    "ShowConflictHistory": False, "TemplateRemoveExtData": False, "VBASigned": False, "WriteReserved": False,
    "EnvelopeVisible": False,
}
#: Those a set changes: a Boolean, or a Long.
SETTABLE: Final = {"PersonalViewListSettings": bool, "PersonalViewPrintSettings": bool, "TemplateRemoveExtData": bool,
                   "DoNotPromptForConvert": bool, "KeepChangeHistory": bool, "ConflictResolution": int, "IsAddin": bool}
#: Those a set leaves as they are: ChangeHistoryDuration, which belongs to a shared workbook.
IGNORED_SET: Final = frozenset({"ChangeHistoryDuration"})
#: Those Excel refuses to set on a workbook not shared, and the message it gives.
REFUSED_SET: Final = {"HighlightChangesOnScreen": "", "ListChangesOnNewSheet": "", "AutoUpdateFrequency": "",
                      "AutoSaveOn": "", "ShowConflictHistory": "Unable to set the ShowConflictHistory property of the "
                                                               "Workbook class"}


def session_value(book: Workbook, name: str) -> object:
    value = book.session_properties.get(name, SESSION[name])
    return VBAInt(value, "Long") if isinstance(value, int) and not isinstance(value, bool) else value


def set_session_value(book: Workbook, name: str, value: object) -> None:
    if name in REFUSED_SET:
        raise error(1004, REFUSED_SET[name] or "Application-defined or object-defined error")
    if name in IGNORED_SET:
        to_integer(value)
        return
    kind = SETTABLE.get(name)
    if kind is None:
        raise VBAUnsupportedError(f"Workbook.{name} set is not implemented")
    book.session_properties[name] = to_bool(value) if kind is bool else int(to_integer(value))
