"""Workbook.BuiltinDocumentProperties and CustomDocumentProperties, as Excel keeps a workbook's properties and saves
them as its docProps parts (tests/fixtures/document_properties/, scripts/measure_document_properties.py).

A workbook has Office's 34 built-in properties, each of a type: text, a whole number or a date. A new workbook's
author is the user, its company empty, its application Microsoft Excel and its creation date the moment it was
made; a workbook opened from a file has what the file keeps. A property nothing set reads Empty where it is text
and fails where it is a number or a date. A value set is turned into the property's type as Office turns a
Variant: a Boolean becomes -1 or 0 in text, and Null or text that is no number is a type mismatch. Custom
properties are added by name, any case one name, as text, a whole number, a Boolean, a date or a number, or
linked to a named range, which they take their value and its type from when the workbook is saved.

Office keeps a date as a moment in UTC, whole seconds of it. A Date from day zero on is cut to its second; one
before day zero is taken as VBA shows it, rounded, and then a day earlier, as Office does, and one that comes
before 1601 fails. The local time goes to UTC by the machine's time zone, as Windows gives it for each year,
daylight time in the hour a clock skips or repeats, so a time in the hour skipped reads back an hour early.
Elsewhere the local zone is Python's.

A save writes the last author as the user and the last save time as the moment, a creation date where there is
none, and core.xml and app.xml anew: each property where Excel keeps it, in Excel's order, app.xml's sheets and
named ranges as Excel lists them, and custom.xml while the workbook has a custom property, with its relationship
and content type. The number of bytes and the content type are kept by no part, and Excel writes its own
application and security.
"""
from __future__ import annotations

import calendar
import datetime as dt
import functools
import math
import re
import sys
import time
from collections.abc import Callable
from dataclasses import dataclass, field, replace
from typing import TYPE_CHECKING, Final

from pyopenvba._a1 import Area, quote_sheet, split_sheet
from pyopenvba._opc_xml import opc_part
from pyopenvba._relationships import in_office_order
from pyopenvba._xml import unescape
from pyopenvba.apps.excel._model import ExcelObject, Range
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.formula._parse import Binary, FormulaError, NameNode, Node, Reference, Structured, parse
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import (
    EMPTY,
    MISSING,
    NULL,
    VBADate,
    VBAErrorValue,
    VBAInt,
    error,
    number_text,
    to_bool,
    to_date,
    to_integer,
    to_number,
    to_text,
)

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Workbook
    from pyopenvba.powerquery._opc import OpcFile

#: msoPropertyTypeNumber, Boolean, Date, String and Float.
NUMBER, BOOLEAN, DATE, STRING, FLOAT = 1, 2, 3, 4, 5
UNSPECIFIED: Final = -2147467259
CATASTROPHIC: Final = -2147418113
#: The most characters a custom property's name keeps; Add cuts a longer one (custom_errors).
_LONGEST: Final = 255
#: What VBA's Err.Description reads for each, a space at its end (live probe).
_UNSPECIFIED_TEXT: Final = "Automation error\nUnspecified error "
_CATASTROPHIC_TEXT: Final = "Automation error\nCatastrophic failure "
FMTID: Final = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
CORE_PART, APP_PART, CUSTOM_PART = "docProps/core.xml", "docProps/app.xml", "docProps/custom.xml"
_CORE_NS: Final = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
_APP_NS: Final = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
_CUSTOM_NS: Final = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
_VT_NS: Final = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
_PREFIXES: Final = {"http://purl.org/dc/elements/1.1/": "dc", "http://purl.org/dc/terms/": "dcterms", _CORE_NS: "cp"}
#: Each part's relationship from the package and its content type, in the order Excel numbers the relationships.
_PARTS: Final = {
    CORE_PART: ("http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
                "application/vnd.openxmlformats-package.core-properties+xml"),
    APP_PART: ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
               "application/vnd.openxmlformats-officedocument.extended-properties+xml"),
    CUSTOM_PART: ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
                  "application/vnd.openxmlformats-officedocument.custom-properties+xml"),
}
_DECLARATION: Final = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

#: Office's built-in properties in their order: the name, the type and where a save keeps each, "" for nowhere.
BUILTIN: Final = (
    ("Title", STRING, "dc:title"), ("Subject", STRING, "dc:subject"), ("Author", STRING, "dc:creator"),
    ("Keywords", STRING, "cp:keywords"), ("Comments", STRING, "dc:description"), ("Template", STRING, "Template"),
    ("Last author", STRING, "cp:lastModifiedBy"), ("Revision number", STRING, "cp:revision"),
    ("Application name", STRING, "Application"), ("Last print date", DATE, "cp:lastPrinted"),
    ("Creation date", DATE, "dcterms:created"), ("Last save time", DATE, "dcterms:modified"),
    ("Total editing time", NUMBER, "TotalTime"), ("Number of pages", NUMBER, "Pages"),
    ("Number of words", NUMBER, "Words"), ("Number of characters", NUMBER, "Characters"),
    ("Security", NUMBER, "DocSecurity"), ("Category", STRING, "cp:category"), ("Format", STRING, "PresentationFormat"),
    ("Manager", STRING, "Manager"), ("Company", STRING, "Company"), ("Number of bytes", NUMBER, ""),
    ("Number of lines", NUMBER, "Lines"), ("Number of paragraphs", NUMBER, "Paragraphs"),
    ("Number of slides", NUMBER, "Slides"), ("Number of notes", NUMBER, "Notes"),
    ("Number of hidden Slides", NUMBER, "HiddenSlides"), ("Number of multimedia clips", NUMBER, "MMClips"),
    ("Hyperlink base", STRING, "HyperlinkBase"), ("Number of characters (with spaces)", NUMBER, "CharactersWithSpaces"),
    ("Content type", STRING, ""), ("Content status", STRING, "cp:contentStatus"), ("Language", STRING, "dc:language"),
    ("Document version", STRING, "cp:version"),
)
_AT: Final = {element: index for index, (_name, _kind, element) in enumerate(BUILTIN) if element}
AUTHOR, LAST_AUTHOR, APPLICATION, CREATED, SAVED, SECURITY, COMPANY = (
    _AT[element] for element in ("dc:creator", "cp:lastModifiedBy", "Application", "dcterms:created",
                                 "dcterms:modified", "DocSecurity", "Company"))
#: core.xml's elements in Excel's order.
_CORE: Final = ("dc:title", "dc:subject", "dc:creator", "cp:keywords", "dc:description", "cp:lastModifiedBy",
                "cp:revision", "cp:lastPrinted", "dcterms:created", "dcterms:modified", "cp:category",
                "cp:contentStatus", "dc:language", "cp:version")
#: app.xml's elements in Excel's order, the lists of sheets and names and the settings Excel writes the same way
#: every time among them.
_APP: Final = ("Template", "TotalTime", "Pages", "Words", "Characters", "Application", "DocSecurity",
               "PresentationFormat", "Lines", "Paragraphs", "Slides", "Notes", "HiddenSlides", "MMClips", "ScaleCrop",
               "HeadingPairs", "TitlesOfParts", "Manager", "Company", "LinksUpToDate", "CharactersWithSpaces",
               "SharedDoc", "HyperlinkBase", "HyperlinksChanged", "AppVersion")
_SETTINGS: Final = {"ScaleCrop": "false", "LinksUpToDate": "false", "SharedDoc": "false", "HyperlinksChanged": "false",
                    "AppVersion": "16.0300"}
_VALUE_TAGS: Final = {STRING: "lpwstr", NUMBER: "i4", BOOLEAN: "bool", DATE: "filetime", FLOAT: "r8"}
#: The value types a custom property is read from, Excel's own and the others other writers use for the same kinds.
_READ_TAGS: Final = {"lpwstr": STRING, "lpstr": STRING, "bstr": STRING, "i1": NUMBER, "i2": NUMBER, "i4": NUMBER,
                     "int": NUMBER, "ui1": NUMBER, "ui2": NUMBER, "bool": BOOLEAN, "filetime": DATE, "r4": FLOAT,
                     "r8": FLOAT}
_ERROR_TEXT: Final = {2000: "#NULL!", 2007: "#DIV/0!", 2015: "#VALUE!", 2023: "#REF!", 2029: "#NAME?", 2036: "#NUM!",
                      2042: "#N/A", 2043: "#GETTING_DATA", 2045: "#SPILL!", 2046: "#CONNECT!", 2047: "#BLOCKED!",
                      2048: "#UNKNOWN!", 2049: "#FIELD!", 2050: "#CALC!"}


def _failed(name: str) -> VBARuntimeError:
    return error(UNSPECIFIED, f"Method '{name}' of object 'DocumentProperty' failed")


def _gone(name: str) -> VBARuntimeError:
    """What a property deleted under a macro's reference to it answers: E_UNEXPECTED, for most members."""
    return error(CATASTROPHIC, f"Method '{name}' of object 'DocumentProperty' failed")


class _Early(Exception):
    """A date before 1601, which Office's moments start at."""


# --- moments ---------------------------------------------------------------------------------------------------

_ORIGIN: Final = dt.datetime(1, 1, 1)
#: 400 years of the Gregorian calendar, which repeats after them, in seconds.
_CYCLE: Final = 146097 * 86400
_FIRST: Final = dt.datetime(1601, 1, 1)


def _seconds(moment: dt.datetime) -> int:
    return (moment - _ORIGIN) // dt.timedelta(seconds=1)


def _moment(seconds: int) -> dt.datetime:
    return _ORIGIN + dt.timedelta(seconds=seconds)


#: The first second past 9999, which a moment on 9999-12-31 reaches in UTC west of Greenwich.
_PAST: Final = _seconds(dt.datetime(9999, 12, 31)) + 86400


def _w3c(utc: int) -> str:
    """A UTC moment as Excel writes it, a year past 9999 counted from 0000 (tests/fixtures/document_properties/,
    dates_overflow): a file Excel writes but cannot open."""
    over = utc >= _PAST
    moment = _moment(utc - _CYCLE if over else utc)
    year = (moment.year + 400) % 10000 if over else moment.year
    return f"{year:04d}" + moment.strftime("-%m-%dT%H:%M:%SZ")


def _read_w3c(text: str) -> int:
    """A W3CDTF moment, in UTC: Excel's own spelling, or another writer's with a part second or an offset."""
    found = re.fullmatch(r"\s*(\d{4})-(\d\d)-(\d\d)T(\d\d):(\d\d):(\d\d)(?:\.\d+)?(Z|[+-]\d\d:\d\d)\s*", text)
    if found is None or found.group(1) == "0000":
        raise VBAUnsupportedError(f"a document property's date written {text!r} is not implemented")
    year, month, day, hour, minute, second = (int(part) for part in found.groups()[:6])
    moment = dt.datetime(year, month, day, hour, minute, second)
    zone = found.group(7)
    if zone != "Z":
        offset = (int(zone[1:3]) * 60 + int(zone[4:6])) * (1 if zone[0] == "+" else -1)
        moment -= dt.timedelta(minutes=offset)
    return _seconds(moment)


# --- the time zone ---------------------------------------------------------------------------------------------


@dataclass(frozen=True)
class ZoneRule:
    """A time zone as Windows gives it for a year (GetTimeZoneInformationForYear): UTC is the local time plus bias
    minutes, plus standard_bias or daylight_bias. Daylight time begins at `daylight`, on the standard clock, and
    ends at `standard`, on the daylight clock, each a SYSTEMTIME's year, month, day of the week, day, hour, minute
    and second: no month where the zone has no daylight time, and no year where the day is the week of the month
    the weekday falls in, 5 its last."""

    bias: int
    standard_bias: int
    daylight_bias: int
    standard: tuple[int, ...]
    daylight: tuple[int, ...]


if sys.platform == "win32":

    @functools.cache
    def _system_rule(year: int) -> ZoneRule:
        import ctypes
        from ctypes import wintypes

        class SYSTEMTIME(ctypes.Structure):
            _fields_ = [(part, wintypes.WORD) for part in
                        ("wYear", "wMonth", "wDayOfWeek", "wDay", "wHour", "wMinute", "wSecond", "wMilliseconds")]

        class TIME_ZONE_INFORMATION(ctypes.Structure):
            _fields_ = [("Bias", wintypes.LONG), ("StandardName", wintypes.WCHAR * 32), ("StandardDate", SYSTEMTIME),
                        ("StandardBias", wintypes.LONG), ("DaylightName", wintypes.WCHAR * 32),
                        ("DaylightDate", SYSTEMTIME), ("DaylightBias", wintypes.LONG)]

        def change(found: SYSTEMTIME) -> tuple[int, ...]:
            return (found.wYear, found.wMonth, found.wDayOfWeek, found.wDay, found.wHour, found.wMinute,
                    found.wSecond)

        kernel = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel.GetTimeZoneInformationForYear.argtypes = [wintypes.USHORT, ctypes.c_void_p,
                                                         ctypes.POINTER(TIME_ZONE_INFORMATION)]
        info = TIME_ZONE_INFORMATION()
        if not kernel.GetTimeZoneInformationForYear(year, None, ctypes.byref(info)):
            raise OSError(ctypes.get_last_error(), "GetTimeZoneInformationForYear failed")
        return ZoneRule(info.Bias, info.StandardBias, info.DaylightBias, change(info.StandardDate),
                        change(info.DaylightDate))

    zone_rule: Callable[[int], ZoneRule] | None = _system_rule
else:
    #: The rule for a year, or None for Python's own view of the local zone, which the tests replace.
    zone_rule: Callable[[int], ZoneRule] | None = None


def _change(year: int, when: tuple[int, ...]) -> dt.datetime | None:
    """The wall-clock moment in `year` a rule's change falls on, or None where there is none."""
    rule_year, month, weekday, day, hour, minute, second = when[:7]
    if month == 0:
        return None
    if rule_year == 0:
        first = (dt.date(year, month, 1).weekday() + 1) % 7
        day = 1 + (weekday - first) % 7 + 7 * (day - 1)
        while day > calendar.monthrange(year, month)[1]:
            day -= 7
    return dt.datetime(year, month, day) + dt.timedelta(hours=hour, minutes=minute, seconds=second)


def _in_daylight(local: dt.datetime, rule: ZoneRule) -> bool:
    """Whether a local time is in daylight time: not before the change forward, read on the standard clock, and
    before the change back, read on the daylight one, so the hours a clock skips and repeats are both in it."""
    ends, begins = _change(local.year, rule.standard), _change(local.year, rule.daylight)
    if ends is None or begins is None:
        return False
    before_end, after_start = local < ends, local >= begins
    if rule.daylight[1] < rule.standard[1]:
        return before_end and after_start
    return before_end or after_start


def to_utc(local: dt.datetime) -> int:
    """A local moment's UTC, in seconds from 0001-01-01: a time the clock skips or repeats is in daylight time, as
    Excel has it (dates_zones), though Windows' own TzSpecificLocalTimeToSystemTimeEx takes a skipped one as
    standard time. Back from UTC the two agree."""
    rule = zone_rule(local.year) if zone_rule is not None else None
    if rule is None:
        try:
            offset = local.astimezone().utcoffset()
            bias = -offset // dt.timedelta(minutes=1) if offset is not None else 0
        except (OverflowError, OSError, ValueError):
            bias = time.timezone // 60
    else:
        bias = rule.bias + (rule.daylight_bias if _in_daylight(local, rule) else rule.standard_bias)
    return _seconds(local) + bias * 60


def to_local(utc: int) -> dt.datetime:
    """A UTC moment's local time; one past 9999 is a time on 9999-12-31, worked out 400 years earlier, when the
    calendar was the same."""
    if utc >= _PAST:
        local = to_local(utc - _CYCLE)
        return local.replace(year=local.year + 400)
    if zone_rule is None:
        try:
            return _moment(utc).replace(tzinfo=dt.timezone.utc).astimezone().replace(tzinfo=None)
        except (OverflowError, OSError, ValueError):
            return _moment(utc - time.timezone)
    # The year is the one the standard clock reads, and the changes of that year are put in UTC to compare.
    rule = zone_rule(_moment(utc).year)
    year = _moment(utc - (rule.bias + rule.standard_bias) * 60).year
    rule = zone_rule(year)
    on_standard = _moment(utc - (rule.bias + rule.standard_bias) * 60)
    on_daylight = _moment(utc - (rule.bias + rule.daylight_bias) * 60)
    ends, begins = _change(year, rule.standard), _change(year, rule.daylight)
    if ends is None or begins is None:
        return on_standard
    after_start = utc >= _seconds(begins) + (rule.bias + rule.standard_bias) * 60
    before_end = utc < _seconds(ends) + (rule.bias + rule.daylight_bias) * 60
    daylight = (before_end and after_start) if rule.daylight[1] < rule.standard[1] else (before_end or after_start)
    return on_daylight if daylight else on_standard


def kept_moment(date: VBADate) -> dt.datetime:
    """The local moment Office keeps of a Date: cut to its second from day zero on; before day zero the moment VBA
    shows, rounded to its second, a day earlier (tests/fixtures/document_properties/, dates_zones and
    dates_early)."""
    if date.serial >= 0:
        days = math.floor(date.serial)
        seconds = math.floor((date.serial - days) * 86400.0)
        return VBADate.EPOCH + dt.timedelta(days=days, seconds=seconds)
    return date.to_datetime() - dt.timedelta(days=1)


def _stored_date(value: object) -> int:
    moment = kept_moment(to_date(value))
    if moment < _FIRST:
        raise _Early
    return to_utc(moment)


# --- values ----------------------------------------------------------------------------------------------------


def coerced(value: object, kind: int) -> object:
    """A value as a property of the type keeps it, turned as Office turns a Variant; a date as its moment in UTC.
    Null, and anything that does not turn, is a type mismatch, a number too large to hold an overflow."""
    if value is NULL:
        raise error(13)
    try:
        if kind == STRING:
            if isinstance(value, bool):
                return "-1" if value else "0"
            return to_text(value)
        if kind == NUMBER:
            return VBAInt(int(to_integer(value, "Long")), "Long")
        if kind == BOOLEAN:
            return to_bool(value)
        if kind == DATE:
            return _stored_date(value)
        if isinstance(value, VBADate):
            return value.serial
        return float(to_number(value))
    except VBARuntimeError as exc:
        if exc.number == 6:
            raise
        raise error(13) from None


def shown(kind: int, stored: object) -> object:
    """What VBA reads of a value kept as `coerced` keeps it."""
    if kind == DATE:
        assert isinstance(stored, int)
        return VBADate.from_datetime(to_local(stored))
    return stored


def _linked_default(kind: int) -> object:
    """What a linked property holds until a save gives it its range's value: a date of 1899-12-31."""
    defaults: dict[int, object] = {STRING: "", NUMBER: VBAInt(0, "Long"), BOOLEAN: False, FLOAT: 0.0}
    return defaults[kind] if kind in defaults else to_utc(dt.datetime(1899, 12, 31))


# --- the state -------------------------------------------------------------------------------------------------


@dataclass
class Custom:
    name: str
    kind: int
    #: The value as `coerced` keeps it.
    value: object
    #: The named range a linked property takes its value from, "" for one that is not linked.
    link: str = ""


def _builtin() -> list[object | None]:
    return [None] * len(BUILTIN)


def _custom() -> list[Custom]:
    return []


@dataclass
class Properties:
    #: Each built-in property's value, as `coerced` keeps it, or None while nothing has set it.
    builtin: list[object | None] = field(default_factory=_builtin)
    custom: list[Custom] = field(default_factory=_custom)
    #: True where the file's custom.xml holds a value of a type the model does not read: the part is kept as it is.
    opaque: bool = False
    #: True once RemoveDocumentInformation took the document properties away: a save then writes no last author.
    last_author_removed: bool = False


def state(book: Workbook) -> Properties:
    """The workbook's properties: those its file keeps, or a new workbook's. A new workbook has no path until its
    first save is done, though the save gives it the template's package to fill."""
    if book.properties is None:
        if book.package is not None and book.path:
            book.properties = _read(book.package)
        else:
            found = Properties()
            found.builtin[AUTHOR] = book.made_by
            found.builtin[APPLICATION] = "Microsoft Excel"
            found.builtin[COMPANY] = ""
            found.builtin[CREATED] = to_utc(book.made_at)
            book.properties = found
    return book.properties


# --- reading the parts -----------------------------------------------------------------------------------------

_ESCAPED: Final = re.compile(r"_x([0-9A-Fa-f]{4})_")


def _text(raw: str) -> str:
    """The text a docProps part keeps, its entities, line ends and _xhhhh_ escapes read."""
    text = unescape(raw).replace("\r\n", "\n").replace("\r", "\n")
    return _ESCAPED.sub(lambda found: chr(int(found.group(1), 16)), text)


def _children(data: bytes, namespace: str, root: str) -> list[tuple[str, dict[str, str], str]]:
    """A part's root's children: each one's expanded name, attributes and the markup inside it."""
    try:
        part = opc_part(data, namespace, root, max_children=10000)
    except ValueError as exc:
        raise VBAUnsupportedError(f"a docProps part that cannot be read is not implemented: {exc}") from None
    found: list[tuple[str, dict[str, str], str]] = []
    for name, attributes, start, end in part.children:
        element = data[start:end].decode("utf-8")
        opening = re.match(r"<[^>]*>", element)
        assert opening is not None
        inner = "" if opening.group().endswith("/>") else element[opening.end():element.rfind("</")]
        found.append((name, attributes, inner))
    return found


def _read(package: OpcFile) -> Properties:
    found = Properties()
    if package.has(CORE_PART):
        for name, _attributes, inner in _children(package.read(CORE_PART), _CORE_NS, "coreProperties"):
            namespace, _, local = name.rpartition("}")
            index = _AT.get(f"{_PREFIXES.get(namespace, '')}:{local}")
            if index is not None:
                found.builtin[index] = _read_builtin(index, inner)
    if package.has(APP_PART):
        for name, _attributes, inner in _children(package.read(APP_PART), _APP_NS, "Properties"):
            namespace, _, local = name.rpartition("}")
            index = _AT.get(local) if namespace == _APP_NS else None
            if index is not None:
                found.builtin[index] = _read_builtin(index, inner)
    if package.has(CUSTOM_PART):
        try:
            found.custom = [_read_custom(attributes, inner) for name, attributes, inner
                            in _children(package.read(CUSTOM_PART), _CUSTOM_NS, "Properties")
                            if name == _CUSTOM_NS + "}property"]
        except VBAUnsupportedError:
            found.opaque = True
    return found


def _read_builtin(index: int, inner: str) -> object:
    kind = BUILTIN[index][1]
    if kind == DATE:
        return _read_w3c(_text(inner))
    if kind == NUMBER:
        try:
            return VBAInt(int(_text(inner)), "Long")
        except ValueError:
            raise VBAUnsupportedError(f"the number {BUILTIN[index][0]} written {inner!r} is not implemented") from None
    return _text(inner)


def _read_custom(attributes: dict[str, str], inner: str) -> Custom:
    value = re.fullmatch(r"\s*<(?:[\w.-]+:)?(\w+)\b[^>]*?(?:/>|>(.*)</(?:[\w.-]+:)?\1>)\s*", inner, re.DOTALL)
    if value is None or value.group(1) not in _READ_TAGS:
        raise VBAUnsupportedError("a custom document property of a type Excel does not write is not implemented")
    kind, text = _READ_TAGS[value.group(1)], _text(value.group(2) or "")
    name = _ESCAPED.sub(lambda found: chr(int(found.group(1), 16)), attributes.get("name", ""))
    try:
        if kind == STRING:
            kept: object = text
        elif kind == NUMBER:
            kept = VBAInt(int(text), "Long")
        elif kind == BOOLEAN:
            kept = text.strip().lower() in ("true", "1")
        elif kind == DATE:
            kept = _read_w3c(text)
        else:
            kept = float(text)
    except (ValueError, OverflowError):
        raise VBAUnsupportedError(f"the custom document property {name!r} written {text!r} is not implemented") \
            from None
    return Custom(name, kind, kept, attributes.get("linkTarget", ""))


# --- the objects -----------------------------------------------------------------------------------------------


class DocumentProperties(VBACollection, ExcelObject):
    """A workbook's built-in or custom properties."""

    vba_type_name = "DocumentProperties"

    def __init__(self, book: Workbook, builtin: bool) -> None:
        self.book = book
        self.builtin = builtin

    def _found(self) -> Properties:
        found = state(self.book)
        if found.opaque and not self.builtin:
            raise VBAUnsupportedError("custom document properties of a type Excel does not write are not "
                                      "implemented")
        return found

    def vba_items(self) -> list[object]:
        if self.builtin:
            return [DocumentProperty(self.book, index) for index in range(len(BUILTIN))]
        return [DocumentProperty(self.book, -1, entry) for entry in self._found().custom]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        """By name, any case one name, or by place; anything else is error 5."""
        if isinstance(index, str):
            wanted = index.casefold()
            for item in items:
                assert isinstance(item, DocumentProperty)
                if item.name().casefold() == wanted:
                    return item
            raise error(5)
        place = int(to_integer(index, "Long"))
        if not 1 <= place <= len(items):
            raise error(5)
        return items[place - 1]

    @member
    def Parent(self) -> object:
        return self.book

    @member
    def Application(self) -> object:
        return self.book.application

    @member
    def Creator(self) -> object:
        return VBAInt(1480803660, "Long")

    @method
    def Add(self, Name: object = MISSING, LinkToContent: object = MISSING, Type: object = MISSING,  # noqa: A002
            Value: object = MISSING, LinkSource: object = MISSING) -> object:
        """A custom property of the type given, its value or a named range's (tests/fixtures/document_properties/,
        custom_errors, custom_link_changes and defaults)."""
        if self.builtin:
            raise error(CATASTROPHIC, _CATASTROPHIC_TEXT)
        if Name is MISSING or LinkToContent is MISSING:
            raise error(450, "Wrong number of arguments or invalid property assignment")
        name, linked = to_text(Name)[:_LONGEST], to_bool(LinkToContent)
        if Type is MISSING:
            raise error(5)
        kind = int(to_integer(Type, "Long"))
        if kind not in _VALUE_TAGS:
            raise error(5)
        if not name:
            raise error(CATASTROPHIC, _CATASTROPHIC_TEXT)
        found = self._found()
        if linked:
            if LinkSource is MISSING:
                raise error(5)
            entry = Custom(name, kind, _linked_default(kind), to_text(LinkSource))
        else:
            if Value is MISSING:
                raise error(5)
            try:
                entry = Custom(name, kind, coerced(Value, kind))
            except _Early:
                raise error(UNSPECIFIED, _UNSPECIFIED_TEXT) from None
        if any(other.name.casefold() == name.casefold() for other in found.custom):
            raise error(UNSPECIFIED, _UNSPECIFIED_TEXT)
        found.custom.append(entry)
        self.book.saved = False
        return DocumentProperty(self.book, -1, entry)


class DocumentProperty(ExcelObject):
    """One property: a built-in one by its place, or a custom one."""

    vba_type_name = "DocumentProperty"

    def __init__(self, book: Workbook, index: int, entry: Custom | None = None) -> None:
        self.book = book
        self.index = index
        self.entry = entry

    def name(self) -> str:
        return BUILTIN[self.index][0] if self.entry is None else self.entry.name

    def _custom(self, name: str) -> Custom:
        """The custom property, or the error a built-in one answers this member with."""
        if self.entry is None:
            raise _failed(name)
        return self.entry

    def _alive(self, entry: Custom, failure: VBARuntimeError | None = None) -> Custom:
        """A custom property still in the workbook, or the error a deleted one answers (custom_stale)."""
        if all(other is not entry for other in state(self.book).custom):
            if failure is None:
                raise VBAUnsupportedError("this member of a deleted custom document property is not implemented")
            raise failure
        return entry

    @member
    def Name(self) -> object:
        if self.entry is not None:
            self._alive(self.entry, _gone("Name"))
        return self.name()

    @setter("Name")
    def _set_name(self, value: object) -> None:
        """A custom property takes a name no other one has; a built-in one keeps its own."""
        if self.entry is not None:
            entry, name = self._alive(self.entry), to_text(value)[:_LONGEST]
            if any(other is not entry and other.name.casefold() == name.casefold() for other in state(self.book).custom):
                raise error(5)
            entry.name = name
        self.book.saved = False

    @member
    def Type(self) -> object:
        if self.entry is None:
            return VBAInt(BUILTIN[self.index][1], "Long")
        return VBAInt(self._alive(self.entry).kind, "Long")

    @setter("Type")
    def _set_type(self, value: object) -> None:
        """A custom property's value turned into the type, or error -2147467259 where it does not turn."""
        entry = self._alive(self._custom("Type"))
        kind = int(to_integer(value, "Long"))
        if kind not in _VALUE_TAGS:
            raise error(5)
        try:
            turned = coerced(shown(entry.kind, entry.value), kind)
        except (VBARuntimeError, _Early):
            raise _failed("Type") from None
        entry.kind, entry.value = kind, turned
        self.book.saved = False

    @member(default=True)
    def Value(self) -> object:
        if self.entry is not None:
            entry = self._alive(self.entry, _failed("Value"))
            return shown(entry.kind, entry.value)
        found = state(self.book).builtin[self.index]
        kind = BUILTIN[self.index][1]
        if found is None:
            if kind == STRING:
                return EMPTY
            raise _failed("Value")
        return shown(kind, found)

    @setter("Value")
    def _set_value(self, value: object) -> None:
        try:
            if self.entry is not None:
                entry = self._alive(self.entry, _gone("Value"))
                if entry.link:
                    raise _failed("Value")
                entry.value = coerced(value, entry.kind)
            else:
                state(self.book).builtin[self.index] = coerced(value, BUILTIN[self.index][1])
        except _Early:
            raise _failed("Value") from None
        self.book.saved = False

    @member
    def LinkToContent(self) -> object:
        return self.entry is not None and bool(self._alive(self.entry).link)

    @setter("LinkToContent")
    def _set_link_to_content(self, value: object) -> None:
        """Ignored on a custom property, linked or not (custom_link_changes)."""
        self._alive(self._custom("LinkToContent"))
        self.book.saved = False

    @member
    def LinkSource(self) -> object:
        return self._alive(self._custom("LinkSource")).link

    @setter("LinkSource")
    def _set_link_source(self, value: object) -> None:
        """A linked property takes another range's name; one that is not linked stays as it is."""
        entry = self._alive(self._custom("LinkSource"))
        if entry.link:
            entry.link = to_text(value)
        self.book.saved = False

    @method
    def Delete(self) -> object:
        entry = self._alive(self._custom("Delete"), _gone("Delete"))
        found = state(self.book)
        found.custom = [other for other in found.custom if other is not entry]
        self.book.saved = False
        return EMPTY

    @member
    def Parent(self) -> object:
        return self.book

    @member
    def Application(self) -> object:
        return self.book.application

    @member
    def Creator(self) -> object:
        return VBAInt(1480803660, "Long")


# --- the save --------------------------------------------------------------------------------------------------


def _cells(node: Node, tables: set[str]) -> bool:
    """Whether a name's formula stands for cells of one sheet of the workbook: references and their unions,
    intersections and spans, and tables (names_more, names_tables)."""
    if isinstance(node, Reference):
        return ":" not in node.sheet and "[" not in node.sheet
    if isinstance(node, Binary):
        return (node.op in (":", " ", ",") and node.left is not None and node.right is not None
                and _cells(node.left, tables) and _cells(node.right, tables))
    if isinstance(node, NameNode):
        return not node.sheet and node.name.casefold() in tables
    return isinstance(node, Structured)


def _named_range(text: str, tables: set[str]) -> bool:
    body = text.removeprefix("=")
    for formula in (body, f"({body})"):
        try:
            return _cells(parse(formula), tables)
        except FormulaError:
            continue
    return False


def _titles(book: Workbook) -> tuple[list[str], list[str]]:
    """The sheets in tab order, and the visible names that stand for cells in the order the Names collection has
    them, a sheet's own spelt with its sheet. Excel lists chart sheets under a heading of their own after the
    worksheets (chart_sheet), which the model, holding none, never writes."""
    from pyopenvba.apps.excel._editing import name_scope
    from pyopenvba.apps.excel._model import DefinedName

    sheets = [sheet.name for sheet in book.sheets_]
    tables = {table.name.casefold() for sheet in book.sheets_ for table in sheet.tables}
    names: list[str] = []
    for item in book.names_.vba_items():
        assert isinstance(item, DefinedName)
        entry = item.entry
        if entry.visible and _named_range(entry.refers_to, tables):
            scope, bare = name_scope(book, entry), split_sheet(entry.name)[1]
            names.append(quote_sheet(scope) + "!" + bare if scope else bare)
    return sheets, names


#: What Excel escapes in a docProps part: the underscore of a run that reads as an escape, and the characters XML
#: cannot hold but a tab and a line feed.
_ESCAPES: Final = re.compile("_(?=x[0-9A-Fa-f]{4}_)|[\x00-\x08\x0b-\x1f￾￿]")


def _spelled(text: str, *, attribute: bool = False) -> str:
    """Text as Excel spells it in a docProps part (custom_text): what XML cannot hold _xhhhh_, in lower case where
    a cell's text has upper, a run that reads as an escape with its underscore escaped, a line feed CRLF, and &, <
    and > entities, with " too in an attribute. The two halves of a surrogate pair VBA holds are one character."""
    if any("\ud800" <= character <= "\udfff" for character in text):
        text = text.encode("utf-16-le", "surrogatepass").decode("utf-16-le", "surrogatepass")
        if any("\ud800" <= character <= "\udfff" for character in text):
            raise VBAUnsupportedError("a document property holding half of a surrogate pair is not implemented")
    text = _ESCAPES.sub(lambda found: f"_x{ord(found.group()):04x}_", text)
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return text.replace('"', "&quot;") if attribute else text.replace("\n", "\r\n")


def _element(element: str, value: object, kind: int) -> str:
    if kind == DATE:
        assert isinstance(value, int)
        typed = ' xsi:type="dcterms:W3CDTF"' if element.startswith("dcterms:") else ""
        return f"<{element}{typed}>{_w3c(value)}</{element}>"
    text = str(int(value)) if isinstance(value, int) else _spelled(str(value))
    return f"<{element}>{text}</{element}>"


def _core(found: Properties) -> str:
    body = "".join(_element(element, value, BUILTIN[_AT[element]][1]) for element in _CORE
                   if (value := found.builtin[_AT[element]]) is not None)
    return (_DECLARATION + f'<cp:coreProperties xmlns:cp="{_CORE_NS}" xmlns:dc="http://purl.org/dc/elements/1.1/" '
            'xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' + body + "</cp:coreProperties>")


def _app(book: Workbook, found: Properties) -> str:
    sheets, names = _titles(book)
    headings = [(heading, len(titles)) for heading, titles in (("Worksheets", sheets), ("Named Ranges", names))
                if titles]
    parts: list[str] = []
    for element in _APP:
        if element in _SETTINGS:
            parts.append(f"<{element}>{_SETTINGS[element]}</{element}>")
        elif element == "HeadingPairs":
            items = "".join(f"<vt:variant><vt:lpstr>{heading}</vt:lpstr></vt:variant><vt:variant><vt:i4>{count}"
                            "</vt:i4></vt:variant>" for heading, count in headings)
            parts.append(f'<HeadingPairs><vt:vector size="{2 * len(headings)}" baseType="variant">{items}'
                         "</vt:vector></HeadingPairs>")
        elif element == "TitlesOfParts":
            titles = sheets + names
            items = "".join(f"<vt:lpstr>{_spelled(title)}</vt:lpstr>" for title in titles)
            parts.append(f'<TitlesOfParts><vt:vector size="{len(titles)}" baseType="lpstr">{items}</vt:vector>'
                         "</TitlesOfParts>")
        elif (value := found.builtin[_AT[element]]) is not None:
            parts.append(_element(element, value, BUILTIN[_AT[element]][1]))
    return (_DECLARATION + f'<Properties xmlns="{_APP_NS}" xmlns:vt="{_VT_NS}">' + "".join(parts) + "</Properties>")


def _custom_xml(found: Properties) -> str:
    items: list[str] = []
    for pid, entry in enumerate(found.custom, start=2):
        tag, value = _VALUE_TAGS[entry.kind], entry.value
        if entry.kind == DATE:
            assert isinstance(value, int)
            text = _w3c(value)
        elif entry.kind == BOOLEAN:
            text = "true" if value else "false"
        elif entry.kind == NUMBER:
            text = str(value)
        elif entry.kind == FLOAT:
            assert isinstance(value, float)
            text = number_text(value)
        else:
            text = _spelled(str(value))
        link = f' linkTarget="{_spelled(entry.link, attribute=True)}"' if entry.link else ""
        items.append(f'<property fmtid="{FMTID}" pid="{pid}" name="{_spelled(entry.name, attribute=True)}"{link}>'
                     f"<vt:{tag}>{text}</vt:{tag}></property>")
    return _DECLARATION + f'<Properties xmlns="{_CUSTOM_NS}" xmlns:vt="{_VT_NS}">' + "".join(items) + "</Properties>"


def _resolved(book: Workbook, link: str) -> tuple[int, object] | None:
    """A linked property's type and value from the first cell of the range its name stands for: a date, a Boolean,
    a number as a Float, and anything else as text (custom_link_kinds); None where the name stands for none."""
    found = book.names_.find(link)
    if found is None:
        return None
    try:
        target = found.refers_to_range()
        first = target.areas[0]
        value = Range(target.sheet, [Area(first.top, first.left, first.top, first.left, first.sheet)]).vba_get("Value")
    except (VBARuntimeError, VBAUnsupportedError, ValueError):
        return None
    if isinstance(value, bool):
        return BOOLEAN, value
    if isinstance(value, VBADate):
        return DATE, to_utc(kept_moment(value))
    if isinstance(value, VBAErrorValue):
        return STRING, _ERROR_TEXT.get(value.number, "#VALUE!")
    if value is EMPTY:
        return STRING, ""
    if isinstance(value, str):
        return STRING, value
    try:
        return FLOAT, float(to_number(value))
    except VBARuntimeError:
        return None


def _related(package: OpcFile, part: str) -> None:
    """A docProps part's relationship from the package and its content type, added after those there where
    missing, numbered one past the highest."""
    relationship, content = _PARTS[part]
    rels = package.read("_rels/.rels").decode("utf-8")
    if f'Type="{relationship}"' not in rels:
        taken = [int(number) for number in re.findall(r'\bId="rId(\d+)"', rels)]
        rels = rels.replace("</Relationships>", f'<Relationship Id="rId{max(taken, default=0) + 1}" '
                            f'Type="{relationship}" Target="{part}"/></Relationships>')
        package.write("_rels/.rels", in_office_order(rels).encode("utf-8"))
    types = package.read("[Content_Types].xml").decode("utf-8")
    if f'PartName="/{part}"' not in types:
        types = types.replace("</Types>", f'<Override PartName="/{part}" ContentType="{content}"/></Types>')
        package.write("[Content_Types].xml", types.encode("utf-8"))


def _unrelated(package: OpcFile, part: str) -> None:
    """A docProps part gone, with its relationship and its content type."""
    relationship, _content = _PARTS[part]
    package.remove(part)
    rels = package.read("_rels/.rels").decode("utf-8")
    kept = re.sub(r'<Relationship\b[^>]*\bType="' + re.escape(relationship) + r'"[^>]*/>', "", rels)
    if kept != rels:
        package.write("_rels/.rels", in_office_order(kept).encode("utf-8"))
    types = package.read("[Content_Types].xml").decode("utf-8")
    kept = re.sub(r'<Override PartName="/' + re.escape(part) + r'"[^>]*/>', "", types)
    if kept != types:
        package.write("[Content_Types].xml", kept.encode("utf-8"))


def _now(book: Workbook) -> dt.datetime:
    interpreter = book.application.interpreter
    moment = interpreter.clock() if interpreter is not None else dt.datetime.now()
    return moment.replace(microsecond=0)


def stamp(book: Workbook) -> Properties:
    """The properties a save sets, set: the last author, the save time, a creation date where there is none, and
    Excel's application and security."""
    found = state(book)
    saved = to_utc(_now(book))
    found.builtin[LAST_AUTHOR] = None if found.last_author_removed else book.application.user_name
    found.builtin[SAVED] = saved
    if found.builtin[CREATED] is None:
        found.builtin[CREATED] = saved
    found.builtin[APPLICATION] = "Microsoft Excel"
    found.builtin[SECURITY] = VBAInt(0, "Long")
    return found


#: XlRemoveDocInfoType: Comments 1, RemovePersonalInformation 4, the rest through InlineWebExtensions, and All 99.
REMOVALS: Final = frozenset({1, 4, 5, 6, 7, 8, 10, 11, 12, 13, 14, 15, 16, 18, 19, 20, 21, 22, 23, 99})


def remove_information(book: Workbook, kind: object) -> None:
    """Workbook.RemoveDocumentInformation, as measured (tests/fixtures/remove_document_information.json).

    Document properties (8) and All (99) empty every built-in text property a value was given, drop the custom
    properties, take the last author away and turn RemovePersonalInformation on; RemovePersonalInformation (4) turns
    it on. Comments (1) and defined-name comments (18) were measured on a workbook holding neither and change nothing
    there; a workbook holding comments, and the other kinds, are refused. Anything else is error 5.
    """
    from pyopenvba.interpreter._values import error, to_integer

    wanted = int(to_integer(kind, "Long"))
    if wanted not in REMOVALS:
        raise error(5)
    if wanted in (1, 99) and any(getattr(sheet, "notes", None) for sheet in book.sheets_):
        raise VBAUnsupportedError("Removing the comments of a workbook is not implemented")
    if wanted in (18, 99) and any(entry.comment for entry in book.names_.entries):
        raise VBAUnsupportedError("Removing defined name comments is not implemented")
    if wanted not in (1, 4, 8, 18, 99):
        raise VBAUnsupportedError("This kind of document information removal is not implemented")
    if wanted in (4, 8, 99):
        book.remove_personal = True
    if wanted in (8, 99):
        found = state(book)
        for index, (_name, kind_of, _element) in enumerate(BUILTIN):
            if kind_of == STRING and found.builtin[index] is not None and index != APPLICATION:
                found.builtin[index] = ""
        found.builtin[LAST_AUTHOR] = None
        found.last_author_removed = True
        if not found.opaque:
            found.custom.clear()


def write(book: Workbook, package: OpcFile) -> None:
    """The docProps parts as a save writes them: the properties it stamps, each linked property's range's value,
    and every property in its part (session_after_save, no_docprops)."""
    found = stamp(book)
    if not found.opaque:
        for entry in found.custom:
            resolved = _resolved(book, entry.link) if entry.link else None
            if resolved is not None:
                entry.kind, entry.value = resolved
    written = found
    if book.remove_personal:
        # A workbook that removes personal information is saved with no author, last author, manager or company,
        # though the open workbook still reads them (tests/fixtures/remove_document_information.json).
        written = replace(found, builtin=list(found.builtin))
        for slot in (AUTHOR, LAST_AUTHOR, _AT["Manager"], COMPANY):
            if written.builtin[slot] is not None:
                written.builtin[slot] = ""
    package.write(CORE_PART, _core(written).encode("utf-8"))
    package.write(APP_PART, _app(book, written).encode("utf-8"))
    _related(package, CORE_PART)
    _related(package, APP_PART)
    if found.opaque:
        return
    if found.custom:
        package.write(CUSTOM_PART, _custom_xml(found).encode("utf-8"))
        _related(package, CUSTOM_PART)
    elif package.has(CUSTOM_PART):
        _unrelated(package, CUSTOM_PART)
