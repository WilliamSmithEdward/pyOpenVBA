"""A disconnected ADODB.Recordset, as tests/fixtures/adodb_recordset.json measured it in live Excel: a table in memory, made by
appending fields. It reaches nothing outside the process, so it needs no leave from the sandbox; opening one on a source is refused.

Rows belong to a table that clones share, and each recordset has its own place in them, its own Sort and its own Filter. A row
added since the last UpdateBatch stays pending, which is why EditMode goes on saying adEditAdd after Update, and CancelUpdate takes
the row out however long ago it was updated. Filter's grammar is ADO's own and narrow: no IN, a wildcard only at the end or at both
ends, and no AND over an OR.
"""
from __future__ import annotations

import math
import re
from collections.abc import Callable, Iterator
from dataclasses import dataclass, field
from decimal import ROUND_DOWN, Decimal

from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter._objects import VBAObject, member, method, setter
from pyopenvba.interpreter._values import (
    EMPTY,
    MISSING,
    NOTHING,
    NULL,
    VBAArray,
    VBACurrency,
    VBADecimal,
    VBAInt,
    VBASingle,
    error,
    to_bool,
    to_date,
    to_integer,
    to_number,
    to_text,
)

_CLOSED = 3704
_NO_RECORD = 3021
_BAD_ARGUMENTS = 3001
_NO_FIELD = 3265
_OPEN_ALREADY = 3705
_NOT_ALLOWED_OPEN = 3219
_NO_FIELDS = 3709
_DUPLICATE_FIELD = 3367
_ROW_DELETED = -2147217885
_BAD_VALUE = -2147217887

_FIXED_TEXT = {129, 130}
_TEXT = {129, 130, 200, 201, 202, 203}
_SIZES = {2: 2, 3: 4, 4: 4, 5: 8, 6: 8, 7: 8, 11: 2, 12: 24, 14: 16, 16: 1, 17: 1, 18: 2, 19: 4, 20: 8, 21: 8, 64: 8, 72: 16, 131: 19, 133: 6, 135: 16}
_VARIABLE = {128, 204, 205}
#: Types whose values go in and then cannot be read as VBA has no such type.
_UNREADABLE = {16, 18, 19}
#: What Supports answers, as measured; the other options are not known.
_SUPPORTED = {16778240: True, 512: True, 8192: True, 16779264: True, 524288: True, 16809984: True, 65536: True, 131072: True, 256: True, 16384: True,
              262144: True, 1048576: False, 2097152: False}


@dataclass
class _Field:
    name: str
    kind: int
    size: int
    attributes: int


@dataclass(eq=False)
class _Row:
    number: int
    values: list[object]
    #: As of the last UpdateBatch; None for a row added since.
    original: list[object] | None = None
    modified: bool = False
    deleted: bool = False


@dataclass
class _Table:
    fields: list[_Field] = field(default_factory=lambda: [])
    rows: list[_Row] = field(default_factory=lambda: [])
    made: int = 0

    def where(self, name: object) -> int:
        """A field by name or from 0."""
        if isinstance(name, str):
            wanted = name.lower()
            for place, one in enumerate(self.fields):
                if one.name.lower() == wanted:
                    return place
            raise error(_NO_FIELD)
        place = int(to_integer(name, "Long"))
        if not 0 <= place < len(self.fields):
            raise error(_NO_FIELD)
        return place


def _stored(column: _Field, value: object, *, initial: bool = False) -> object:
    """A value as its field keeps it; what will not go is one error."""
    if value is NULL or value is EMPTY:
        if not initial and not column.attributes & 32:
            raise error(_BAD_VALUE)
        return NULL
    try:
        return _converted(column, value)
    except VBARuntimeError as failure:
        raise error(_BAD_VALUE) from failure


def _converted(column: _Field, value: object) -> object:
    kind = column.kind
    if kind in _TEXT:
        text = to_text(value)
        if len(text) > column.size:
            raise error(_BAD_VALUE)
        return text.ljust(column.size) if kind in _FIXED_TEXT else text
    if kind == 72:
        return to_text(value)
    if kind == 2:
        return to_integer(value, "Integer")
    if kind == 3:
        return to_integer(value, "Long")
    if kind == 17:
        return to_integer(value, "Byte")
    if kind == 4:
        return VBASingle(float(to_number(value)))
    if kind == 5:
        return float(to_number(value))
    if kind == 6:
        return VBACurrency(to_number(value))
    if kind in (7, 64, 133, 135):
        return to_date(value)
    if kind == 11:
        return to_bool(value)
    if kind == 12:
        return value
    if kind in (14, 20, 21):
        return VBADecimal(Decimal(str(to_number(value))).quantize(Decimal(1), rounding=ROUND_DOWN))
    if kind in _UNREADABLE:
        return to_integer(value, "Long")
    raise error(_BAD_VALUE)


# --- Filter and Find ------------------------------------------------------------------------------------------------

_CLAUSE_TOKEN = re.compile(r"""\s*(?:(?P<op><>|<=|>=|=|<|>)|(?P<paren>[()])|(?P<single>'(?:[^']|'')*')|(?P<double>"[^"]*")|(?P<date>\#[^#]*\#)
                           |(?P<number>-?\d+(?:\.\d*)?|-?\.\d+)|(?P<word>[^\s()=<>'"#]+))""", re.X)
_Test = Callable[[_Row], bool]


def _clause_tokens(text: str) -> list[tuple[str, str]]:
    found: list[tuple[str, str]] = []
    at, end = 0, len(text.rstrip())
    while at < end:
        match = _CLAUSE_TOKEN.match(text, at)
        if match is None:
            raise error(_BAD_ARGUMENTS)
        kind = match.lastgroup or ""
        found.append((kind, match.group(kind)))
        at = match.end()
    return found


def _compared(left: object, operator: str, right: object) -> bool:
    if left is NULL:
        return False
    if isinstance(right, str):
        one: object = to_text(left).rstrip().lower()
        other: object = right.lower()
    else:
        assert isinstance(right, float)
        one, other = _numeric(left), right
        if one is None:
            return False
    assert isinstance(one, (str, float)) and isinstance(other, (str, float))
    if operator == "=":
        return one == other
    if operator == "<>":
        return one != other
    assert type(one) is type(other)
    if operator == "<":
        return one < other  # type: ignore[operator]
    if operator == "<=":
        return one <= other  # type: ignore[operator]
    if operator == ">":
        return one > other  # type: ignore[operator]
    return one >= other  # type: ignore[operator]


def _numeric(value: object) -> float | None:
    try:
        return float(to_number(value))
    except VBARuntimeError:
        return None


def _like(pattern: str) -> Callable[[str], bool]:
    """LIKE: a wildcard at the end, or at both ends; one only at the start fails."""
    wild = "*%"
    leading, trailing = pattern[:1] in wild and pattern != "", pattern[-1:] in wild and pattern != ""
    core = pattern.strip(wild).lower()
    if any(char in core for char in wild) or (leading and not trailing):
        raise error(_BAD_ARGUMENTS)
    if leading and trailing:
        return lambda text: core in text
    return (lambda text: text.startswith(core)) if trailing else (lambda text: text == core)


class _CriteriaParser:
    def __init__(self, table: _Table, text: str) -> None:
        self.table = table
        self.tokens = _clause_tokens(text)
        self.at = 0
        self.clauses = 0

    def peek(self) -> tuple[str, str]:
        return self.tokens[self.at] if self.at < len(self.tokens) else ("", "")

    def take(self) -> tuple[str, str]:
        found = self.peek()
        if not found[0]:
            raise error(_BAD_ARGUMENTS)
        self.at += 1
        return found

    def whole(self) -> _Test:
        found, _ = self.joined()
        if self.at != len(self.tokens):
            raise error(_BAD_ARGUMENTS)
        return found

    def joined(self) -> tuple[_Test, str]:
        """Groups joined by one of AND and OR, never both, and no AND over a group that holds an OR."""
        first, inner = self.group()
        parts, joiner = [first], ""
        holds_or = inner == "or"
        while self.peek()[0] == "word" and self.peek()[1].lower() in ("and", "or"):
            word = self.take()[1].lower()
            if joiner and word != joiner:
                raise error(_BAD_ARGUMENTS)
            joiner = word
            part, inner = self.group()
            holds_or = holds_or or inner == "or"
            parts.append(part)
        if joiner == "and" and holds_or:
            raise error(_BAD_ARGUMENTS)
        if joiner == "or":
            return (lambda row: any(part(row) for part in parts)), "or"
        return (lambda row: all(part(row) for part in parts)), (joiner or inner)

    def group(self) -> tuple[_Test, str]:
        if self.peek() == ("paren", "("):
            self.take()
            found = self.joined()
            if self.take() != ("paren", ")"):
                raise error(_BAD_ARGUMENTS)
            return found
        return self.clause(), ""

    def clause(self) -> _Test:
        self.clauses += 1
        kind, name = self.take()
        if kind != "word":
            raise error(_BAD_ARGUMENTS)
        place = self.table.where(name.strip("[]"))
        kind, operator = self.take()
        if kind == "word" and operator.lower() == "like":
            value_kind, literal = self.take()
            if value_kind != "single":
                raise error(_BAD_ARGUMENTS)
            matches = _like(literal[1:-1].replace("''", "'"))
            return lambda row: row.values[place] is not NULL and matches(to_text(row.values[place]).rstrip().lower())
        if kind != "op":
            raise error(_BAD_ARGUMENTS)
        value = self._value()
        return lambda row: _compared(row.values[place], operator, value)

    def _value(self) -> object:
        kind, text = self.take()
        if kind == "single":
            literal = text[1:-1].replace("''", "'")
            number = _numeric(literal)
            return number if number is not None and _all_numeric(self.table, self.tokens, self.at) else literal
        if kind == "number":
            return float(text)
        if kind == "double":
            return text  # Kept with its quotes, so it matches nothing, as measured.
        if kind == "date":
            return float(to_date(text[1:-1]).serial)
        raise error(_BAD_ARGUMENTS)


def _all_numeric(table: _Table, tokens: list[tuple[str, str]], after: int) -> bool:
    """Whether the clause just read is about a number field, so that '1' compares as 1."""
    name = tokens[after - 3][1].strip("[]")
    return table.fields[table.where(name)].kind not in _TEXT


def _criteria(table: _Table, text: str, *, single: bool = False) -> _Test:
    parser = _CriteriaParser(table, text)
    found = parser.whole()
    if single and parser.clauses != 1:
        raise error(_BAD_ARGUMENTS)
    return found


def _sort_key(value: object) -> tuple[int, float, str]:
    """Null first, then numbers, then text without regard to case."""
    if value is NULL:
        return (0, 0.0, "")
    if isinstance(value, str):
        return (2, 0.0, value.lower())
    number = _numeric(value)
    return (1, 0.0 if number is None else number, "")


# --- the objects ----------------------------------------------------------------------------------------------------


class AdoField(VBAObject):
    vba_type_name = "Field"

    def __init__(self, owner: AdoRecordset, place: int) -> None:
        self.owner = owner
        self.place = place

    @property
    def column(self) -> _Field:
        return self.owner.table.fields[self.place]

    def vba_information_value(self) -> object:
        """What IsNull sees: Null only in a field that may hold it, though one that may not starts with nothing in it
        (tests/fixtures/adodb_recordset.json)."""
        value = self.Value()
        return EMPTY if value is NULL and not self.column.attributes & 32 else value

    @member
    def Name(self) -> object:
        return self.column.name

    @member
    def Type(self) -> object:
        return VBAInt(self.column.kind, "Long")

    @member
    def DefinedSize(self) -> object:
        return VBAInt(self.owner.shown_sizes.get(id(self.column), self.column.size), "Long")

    @member
    def Attributes(self) -> object:
        return VBAInt(self.column.attributes, "Long")

    @member
    def Precision(self) -> object:
        return VBAInt(0, "Byte")

    @member
    def NumericScale(self) -> object:
        return VBAInt(0, "Byte")

    @member
    def ActualSize(self) -> object:
        value = self.owner.row_to_read().values[self.place]
        if value is NULL:
            return VBAInt(-1, "Long")
        return VBAInt(len(value) if isinstance(value, str) else self.column.size, "Long")

    @member(default=True)
    def Value(self) -> object:
        value = self.owner.row_to_read().values[self.place]
        if self.column.kind in _UNREADABLE and value is not NULL:
            raise error(458, "Variable uses an Automation type not supported in Visual Basic")
        return value

    @setter("Value")
    def _set_value(self, value: object) -> None:
        self.owner.write(self.place, value)

    @member
    def OriginalValue(self) -> object:
        row = self.owner.row_to_read()
        return NULL if row.original is None else row.original[self.place]

    @member
    def UnderlyingValue(self) -> object:
        return self.OriginalValue()


class AdoFields(VBAObject):
    vba_type_name = "Fields"

    def __init__(self, owner: AdoRecordset) -> None:
        self.owner = owner

    @member
    def Count(self) -> object:
        return VBAInt(len(self.owner.table.fields), "Long")

    @member(default=True)
    def Item(self, Index: object = MISSING) -> object:
        if Index is MISSING:
            raise error(449)
        return AdoField(self.owner, self.owner.table.where(Index))

    def vba_iterate(self) -> Iterator[object]:
        return iter([AdoField(self.owner, place) for place in range(len(self.owner.table.fields))])

    @method
    def Append(self, Name: object = MISSING, Type: object = MISSING, DefinedSize: object = MISSING, Attrib: object = MISSING,
               FieldValue: object = MISSING) -> object:
        owner = self.owner
        if owner.is_open:
            raise error(_NOT_ALLOWED_OPEN)
        name = to_text(Name)
        kind = int(to_integer(Type, "Long"))
        attributes = 4 if Attrib is MISSING else int(to_integer(Attrib, "Long")) | 4
        if kind == 134:
            raise VBAUnsupportedError("a field of type adDBTime is not implemented")
        if kind == 8:
            # adBSTR is kept as adVarWChar of no limit (tests/fixtures/adodb_recordset.json).
            kind, size = 203, 2147483647
            shown = -1
        elif kind in _TEXT or kind in _VARIABLE:
            if DefinedSize is MISSING:
                raise error(_BAD_ARGUMENTS)
            size = shown = int(to_integer(DefinedSize, "Long"))
        elif kind in _SIZES:
            size = shown = _SIZES[kind]
            attributes |= 16
        else:
            raise error(_BAD_ARGUMENTS)
        if any(one.name.lower() == name.lower() for one in owner.table.fields):
            raise error(_DUPLICATE_FIELD)
        made = _Field(name, kind, size, attributes)
        owner.table.fields.append(made)
        owner.shown_sizes[id(made)] = shown
        return EMPTY

    @method
    def Delete(self, Index: object = MISSING) -> object:
        if self.owner.is_open:
            raise error(_NOT_ALLOWED_OPEN)
        del self.owner.table.fields[self.owner.table.where(Index)]
        return EMPTY

    @method
    def Refresh(self) -> object:
        return EMPTY


class AdoRecordset(VBAObject):
    vba_type_name = "Recordset"

    def __init__(self) -> None:
        self.table = _Table()
        self.is_open = False
        self.shown_sizes: dict[int, int] = {}
        self.cursor_type, self.lock_type, self.cursor_location = 0, 1, 2
        self.at = "none"  # row, bof, eof, none, or gone once its row is deleted
        self.place = 0
        self.row: _Row | None = None
        self.sort_text = ""
        self.sort_order: list[tuple[int, bool]] = []
        self.filter_text: str | None = None
        self.filter_test: _Test | None = None
        self.kept: list[_Row] = []

    # --- where it is

    def _opened(self) -> None:
        if not self.is_open:
            raise error(_CLOSED)

    def _matching(self) -> list[_Row]:
        rows = [row for row in self.table.rows if not row.deleted and (self.filter_test is None or self.filter_test(row))]
        for place, descending in reversed(self.sort_order):
            rows.sort(key=lambda row: _sort_key(row.values[place]), reverse=descending)
        return rows

    def view(self) -> list[_Row]:
        """The rows it walks: those the Filter lets through, and any it added since, which stay until the Filter is set again."""
        rows = self._matching()
        return rows + [row for row in self.kept if row not in rows and not row.deleted and row in self.table.rows]

    def _land(self, rows: list[_Row], place: int) -> None:
        if place < 0:
            self.at, self.row, self.place = "bof", None, -1
        elif place >= len(rows):
            self.at, self.row, self.place = "eof", None, len(rows)
        else:
            self.at, self.row, self.place = "row", rows[place], place

    def _first(self) -> None:
        rows = self.view()
        if rows:
            self._land(rows, 0)
        else:
            self.at, self.row, self.place = "none", None, 0

    def _place_now(self, rows: list[_Row]) -> int:
        if self.at == "row" and self.row is not None and self.row in rows:
            return rows.index(self.row)
        return self.place

    def row_to_read(self) -> _Row:
        self._opened_for_field()
        if self.at == "gone":
            raise error(_ROW_DELETED)
        if self.at != "row" or self.row is None:
            raise error(_NO_RECORD)
        return self.row

    def _opened_for_field(self) -> None:
        if not self.is_open:
            raise error(_NO_RECORD)

    def write(self, place: int, value: object) -> None:
        row = self.row_to_read()
        row.values[place] = _stored(self.table.fields[place], value)
        row.modified = True

    # --- opening and closing

    @member(default=True)
    def Fields(self, Index: object = MISSING) -> object:
        fields = AdoFields(self)
        return fields if Index is MISSING else fields.Item(Index)

    @member
    def Item(self, Index: object = MISSING) -> object:
        """Where the interpreter takes r("N") = value; ADO itself has no such member."""
        return AdoFields(self).Item(Index)

    @method
    def Open(self, Source: object = MISSING, ActiveConnection: object = MISSING, CursorType: object = MISSING, LockType: object = MISSING,
             Options: object = MISSING) -> object:
        if Source is not MISSING or ActiveConnection is not MISSING:
            raise VBAUnsupportedError("a Recordset opened on a source is not implemented: only one made by appending fields")
        if self.is_open:
            raise error(_OPEN_ALREADY)
        if not self.table.fields:
            raise error(_NO_FIELDS)
        self.is_open = True
        self.cursor_type, self.lock_type, self.cursor_location = 3, 4, 3
        self.at, self.row, self.place = "none", None, 0
        return EMPTY

    @method
    def Close(self) -> object:
        self._opened()
        self.is_open = False
        self.table = _Table()
        self.at, self.row, self.place = "none", None, 0
        self.sort_text, self.sort_order, self.filter_text, self.filter_test, self.kept = "", [], None, None, []
        return EMPTY

    @method
    def Clone(self, LockType: object = MISSING) -> object:
        self._opened()
        made = AdoRecordset()
        made.table, made.is_open, made.shown_sizes = self.table, True, self.shown_sizes
        made.cursor_type, made.lock_type, made.cursor_location = self.cursor_type, self.lock_type, self.cursor_location
        made._first()
        return made

    @method
    def Requery(self, Options: object = MISSING) -> object:
        raise error(_NOT_ALLOWED_OPEN)

    @method
    def Supports(self, CursorOptions: object = MISSING) -> object:
        wanted = int(to_integer(CursorOptions, "Long"))
        if wanted not in _SUPPORTED:
            raise VBAUnsupportedError(f"Recordset.Supports({wanted}) is not implemented")
        return _SUPPORTED[wanted]

    # --- properties

    @member
    def State(self) -> object:
        return VBAInt(1 if self.is_open else 0, "Long")

    @member
    def CursorType(self) -> object:
        return VBAInt(self.cursor_type, "Long")

    @setter("CursorType")
    def _set_cursor_type(self, value: object) -> None:
        self.cursor_type = int(to_integer(value, "Long"))

    @member
    def LockType(self) -> object:
        return VBAInt(self.lock_type, "Long")

    @setter("LockType")
    def _set_lock_type(self, value: object) -> None:
        self.lock_type = int(to_integer(value, "Long"))

    @member
    def CursorLocation(self) -> object:
        return VBAInt(self.cursor_location, "Long")

    @setter("CursorLocation")
    def _set_cursor_location(self, value: object) -> None:
        self.cursor_location = int(to_integer(value, "Long"))

    @member
    def ActiveConnection(self) -> object:
        return NOTHING

    @member
    def Source(self) -> object:
        return ""

    @member
    def PageSize(self) -> object:
        return VBAInt(10, "Long")

    @member
    def PageCount(self) -> object:
        return VBAInt(math.ceil(len(self._matching()) / 10), "Long")

    @member
    def MaxRecords(self) -> object:
        return VBAInt(0, "Long")

    @member
    def CacheSize(self) -> object:
        return VBAInt(1, "Long")

    @member
    def EOF(self) -> object:
        self._opened()
        return self.at in ("eof", "none")

    @member
    def BOF(self) -> object:
        self._opened()
        return self.at in ("bof", "none")

    @member
    def RecordCount(self) -> object:
        self._opened()
        return VBAInt(len(self._matching()), "Long")

    @member
    def AbsolutePosition(self) -> object:
        self._opened()
        if self.at == "row":
            return VBAInt(self._place_now(self.view()) + 1, "Long")
        return VBAInt({"bof": -2, "eof": -3}.get(self.at, -1), "Long")

    @setter("AbsolutePosition")
    def _set_absolute_position(self, value: object) -> None:
        self._opened()
        self._land(self.view(), int(to_integer(value, "Long")) - 1)

    @member
    def Bookmark(self) -> object:
        return float(self.row_to_read().number)

    @setter("Bookmark")
    def _set_bookmark(self, value: object) -> None:
        self._opened()
        rows = self.view()
        wanted = int(to_integer(value, "Long"))
        for place, row in enumerate(rows):
            if row.number == wanted:
                self._land(rows, place)
                return
        raise error(_BAD_ARGUMENTS)

    @member
    def EditMode(self) -> object:
        self._opened()
        if self.at == "gone":
            return VBAInt(0, "Long")
        if self.at != "row" or self.row is None:
            raise error(_NO_RECORD)
        return VBAInt(2 if self.row.original is None else 1 if self.row.modified else 0, "Long")

    @member
    def Status(self) -> object:
        row = self.row_to_read()
        return VBAInt(1 if row.original is None else 2 if row.modified else 0, "Long")

    @member
    def Collect(self, Index: object = MISSING) -> object:
        return self.row_to_read().values[self.table.where(Index)]

    # --- moving

    @method
    def MoveFirst(self) -> object:
        self._opened()
        rows = self.view()
        if not rows:
            raise error(_NO_RECORD)
        self._land(rows, 0)
        return EMPTY

    @method
    def MoveLast(self) -> object:
        self._opened()
        rows = self.view()
        if not rows:
            raise error(_NO_RECORD)
        self._land(rows, len(rows) - 1)
        return EMPTY

    @method
    def MoveNext(self) -> object:
        self._opened()
        if self.at in ("eof", "none"):
            raise error(_NO_RECORD)
        rows = self.view()
        if self.at == "gone":
            self._land(rows, self.place)
        else:
            self._land(rows, self._place_now(rows) + 1)
        return EMPTY

    @method
    def MovePrevious(self) -> object:
        self._opened()
        if self.at in ("bof", "none"):
            raise error(_NO_RECORD)
        rows = self.view()
        self._land(rows, self._place_now(rows) - 1)
        return EMPTY

    def _anchor(self, rows: list[_Row], start: object) -> int:
        if start is MISSING:
            return self._place_now(rows)
        wanted = int(to_integer(start, "Long"))
        if wanted == 0:
            return self._place_now(rows)
        if wanted == 1:
            return 0
        if wanted == 2:
            return len(rows) - 1
        for place, row in enumerate(rows):
            if row.number == wanted:
                return place
        raise error(_BAD_ARGUMENTS)

    @method
    def Move(self, NumRecords: object = MISSING, Start: object = MISSING) -> object:
        self._opened()
        count = int(to_integer(NumRecords, "Long"))
        rows = self.view()
        if not rows or self.at == "none":
            raise error(_NO_RECORD)
        if count == 0 and Start is MISSING:
            return EMPTY
        self._land(rows, self._anchor(rows, Start) + count)
        return EMPTY

    # --- changing

    def _set_many(self, row: _Row, names: object, values: object) -> None:
        if names is MISSING:
            return
        if isinstance(names, VBAArray):
            pairs = zip(names.elements(), values.elements() if isinstance(values, VBAArray) else [values], strict=False)
        else:
            pairs = zip([names], [values], strict=False)
        for name, value in pairs:
            place = self.table.where(name)
            row.values[place] = _stored(self.table.fields[place], value)

    @method
    def AddNew(self, FieldList: object = MISSING, Values: object = MISSING) -> object:
        self._opened()
        self.table.made += 1
        row = _Row(self.table.made, [_stored(one, NULL, initial=True) for one in self.table.fields])
        self._set_many(row, FieldList, Values)
        self.table.rows.append(row)
        if self.filter_test is not None:
            self.kept.append(row)
        rows = self.view()
        self.at, self.row, self.place = "row", row, rows.index(row)
        return EMPTY

    def _current_to_change(self) -> _Row:
        self._opened()
        if self.at == "gone":
            raise error(_ROW_DELETED)
        if self.at != "row" or self.row is None:
            raise error(_NO_RECORD)
        return self.row

    @method
    def Update(self, Fields: object = MISSING, Values: object = MISSING) -> object:
        row = self._current_to_change()
        if Fields is not MISSING:
            self._set_many(row, Fields, Values)
            row.modified = True
        return EMPTY

    @method
    def CancelUpdate(self) -> object:
        row = self._current_to_change()
        if row.original is None:
            self.table.rows.remove(row)
            self.at, self.row, self.place = "none", None, 0
        else:
            row.values, row.modified = list(row.original), False
        return EMPTY

    @method
    def Delete(self, AffectRecords: object = MISSING) -> object:
        row = self._current_to_change()
        self.place = self._place_now(self.view())
        if row.original is None:
            self.table.rows.remove(row)
        else:
            row.deleted = True
        self.at, self.row = "gone", None
        return EMPTY

    @method
    def UpdateBatch(self, AffectRecords: object = MISSING) -> object:
        self._opened()
        self.table.rows = [row for row in self.table.rows if not row.deleted]
        for row in self.table.rows:
            row.original, row.modified = list(row.values), False
        return EMPTY

    @method
    def CancelBatch(self, AffectRecords: object = MISSING) -> object:
        self._opened()
        kept: list[_Row] = []
        for row in self.table.rows:
            if row.original is None:
                continue
            row.values, row.modified, row.deleted = list(row.original), False, False
            kept.append(row)
        self.table.rows = kept
        if self.row is not None and self.row not in kept:
            self.at, self.row = "gone", None
        return EMPTY

    # --- Sort, Filter, Find

    @member
    def Sort(self) -> object:
        return self.sort_text

    @setter("Sort")
    def _set_sort(self, value: object) -> None:
        text = to_text(value)
        order: list[tuple[int, bool]] = []
        for part in filter(None, (piece.strip() for piece in text.split(","))):
            words = part.split()
            descending = len(words) > 1 and words[-1].upper() == "DESC"
            if len(words) > 1 and words[-1].upper() in ("ASC", "DESC"):
                words = words[:-1]
            order.append((self.table.where(" ".join(words).strip("[]")), descending))
        self.sort_text, self.sort_order = text, order
        self._first()

    @member
    def Filter(self) -> object:
        return VBAInt(0, "Long") if self.filter_text is None else self.filter_text

    @setter("Filter")
    def _set_filter(self, value: object) -> None:
        if isinstance(value, str):
            text: str | None = value or None
            test = _criteria(self.table, value) if value else None
        elif int(to_integer(value, "Long")) == 0:
            text, test = None, None
        else:
            raise VBAUnsupportedError("Recordset.Filter by a FilterGroupEnum other than adFilterNone is not implemented")
        self.filter_text, self.filter_test, self.kept = text, test, []
        self._first()

    @method
    def Find(self, Criteria: object = MISSING, SkipRows: object = MISSING, SearchDirection: object = MISSING, Start: object = MISSING) -> object:
        self._opened()
        test = _criteria(self.table, to_text(Criteria), single=True)
        rows = self.view()
        step = -1 if SearchDirection is not MISSING and int(to_integer(SearchDirection, "Long")) < 0 else 1
        skip = 0 if SkipRows is MISSING else int(to_integer(SkipRows, "Long"))
        if not rows or self.at == "none":
            self.at, self.row = "none", None
            return EMPTY
        place = self._anchor(rows, Start) + skip * step
        while 0 <= place < len(rows):
            if test(rows[place]):
                self._land(rows, place)
                return EMPTY
            place += step
        self._land(rows, -1 if step < 0 else len(rows))
        return EMPTY

    # --- whole rows

    def take(self, count: int | None) -> list[_Row]:
        """Up to ``count`` rows from the current one, which it moves past."""
        rows = self.view()
        start = self._place_now(rows)
        end = len(rows) if count is None else min(start + count, len(rows))
        taken = rows[start:end]
        self._land(rows, end)
        return taken

    @method
    def GetRows(self, Rows: object = MISSING, Start: object = MISSING, Fields: object = MISSING) -> object:
        self._opened()
        count = None if Rows is MISSING or int(to_integer(Rows, "Long")) < 0 else int(to_integer(Rows, "Long"))
        if count == 0:
            raise error(_NOT_ALLOWED_OPEN)
        if isinstance(Fields, VBAArray):
            places = [self.table.where(one) for one in Fields.elements()]
        elif Fields is MISSING:
            places = list(range(len(self.table.fields)))
        else:
            places = [self.table.where(Fields)]
        rows = self.view()
        if Start is not MISSING and rows:
            self._land(rows, self._anchor(rows, Start))
        if self.at != "row":
            raise error(_ROW_DELETED if self.at == "gone" else _NO_RECORD)
        taken = self.take(count)
        made = VBAArray([(0, len(places) - 1), (0, len(taken) - 1)])
        for across, place in enumerate(places):
            for down, row in enumerate(taken):
                made.set([across, down], row.values[place])
        return made

    @method
    def GetString(self, StringFormat: object = MISSING, NumRows: object = MISSING, ColumnDelimeter: object = MISSING, RowDelimeter: object = MISSING,
                  NullExpr: object = MISSING) -> object:
        self._opened()
        if StringFormat is not MISSING and int(to_integer(StringFormat, "Long")) != 2:
            raise error(_BAD_ARGUMENTS)
        if self.at != "row":
            raise error(_ROW_DELETED if self.at == "gone" else _NO_RECORD)
        column = "\t" if ColumnDelimeter is MISSING else to_text(ColumnDelimeter)
        line = "\r" if RowDelimeter is MISSING else to_text(RowDelimeter)
        null = "" if NullExpr is MISSING else to_text(NullExpr)
        count = None if NumRows is MISSING or int(to_integer(NumRows, "Long")) < 0 else int(to_integer(NumRows, "Long"))
        return "".join(column.join(null if value is NULL else to_text(value) for value in row.values) + line for row in self.take(count))

