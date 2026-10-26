"""Workbooks.Open of CSV and text files, as Excel reads them on this model's US English system.

Measured in live Excel (scripts/measure_text_files.py, tests/fixtures/text_files.json):

- A file whose extension is not a workbook's is text. A .csv file is a
  CSV: split at commas, or at the one character a first line of sep=
  names, that line dropped, none naming a comma; FileFormat xlCSV, or
  xlCSVUTF8 after a UTF-8 byte order mark. A .csv file of three bytes or
  fewer, one in UTF-16, or one whose sep= line names more than a
  character reads as any other text file.
- Any other text file splits at what Workbooks.Open's Format last named
  in the session, whatever the file that open read, a tab until one
  names something: 1 a tab, 2 a comma, 3 a space, 4 a semicolon, 5
  nothing, which turns each tab into spaces up to the next eighth column,
  and 6 Delimiter's first character. Its FileFormat is
  xlCurrentPlatformText, or xlCSV where Format 2 is given; xlUnicodeText
  in UTF-16 and 65 in UTF-8.
- Bytes are Windows-1252 unless a byte order mark says UTF-8 or UTF-16
  little-endian. A file holding a NUL is not read here: Excel refuses
  some such files and drops the NULs of others.
- Lines end at CR LF, CR or LF. A field that begins with a double quote
  runs to the next quote not doubled, line ends and delimiters in it, and
  what follows that quote up to the delimiter is added as it stands; a
  quote anywhere else is a character. A field holds 32,759 characters at
  most: past them Excel drops the next character and reads the rest as a
  new line.
- Each field is typed as TextToColumns types a General column's (see
  _text_columns.typed_field); one that begins with = and has more is a
  formula, which leaves the workbook unsaved.
- The workbook is named after the file, its brackets made parentheses,
  and has one sheet, named after the file without its extension, cut to
  31 characters.
"""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import MISSING, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Application, Workbook

CSV, CSV_UTF8, UNICODE_TEXT, UTF8_TEXT, TEXT = 6, 62, 42, 65, -4158
#: The FileFormats of a workbook read from text, which the model does not save.
TEXT_FORMATS = frozenset({CSV, CSV_UTF8, UNICODE_TEXT, UTF8_TEXT, TEXT})
#: The extensions of workbooks; any other file opens as text.
WORKBOOK_EXTENSIONS = frozenset({".xlsx", ".xlsm", ".xltx", ".xltm", ".xlam", ".xlsb", ".xls", ".xlt", ".xla"})
#: What each Format of Workbooks.Open splits a text file at; 6 takes Delimiter's first character.
_FORMATS: dict[int, str | None] = {1: "\t", 2: ",", 3: " ", 4: ";", 5: None}
_UTF8_BOM, _UTF16_BOM = b"\xef\xbb\xbf", b"\xff\xfe"
#: A .csv file shorter than this reads as text.
_SHORTEST_CSV = 4
#: The most characters a field holds.
_LONGEST_FIELD = 32759
_TAB_STOP = 8
_LONGEST_SHEET_NAME = 31
#: Windows-1252's characters for the bytes 0x80 to 0x9F; the five it leaves out read as the control characters.
_WINDOWS_1252 = {code: bytes([code]).decode("cp1252") for code in range(0x80, 0xA0)
                 if code not in (0x81, 0x8D, 0x8F, 0x90, 0x9D)}


def is_text(path: Path) -> bool:
    """Whether a file opens as text rather than as a workbook: its extension is not a workbook's."""
    return path.suffix.lower() not in WORKBOOK_EXTENSIONS


def remember_format(application: Application, format_: object, delimiter: object) -> int | None:
    """Workbooks.Open's Format, which sets what later text files split at, whatever file this open reads; its
    number, or None where it is left out."""
    if format_ is MISSING:
        return None
    code = int(to_integer(format_, "Long"))
    if code == 6:
        named = "" if delimiter is MISSING else to_text(delimiter)
        if not named:
            raise VBAUnsupportedError("Workbooks.Open with Format 6 and no Delimiter is not implemented")
        application.text_delimiter = named[0]
    elif code in _FORMATS:
        application.text_delimiter = _FORMATS[code]
    else:
        raise VBAUnsupportedError(f"Workbooks.Open with Format {code} is not implemented: Format is 1 to 6")
    return code


def open_text(application: Application, path: Path, format_code: int | None) -> Workbook:
    """A workbook read from the text file at ``path``, as Workbooks.Open reads it."""
    from pyopenvba.apps.excel._io import WorkbookFileError
    from pyopenvba.apps.excel._model import Range, Workbook
    from pyopenvba.apps.excel._text_columns import typed_field

    if not path.is_file():
        raise WorkbookFileError(f"there is no file at {path}")
    data = path.read_bytes()
    text, encoding = _decoded(data)
    found = None
    if path.suffix.lower() == ".csv" and len(data) >= _SHORTEST_CSV and encoding != "utf-16":
        found = _separator(text)
    if found is not None:
        delimiter, text = found
        file_format = CSV_UTF8 if encoding == "utf-8" else CSV
    else:
        delimiter = application.text_delimiter
        file_format = {"utf-16": UNICODE_TEXT, "utf-8": UTF8_TEXT}.get(encoding, CSV if format_code == 2 else TEXT)
    rows = _rows(text, delimiter)
    if len(rows) > MAX_ROWS or any(len(fields) > MAX_COLUMNS for fields in rows):
        raise VBAUnsupportedError("opening a text file past 1,048,576 lines or 16,384 fields is not implemented: "
                                  "Excel loads part of it and warns")
    book = Workbook(application, _parenthesized(path.name), str(path.parent))
    sheet = book.add_sheet(_parenthesized(path.stem)[:_LONGEST_SHEET_NAME])
    target = Range(sheet, [Area(1, 1, 1, 1, sheet.name)])
    formulas = False
    for row, fields in enumerate(rows, start=1):
        for column, field in enumerate(fields, start=1):
            if not field:
                continue
            if field.startswith("=") and len(field) > 1:
                try:
                    target.put_as_value(row, column, field)
                except Exception as exc:
                    raise VBAUnsupportedError(f"opening a text file with the field {field!r}, which Excel cannot "
                                              "read as a formula, is not implemented") from exc
                formulas = True
            else:
                target.store_typed(row, column, typed_field(field, "General"))
    book.format_code = file_format
    # Formulas are worked out as the file opens, which leaves the workbook to save.
    book.saved = not formulas
    return book


def _parenthesized(name: str) -> str:
    """A file's name as its workbook and sheet take it, brackets made parentheses."""
    return name.replace("[", "(").replace("]", ")")


def _decoded(data: bytes) -> tuple[str, str]:
    """A file's text and how its bytes read: "utf-8" or "utf-16" after a byte order mark, "ansi" otherwise."""
    if data.startswith(_UTF8_BOM):
        text, encoding = data[len(_UTF8_BOM):].decode("utf-8", errors="replace"), "utf-8"
    elif data.startswith(_UTF16_BOM):
        text, encoding = data[len(_UTF16_BOM):].decode("utf-16-le", errors="replace"), "utf-16"
    else:
        text, encoding = data.decode("latin-1").translate(_WINDOWS_1252), "ansi"
    if "\0" in text:
        raise VBAUnsupportedError("opening a text file that holds a NUL is not implemented: Excel refuses some such "
                                  "files and drops the NULs of others")
    return text, encoding


def _separator(text: str) -> tuple[str, str] | None:
    """A CSV's delimiter and its text after a first line of sep=, if it has one: the character the line names, a
    comma where it names none; None where it names more than one, as Excel then reads the file as text."""
    end = len(text)
    for mark in ("\r", "\n"):
        found = text.find(mark)
        if found != -1:
            end = min(end, found)
    line = text[:end].lstrip(" ")
    if not line.lower().startswith("sep="):
        return ",", text
    named = line[len("sep="):]
    if len(named) > 1:
        return None
    rest = text[end:]
    rest = rest[2:] if rest.startswith("\r\n") else rest[1:]
    return named or ",", rest


def _rows(text: str, delimiter: str | None) -> list[list[str]]:
    """The fields of each line of ``text``, split at ``delimiter`` or not at all (see the module's notes)."""
    rows: list[list[str]] = []
    fields: list[str] = []
    field: list[str] = []
    quoted = started = False
    index, end = 0, len(text)
    while index < end:
        char = text[index]
        index += 1
        if quoted:
            if char == '"':
                if index < end and text[index] == '"':
                    index += 1
                else:
                    quoted = False
                    continue
        elif char == delimiter:
            fields.append("".join(field))
            field, started = [], False
            continue
        elif char in "\r\n":
            if char == "\r" and index < end and text[index] == "\n":
                index += 1
            fields.append("".join(field))
            rows.append(fields)
            fields, field, started = [], [], False
            continue
        elif char == '"' and not started:
            quoted = started = True
            continue
        if len(field) == _LONGEST_FIELD:
            # Excel drops this character and reads on as a new line.
            fields.append("".join(field))
            rows.append(fields)
            fields, field, quoted, started = [], [], False, False
            continue
        field.append(char)
        started = True
    if field or fields or started:
        fields.append("".join(field))
        rows.append(fields)
    if delimiter is None:
        rows = [[_expanded(value) for value in line] for line in rows]
    return rows


def _expanded(text: str) -> str:
    """A line with its tabs turned into spaces up to the next eighth column, as Format 5 reads it."""
    out: list[str] = []
    width = 0
    for char in text:
        if char == "\t":
            gap = _TAB_STOP - width % _TAB_STOP
            out.append(" " * gap)
            width += gap
        else:
            out.append(char)
            width += 1
    return "".join(out)
