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

Workbooks.OpenText reads a text file with the arguments it is given
(scripts/measure_open_text.py, tests/fixtures/open_text.json):

- A .csv file reads as Workbooks.Open reads one, whatever the arguments.
- Each line of any other file is a record, a line end in quotes ending
  it too, split at the delimiters in force or cut at FieldInfo's fixed
  widths, each fixed-width field without the spaces round it; StartRow's
  lines before it are left out. A field is typed as TextToColumns types
  it, its FieldInfo kind and the separators and trailing minus given
  taken as TextToColumns takes them. The workbook is saved.
- The delimiters, ConsecutiveDelimiter and Origin left out are those the
  session last used (see ImportSettings); TextQualifier is a double
  quote, StartRow 1 and DataType delimited where they are left out.
- Origin's code page reads the bytes, a UTF-8 byte order mark left out;
  UTF-16 reads after its own whatever the code page. The workbook
  answers xlCurrentPlatformText, xlUnicodeText for UTF-16, and after a
  UTF-8 byte order mark xlCSVUTF8, or 65 where OpenText is given only
  its Filename.
- Not read here: DataType left out of text holding spaces and none of
  the delimiters, and fixed widths with no FieldInfo, where Excel guesses
  the columns; a field of 32,759 characters or more, which Excel cuts to
  a length that varies with it; and the Macintosh code page past ASCII.

Workbook.SaveAs in a text format writes the active sheet alone
(scripts/measure_text_saves.py, tests/fixtures/text_saves.json):

- Its used block, from the block's first row and column, every line as
  wide as the block, empty fields included, and an empty sheet as one
  empty line; hidden rows and columns are written too, a merged cell's
  value once.
- Each cell as its own format shows it whatever its column's width, a
  number in General in at most eleven characters, a fill left out, a
  formula as its value; a conditional format's number format is not
  used, and a value its format cannot show is 255 #. A field holding a
  comma, a double quote, a tab, a CR or an LF is quoted, its quotes
  doubled, whatever the delimiter.
- xlCSV, xlCSVWindows and xlCSVMSDOS split at commas, the text formats
  at tabs; xlCSVMSDOS and xlTextMSDOS in code page 437, xlCSVUTF8 in
  UTF-8 and xlUnicodeText in UTF-16, each after its byte order mark, the
  rest in Windows-1252, a character a code page lacks written as ?. Every
  line ends with CR LF.
- The workbook takes the file's name, .csv or .txt added where it has
  none but in xlUnicodeText, and the FileFormat Excel answers for it; the
  active sheet takes the file's name, unless another sheet has it or the
  format is an MS-DOS one; and a workbook of more than one sheet is left
  unsaved. A later Save writes the same way. The Mac formats and
  xlTextPrinter are not written here.
"""

from __future__ import annotations

import codecs
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._values import ExcelError
from pyopenvba.interpreter._values import EMPTY, MISSING, to_bool, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Application, Workbook, Worksheet

CSV, CSV_UTF8, UNICODE_TEXT, UTF8_TEXT, TEXT, TEXT_MSDOS, CSV_MSDOS = 6, 62, 42, 65, -4158, 21, 24
#: The FileFormats a workbook read from text, or saved as text, answers.
TEXT_FORMATS = frozenset({CSV, CSV_UTF8, UNICODE_TEXT, UTF8_TEXT, TEXT, TEXT_MSDOS, CSV_MSDOS})
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
#: The code page each of OpenText's Origins names: xlMacintosh, xlWindows and xlMSDOS; any other is a code page.
_ORIGINS = {1: 10000, 2: 1252, 3: 437}
#: Windows-1252's characters for the bytes 0x80 to 0x9F; the five it leaves out read as the control characters.
_WINDOWS_1252 = {code: bytes([code]).decode("cp1252") for code in range(0x80, 0xA0)
                 if code not in (0x81, 0x8D, 0x8F, 0x90, 0x9D)}


@dataclass(slots=True)
class ImportSettings:
    """What Excel remembers through a session of the text a macro last split: the delimiters, whether a run of
    them is one, and OpenText's code page. Workbooks.OpenText and Range.TextToColumns use each one they are not
    given and keep each one they are; Workbooks.Open's Format leaves its delimiter alone in force, and opening a
    text file with Workbooks.Open puts the code page back to Windows-1252. TextQualifier, StartRow, FieldInfo, the
    separators and TrailingMinusNumbers are not remembered (tests/fixtures/open_text.json)."""

    tab: bool = True
    comma: bool = False
    semicolon: bool = False
    space: bool = False
    other: bool = False
    other_char: str = ""
    consecutive: bool = False
    origin: int = 1252

    def take(self, *, tab: object, comma: object, semicolon: object, space: object, other: object,
             other_char: object, consecutive: object) -> None:
        """The arguments a call gives, each kept for the calls after it."""
        for name, value in (("tab", tab), ("comma", comma), ("semicolon", semicolon), ("space", space),
                            ("other", other), ("consecutive", consecutive)):
            if value is not MISSING:
                setattr(self, name, to_bool(value))
        if other_char is not MISSING:
            self.other_char = to_text(other_char)[:1]

    def delimiters(self) -> set[str]:
        """The characters a line splits at."""
        marks = {mark for flag, mark in ((self.tab, "\t"), (self.comma, ","), (self.semicolon, ";"),
                                         (self.space, " ")) if flag}
        if self.other and self.other_char:
            marks.add(self.other_char)
        return marks

    def formatted(self, delimiter: str | None) -> None:
        """What Workbooks.Open's Format leaves in force: its delimiter alone, or none."""
        self.tab, self.comma, self.semicolon, self.space = (delimiter == mark for mark in ("\t", ",", ";", " "))
        self.other = delimiter is not None and delimiter not in ("\t", ",", ";", " ")
        if delimiter is not None and self.other:
            self.other_char = delimiter


def is_text(path: Path) -> bool:
    """Whether a file opens as text rather than as a workbook: its extension is not a workbook's."""
    return path.suffix.lower() not in WORKBOOK_EXTENSIONS


def remember_format(application: Application, format_: object, delimiter: object) -> int | None:
    """Workbooks.Open's Format, which sets what later text files split at, whatever file this open reads, and what
    OpenText and TextToColumns split at; its number, or None where it is left out."""
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
    application.text_import.formatted(application.text_delimiter)
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


def open_text_with(application: Application, path: Path, *, origin: object, start_row: object, data_type: object,
                   qualifier: object, consecutive: object, tab: object, semicolon: object, comma: object,
                   space: object, other: object, other_char: object, field_info: object, decimal: object,
                   thousands: object, trailing_minus: object, given: bool) -> Workbook:
    """A workbook read from a text file as Workbooks.OpenText reads it (see the module's notes); ``given`` is whether
    OpenText was given more than its Filename."""
    from pyopenvba.apps.excel import _text_columns as columns
    from pyopenvba.apps.excel._io import WorkbookFileError
    from pyopenvba.apps.excel._model import Range, Workbook

    settings = application.text_import
    settings.take(tab=tab, comma=comma, semicolon=semicolon, space=space, other=other, other_char=other_char,
                  consecutive=consecutive)
    if origin is not MISSING:
        code = int(to_integer(origin, "Long"))
        settings.origin = _ORIGINS.get(code, code)
    if not path.is_file():
        raise WorkbookFileError(f"there is no file at {path}")
    if path.suffix.lower() == ".csv":
        # A .csv file reads as Workbooks.Open reads one, the arguments only remembered.
        return open_text(application, path, None)
    text, file_format = _imported(path.read_bytes(), settings.origin, given)
    first = 1 if start_row is MISSING else int(to_integer(start_row, "Long"))
    if first < 1:
        raise VBAUnsupportedError("OpenText with a StartRow below 1 is not implemented")
    lines = _lines(text)[first - 1:]
    delimiters = settings.delimiters()
    if data_type is MISSING:
        if " " in text and not any(mark in text for mark in delimiters):
            raise VBAUnsupportedError("OpenText without a DataType of text holding spaces and none of its delimiters "
                                      "is not implemented: Excel may read it as fixed width, guessing the columns")
        kind = columns.DELIMITED
    else:
        kind = int(to_integer(data_type, "Long"))
        if kind not in (columns.DELIMITED, columns.FIXED_WIDTH):
            raise VBAUnsupportedError(f"OpenText with DataType {kind} is not implemented")
    if kind == columns.FIXED_WIDTH and field_info is MISSING:
        raise VBAUnsupportedError("OpenText of fixed widths with no FieldInfo is not implemented: Excel guesses the "
                                  "columns")
    info = columns.field_info_pairs(field_info, fixed=kind == columns.FIXED_WIDTH)
    quote = columns.QUALIFIERS.get(1 if qualifier is MISSING else int(to_integer(qualifier, "Long")), '"')
    separators = ("." if decimal is MISSING else to_text(decimal)[:1] or ".",
                  "," if thousands is MISSING else to_text(thousands)[:1] or ",")
    minus = trailing_minus is not MISSING and to_bool(trailing_minus)
    rows: list[list[str]] = []
    for line in lines:
        fields = columns.fixed_fields(line, info) if kind == columns.FIXED_WIDTH else \
            columns.split_fields(line, delimiters, quote, settings.consecutive)
        if any(len(field) >= _LONGEST_FIELD for field in fields):
            raise VBAUnsupportedError("OpenText of a field of 32,759 characters or more is not implemented: Excel "
                                      "cuts it to a length that varies with it and drops the rest of its line")
        rows.append(fields)
    if len(rows) > MAX_ROWS or any(len(fields) > MAX_COLUMNS for fields in rows):
        raise VBAUnsupportedError("opening a text file past 1,048,576 lines or 16,384 fields is not implemented: "
                                  "Excel loads part of it and warns")
    book = Workbook(application, _parenthesized(path.name), str(path.parent))
    sheet = book.add_sheet(_parenthesized(path.stem)[:_LONGEST_SHEET_NAME])
    target = Range(sheet, [Area(1, 1, 1, 1, sheet.name)])
    kinds = [one for _, one in info]
    for row, fields in enumerate(rows, start=1):
        kept = [(field, kinds[index] if index < len(kinds) else columns.GENERAL) for index, field in enumerate(fields)]
        for column, (field, field_kind) in enumerate([one for one in kept if one[1] != columns.SKIP], start=1):
            if field:
                columns.write_field(target, row, column, field, field_kind, separators, minus)
    book.format_code = file_format
    # Unlike Workbooks.Open's, a workbook OpenText reads is saved, formulas and all.
    book.saved = True
    return book


def _lines(text: str) -> list[str]:
    """A text file's lines, which OpenText reads one record each, a line end in quotes ending its line too."""
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    if lines[-1] == "":
        lines.pop()
    return lines


def _imported(data: bytes, code_page: int, given: bool) -> tuple[str, int]:
    """A text file's text as OpenText reads it, and the FileFormat its workbook answers: UTF-16 after its byte order
    mark, whatever the code page; anything else in the code page, a UTF-8 byte order mark left out, and the
    workbook xlCSVUTF8 after one where OpenText is given more than its Filename, 65 where it is not."""
    if data.startswith(_UTF16_BOM):
        text, file_format = data[len(_UTF16_BOM):].decode("utf-16-le", errors="replace"), UNICODE_TEXT
    else:
        file_format = TEXT
        if data.startswith(_UTF8_BOM):
            data, file_format = data[len(_UTF8_BOM):], CSV_UTF8 if given else UTF8_TEXT
        text = _in_code_page(data, code_page)
    if "\0" in text:
        raise VBAUnsupportedError("opening a text file that holds a NUL is not implemented: Excel refuses some such "
                                  "files and drops the NULs of others")
    return text, file_format


def _in_code_page(data: bytes, code_page: int) -> str:
    """Bytes read in a code page OpenText's Origin names."""
    if code_page == 1252:
        return data.decode("latin-1").translate(_WINDOWS_1252)
    if code_page == 437:
        return data.decode("cp437")
    if code_page == 65001:
        return data.decode("utf-8", errors="replace")
    if code_page == 10000 and data.isascii():
        return data.decode("ascii")
    if code_page == 10000:
        raise VBAUnsupportedError("OpenText of text past ASCII in the Macintosh code page is not implemented: Windows "
                                  "reads some of its characters otherwise than the model can")
    raise VBAUnsupportedError(f"OpenText of text in code page {code_page} is not implemented")


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


# --- writing ------------------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class TextWriter:
    """How a text FileFormat writes: what splits its fields, its code page and byte order mark, the FileFormat the
    workbook answers after it, whether the active sheet takes the file's name, and the extension a name without
    one takes."""

    delimiter: str
    encoding: str
    bom: bytes
    answers: int
    renames: bool
    extension: str


_WRITERS = {
    CSV: TextWriter(",", "cp1252", b"", CSV, True, ".csv"),
    23: TextWriter(",", "cp1252", b"", CSV, True, ".csv"),  # xlCSVWindows
    CSV_MSDOS: TextWriter(",", "cp437", b"", CSV_MSDOS, False, ".csv"),
    CSV_UTF8: TextWriter(",", "utf-8", _UTF8_BOM, CSV_UTF8, True, ".csv"),
    # xlUnicodeText adds no extension to a name without one.
    UNICODE_TEXT: TextWriter("\t", "utf-16-le", _UTF16_BOM, UNICODE_TEXT, True, ""),
    TEXT: TextWriter("\t", "cp1252", b"", TEXT, True, ".txt"),
    20: TextWriter("\t", "cp1252", b"", TEXT, True, ".txt"),  # xlTextWindows
    TEXT_MSDOS: TextWriter("\t", "cp437", b"", TEXT_MSDOS, False, ".txt"),
}
#: Text formats Excel writes that the model does not: the Mac ones, whose line ends and characters were measured
#: once each, the printer's fixed columns, and the UTF-8 text a file of that kind opens as.
_UNWRITTEN = {19: "xlTextMac", 22: "xlCSVMac", 36: "xlTextPrinter", UTF8_TEXT: "UTF-8 text"}
#: What makes a field quoted, whatever the delimiter.
_QUOTED = frozenset(',"\t\r\n')
#: How many # a text file holds for a value its format cannot show.
_UNSHOWN = 255
#: The bytes Windows-1252 leaves out, which Windows writes for the control characters of the same number.
_CONTROL_BYTES = frozenset((0x81, 0x8D, 0x8F, 0x90, 0x9D))


def _question_marks(problem: UnicodeError) -> tuple[bytes, int]:
    """A character a code page lacks written as ?, as Excel writes it, but Windows-1252's five control characters
    as their bytes."""
    assert isinstance(problem, UnicodeEncodeError)
    piece = problem.object[problem.start:problem.end]
    kept = problem.encoding == "cp1252"
    return bytes(ord(char) if kept and ord(char) in _CONTROL_BYTES else ord("?") for char in piece), problem.end


codecs.register_error("pyopenvba.question_marks", _question_marks)


def text_writer(file_format: int) -> TextWriter | None:
    """How SaveAs writes ``file_format``, or None for a workbook's format; the text formats the model does not write
    report themselves unsupported."""
    if file_format in _UNWRITTEN:
        raise VBAUnsupportedError(f"saving a workbook as {_UNWRITTEN[file_format]} (FileFormat {file_format}) is not "
                                  "implemented")
    return _WRITERS.get(file_format)


def save_text(book: Workbook, target: Path, file_format: int) -> None:
    """The workbook's active sheet written to ``target`` as SaveAs writes ``file_format`` (see the module's notes)."""
    writer = text_writer(file_format)
    sheet = book.active_sheet
    if writer is None or sheet is None:
        raise VBAUnsupportedError(f"saving FileFormat {file_format} as text is not implemented")
    bounds = sheet.used_bounds()
    lines: list[str] = []
    if bounds is None:
        lines.append("")
    else:
        top, left, bottom, right = bounds
        for row in range(top, bottom + 1):
            lines.append(writer.delimiter.join(_quoted(_shown(sheet, row, column))
                                               for column in range(left, right + 1)))
    text = "".join(f"{line}\r\n" for line in lines)
    errors = "surrogatepass" if writer.encoding == "utf-16-le" else "pyopenvba.question_marks"
    target.write_bytes(writer.bom + text.encode(writer.encoding, errors=errors))


def saved_as_text(book: Workbook, file_format: int) -> None:
    """What a SaveAs in a text format leaves of the workbook: the FileFormat Excel answers, and the active sheet
    named after the file (tests/fixtures/text_saves.json)."""
    from pyopenvba.apps.excel._protection import APPLICATION_DEFINED, refuse_structure
    from pyopenvba.apps.excel._worksheet_rename import rename

    from pyopenvba.apps.excel import _names

    writer = _WRITERS[file_format]
    book.format_code = writer.answers
    sheet = book.active_sheet
    if writer.renames and sheet is not None:
        wanted = _parenthesized(Path(book.name).stem)[:_LONGEST_SHEET_NAME]
        # A name another sheet has already leaves the sheet its own.
        taken = any(other is not sheet and _names.same_sheet_name(other.name, wanted) for other in book.sheets_)
        if wanted != sheet.name and not taken:
            refuse_structure(book, APPLICATION_DEFINED)
            rename(sheet, wanted)


def _shown(sheet: Worksheet, row: int, column: int) -> str:
    """A cell's field: what its own format shows whatever its column's width, its fill left out and a conditional
    format's number format not used; a value the format cannot show, a negative date, as 255 #
    (see formula._display.shown)."""
    from pyopenvba.formula._display import UndisplayableError, shown

    cell = sheet.cells_.get((row, column))
    if cell is None:
        return ""
    value = sheet.book.calculator.value_of(sheet.name, row, column) if cell.formula else cell.value
    if value is EMPTY:
        return ""
    if isinstance(value, ExcelError):
        return value.name
    try:
        return shown(value, sheet.style_at(row, column).number_format, epoch_1904=sheet.book.epoch_1904)[0]
    except UndisplayableError:
        return "#" * _UNSHOWN


def _quoted(text: str) -> str:
    """A field as a text file holds it: in quotes, its own doubled, where it holds a delimiter or a line end."""
    if any(char in _QUOTED for char in text):
        return '"' + text.replace('"', '""') + '"'
    return text
