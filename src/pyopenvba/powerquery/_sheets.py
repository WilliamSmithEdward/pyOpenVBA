"""Loading a query's result into a worksheet table.

Where a query loads is not the package's decision.  The metadata says
what the Queries pane shows -- ``FillEnabled``, ``FillObjectType``,
``FillTarget`` -- but Excel puts the data on a sheet only when the
workbook also carries the objects that do it: a connection through the
mashup provider, a query table, a table, and the sheet's reference to
that table.  Setting the metadata alone changes nothing at all
(measured: Excel opened such a workbook, refreshed, and created no
table).

So this module writes the objects too.  The column names have to be
given, because knowing them means evaluating the query and only the
mashup engine can do that; Excel reconciles them with the real result on
the first refresh, which is also when it fills the rows in.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from pyopenvba.exceptions import PowerQueryError
from pyopenvba.powerquery._opc import OpcFile

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_CT_CONNECTIONS = "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"
_CT_TABLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"
_CT_QUERY_TABLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"
_CONTENT_TYPES = "[Content_Types].xml"
_WORKBOOK = "xl/workbook.xml"
_WORKBOOK_RELS = "xl/_rels/workbook.xml.rels"
_CONNECTIONS = "xl/connections.xml"
#: The style Excel gives a table it makes for a query.
_TABLE_STYLE = "TableStyleMedium7"
_CELL = re.compile(r"^([A-Za-z]+)([0-9]+)$")
#: A table name may not look like a cell reference, in either notation.
#: Excel does not repair such a workbook, it refuses to open it, so the
#: name is checked before anything is written (measured: a table called
#: ``V1`` gave "Excel cannot open the file").
_LOOKS_LIKE_A_CELL = re.compile(r"^(?:[A-Za-z]{1,3}[0-9]{1,7}|[Rr][0-9]+[Cc][0-9]+|[Cc]|[Rr])$")


def column_letter(index: int) -> str:
    """``1`` is ``A``, ``27`` is ``AA``."""
    if index < 1:
        raise PowerQueryError("a column number starts at one")
    out = ""
    while index:
        index, rest = divmod(index - 1, 26)
        out = chr(ord("A") + rest) + out
    return out


def column_number(letters: str) -> int:
    out = 0
    for char in letters.upper():
        out = out * 26 + (ord(char) - ord("A") + 1)
    return out


@dataclass(frozen=True)
class CellRef:
    """A cell, as a column number and a row number."""

    column: int
    row: int

    @classmethod
    def parse(cls, text: str) -> CellRef:
        match = _CELL.match(text.replace("$", ""))
        if match is None:
            raise PowerQueryError(f"{text!r} is not a cell reference like 'A1'")
        return cls(column_number(match.group(1)), int(match.group(2)))

    def __str__(self) -> str:
        return f"{column_letter(self.column)}{self.row}"


def _escape(value: str) -> str:
    return (
        value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
    )


def _next_relationship(rels: str) -> str:
    used = {int(number) for number in re.findall(r'Id="rId(\d+)"', rels)}
    index = 1
    while index in used:
        index += 1
    return f"rId{index}"


def _relationships(package: OpcFile, part: str) -> str:
    if package.has(part):
        return package.read(part).decode("utf-8")
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        "</Relationships>"
    )


def add_relationship(package: OpcFile, part: str, kind: str, target: str) -> str:
    raw = _relationships(package, part)
    identifier = _next_relationship(raw)
    package.write(
        part,
        raw.replace(
            "</Relationships>",
            f'<Relationship Id="{identifier}" Type="{kind}" Target="{target}"/></Relationships>',
        ).encode("utf-8"),
    )
    return identifier


def drop_relationship(package: OpcFile, part: str, target: str) -> None:
    if not package.has(part):
        return
    raw = package.read(part).decode("utf-8")
    package.write(part, re.sub(rf'<Relationship[^>]*Target="{re.escape(target)}"[^>]*/>', "", raw).encode("utf-8"))


def add_content_type(package: OpcFile, part: str, content_type: str) -> None:
    raw = package.read(_CONTENT_TYPES).decode("utf-8")
    override = f'<Override PartName="/{part}" ContentType="{content_type}"/>'
    if override not in raw:
        package.write(_CONTENT_TYPES, raw.replace("</Types>", override + "</Types>").encode("utf-8"))


def drop_content_type(package: OpcFile, part: str) -> None:
    raw = package.read(_CONTENT_TYPES).decode("utf-8")
    package.write(
        _CONTENT_TYPES,
        re.sub(rf'<Override PartName="/{re.escape(part)}"[^>]*/>', "", raw).encode("utf-8"),
    )


def sheet_part(package: OpcFile, sheet: str | int) -> tuple[str, str]:
    """The worksheet part for a sheet name or one-based position, and the
    name of the sheet."""
    workbook = package.read(_WORKBOOK).decode("utf-8")
    rels = package.read(_WORKBOOK_RELS).decode("utf-8")
    targets = dict(re.findall(r'<Relationship Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    sheets = re.findall(r'<sheet\b[^>]*name="([^"]*)"[^>]*r:id="([^"]+)"', workbook)
    if not sheets:
        raise PowerQueryError("this workbook lists no worksheets")
    if isinstance(sheet, int):
        if not 1 <= sheet <= len(sheets):
            raise PowerQueryError(f"this workbook has {len(sheets)} sheets, so sheet {sheet} is not one of them")
        name, identifier = sheets[sheet - 1]
    else:
        found = [pair for pair in sheets if pair[0] == sheet]
        if not found:
            known = ", ".join(pair[0] for pair in sheets)
            raise PowerQueryError(f"this workbook has no sheet named {sheet!r}; it has: {known}")
        name, identifier = found[0]
    target = targets.get(identifier)
    if target is None:
        raise PowerQueryError(f"the sheet {name!r} has no part behind it")
    return "xl/" + target.lstrip("/").removeprefix("xl/"), name


def occupied(package: OpcFile, part: str) -> set[tuple[int, int]]:
    """Which cells of a worksheet already hold something."""
    raw = package.read(part).decode("utf-8")
    cells: set[tuple[int, int]] = set()
    for reference in re.findall(r'<c\b[^>]*\br="([A-Z]+\d+)"', raw):
        ref = CellRef.parse(reference)
        cells.add((ref.column, ref.row))
    return cells


def _numbered(package: OpcFile, pattern: str) -> int:
    used = [
        int(match.group(1))
        for name in package.names()
        if (match := re.match(pattern, name))
    ]
    return max(used, default=0) + 1


def load_to_sheet(
    package: OpcFile,
    query: str,
    columns: list[str],
    *,
    sheet: str | int = 1,
    cell: str = "A1",
    table_name: str | None = None,
) -> str:
    """Write the objects that put `query`'s result on a worksheet.

    Returns the name of the table.  Excel fills the rows and settles the
    column names on its first refresh; what is written here is the header
    row and the objects that point at the query.
    """
    if not columns:
        raise PowerQueryError("a table needs at least one column")
    if len(set(columns)) != len(columns):
        raise PowerQueryError("two columns of one table cannot share a name")
    name = table_name or query
    if _LOOKS_LIKE_A_CELL.match(name):
        raise PowerQueryError(
            f"a table cannot be called {name!r}: Excel reads that as a cell reference and "
            "refuses to open the workbook; pass table_name= with something else"
        )
    if " " in name:
        raise PowerQueryError(
            f"a table cannot be called {name!r}: Excel allows no spaces in a table name; "
            "pass table_name= with something else"
        )
    part, sheet_name = sheet_part(package, sheet)
    start = CellRef.parse(cell)
    end = CellRef(start.column + len(columns) - 1, start.row + 1)
    taken = occupied(package, part)
    clash = sorted(
        (column, row)
        for column in range(start.column, end.column + 1)
        for row in range(start.row, end.row + 1)
        if (column, row) in taken
    )
    if clash:
        where = CellRef(*clash[0])
        raise PowerQueryError(f"{sheet_name}!{where} already holds something; give the table an empty range")

    number = _numbered(package, r"xl/tables/table(\d+)\.xml$")
    connection_id = _add_connection(package, query)
    table_part = f"xl/tables/table{number}.xml"
    query_part = f"xl/queryTables/queryTable{number}.xml"
    reference = f"{start}:{end}"

    fields = "".join(
        f'<queryTableField id="{index + 1}" name="{_escape(column)}" tableColumnId="{index + 1}"/>'
        for index, column in enumerate(columns)
    )
    package.write(query_part, (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f'<queryTable xmlns="{NS_MAIN}" name="ExternalData_{number}" connectionId="{connection_id}"'
        ' autoFormatId="16" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="1"'
        ' applyPatternFormats="1" applyAlignmentFormats="0" applyWidthHeightFormats="0">'
        f'<queryTableRefresh nextId="{len(columns) + 1}">'
        f'<queryTableFields count="{len(columns)}">{fields}</queryTableFields>'
        "</queryTableRefresh></queryTable>"
    ).encode("utf-8"))
    add_content_type(package, query_part, _CT_QUERY_TABLE)

    table_columns = "".join(
        f'<tableColumn id="{index + 1}" uniqueName="{index + 1}" name="{_escape(column)}"'
        f' queryTableFieldId="{index + 1}"/>'
        for index, column in enumerate(columns)
    )
    package.write(table_part, (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f'<table xmlns="{NS_MAIN}" id="{number}" name="{_escape(name)}" displayName="{_escape(name)}"'
        f' ref="{reference}" tableType="queryTable" totalsRowShown="0">'
        f'<autoFilter ref="{reference}"/>'
        f'<tableColumns count="{len(columns)}">{table_columns}</tableColumns>'
        f'<tableStyleInfo name="{_TABLE_STYLE}" showFirstColumn="0" showLastColumn="0"'
        ' showRowStripes="1" showColumnStripes="0"/></table>'
    ).encode("utf-8"))
    add_content_type(package, table_part, _CT_TABLE)
    add_relationship(
        package, f"xl/tables/_rels/table{number}.xml.rels", f"{_REL}/queryTable",
        f"../queryTables/queryTable{number}.xml",
    )

    rels_part = part.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels"
    identifier = add_relationship(package, rels_part, f"{_REL}/table", f"../tables/table{number}.xml")
    _add_table_to_sheet(package, part, identifier, start, columns, reference)
    _add_defined_name(package, sheet_name, number, reference)
    return name


def _add_connection(package: OpcFile, query: str) -> int:
    if package.has(_CONNECTIONS):
        raw = package.read(_CONNECTIONS).decode("utf-8")
        identifiers = [int(number) for number in re.findall(r'<connection\b[^>]*\bid="(\d+)"', raw)]
        identifier = max(identifiers, default=0) + 1
    else:
        raw = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            f'<connections xmlns="{NS_MAIN}"></connections>'
        )
        identifier = 1
        add_content_type(package, _CONNECTIONS, _CT_CONNECTIONS)
        add_relationship(package, _WORKBOOK_RELS, f"{_REL}/connections", "connections.xml")
    connection = (
        f'<connection id="{identifier}" keepAlive="1" name="Query - {_escape(query)}" type="5"'
        ' refreshedVersion="8" background="1" saveData="1">'
        '<dbPr connection="Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;'
        f'Location={_escape(query)};Extended Properties=&quot;&quot;"'
        f' command="SELECT * FROM [{_escape(query)}]"/></connection>'
    )
    package.write(_CONNECTIONS, raw.replace("</connections>", connection + "</connections>").encode("utf-8"))
    return identifier


def _add_table_to_sheet(
    package: OpcFile, part: str, identifier: str, start: CellRef, columns: list[str], reference: str
) -> None:
    raw = package.read(part).decode("utf-8")
    header = "".join(
        f'<c r="{CellRef(start.column + index, start.row)}" t="inlineStr"><is><t>{_escape(column)}</t></is></c>'
        for index, column in enumerate(columns)
    )
    raw = _write_header(raw, header, start, len(columns))
    if "<tableParts" in raw:
        raw = re.sub(
            r'<tableParts count="(\d+)">',
            lambda match: f'<tableParts count="{int(match.group(1)) + 1}">',
            raw,
        )
        raw = raw.replace("</tableParts>", f'<tablePart r:id="{identifier}"/></tableParts>')
    else:
        raw = raw.replace(
            "</worksheet>", f'<tableParts count="1"><tablePart r:id="{identifier}"/></tableParts></worksheet>'
        )
    raw = _widen_dimension(raw, reference)
    package.write(part, raw.encode("utf-8"))


def _write_header(sheet: str, cells: str, start: CellRef, width: int) -> str:
    """Put the header cells on the sheet.

    A row of that number may already be there -- a second table can share
    a header row with the first -- in which case the cells join it, in
    column order.  Otherwise the row goes in among the others, in order.
    """
    if "<sheetData/>" in sheet:
        sheet = sheet.replace("<sheetData/>", "<sheetData></sheetData>")
    if "<sheetData>" not in sheet:
        raise PowerQueryError("this worksheet has no sheetData to write into")
    body_at = sheet.index("<sheetData>") + len("<sheetData>")
    end_at = sheet.index("</sheetData>")
    spans = f'{start.column}:{start.column + width - 1}'
    row = f'<row r="{start.row}" spans="{spans}">{cells}</row>'
    for match in re.finditer(r'<row\b[^>]*\br="(\d+)"[^>]*>', sheet[body_at:end_at]):
        number = int(match.group(1))
        if number == start.row:
            open_at = body_at + match.end()
            close_at = sheet.index("</row>", open_at)
            inside = sheet[open_at:close_at]
            at = close_at
            for cell in re.finditer(r'<c\b[^>]*\br="([A-Z]+)(\d+)"', inside):
                if column_number(cell.group(1)) > start.column:
                    at = open_at + cell.start()
                    break
            return sheet[:at] + cells + sheet[at:]
        if number > start.row:
            at = body_at + match.start()
            return sheet[:at] + row + sheet[at:]
    return sheet[:end_at] + row + sheet[end_at:]


def _widen_dimension(sheet: str, reference: str) -> str:
    match = re.search(r'<dimension ref="([^"]*)"/>', sheet)
    if match is None:
        return sheet
    wanted = [CellRef.parse(part) for part in reference.split(":")]
    have = [CellRef.parse(part) for part in match.group(1).split(":")]
    if len(have) == 1:
        have = [have[0], have[0]]
    first = CellRef(min(have[0].column, wanted[0].column), min(have[0].row, wanted[0].row))
    last = CellRef(max(have[1].column, wanted[1].column), max(have[1].row, wanted[1].row))
    return sheet.replace(match.group(0), f'<dimension ref="{first}:{last}"/>')


def _add_defined_name(package: OpcFile, sheet_name: str, number: int, reference: str) -> None:
    raw = package.read(_WORKBOOK).decode("utf-8")
    first, last = reference.split(":")
    absolute = f"${column_letter(CellRef.parse(first).column)}${CellRef.parse(first).row}"
    absolute += f":${column_letter(CellRef.parse(last).column)}${CellRef.parse(last).row}"
    quoted = f"'{sheet_name}'" if re.search(r"[^A-Za-z0-9_]", sheet_name) else sheet_name
    defined = (
        f'<definedName name="ExternalData_{number}" localSheetId="0" hidden="1">'
        f"{quoted}!{absolute}</definedName>"
    )
    if "<definedNames>" in raw:
        raw = raw.replace("<definedNames>", "<definedNames>" + defined)
    elif "<calcPr" in raw:
        raw = raw.replace("<calcPr", f"<definedNames>{defined}</definedNames><calcPr", 1)
    else:
        raw = raw.replace("</workbook>", f"<definedNames>{defined}</definedNames></workbook>")
    package.write(_WORKBOOK, raw.encode("utf-8"))


def unload_from_sheet(package: OpcFile, query: str) -> bool:
    """Take away the objects that put `query` on a sheet.

    Gives back whether anything was there to take away.  The table, its
    query table, the connection, the sheet's reference to the table, the
    header cells and the hidden name all go, which is what Excel leaves
    behind when a query goes back to being connection-only.
    """
    table_part = _table_of(package, query)
    if table_part is None:
        return False
    number = re.findall(r"(\d+)", table_part)[-1]
    raw = package.read(table_part).decode("utf-8")
    reference = re.search(r'\bref="([^"]+)"', raw)
    for part in package.names():
        if not part.startswith("xl/worksheets/") or not part.endswith(".xml"):
            continue
        rels_part = part.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels"
        if not package.has(rels_part):
            continue
        rels = package.read(rels_part).decode("utf-8")
        found = re.search(
            rf'<Relationship Id="([^"]+)"[^>]*Target="\.\./tables/table{number}\.xml"[^>]*/>', rels
        )
        if found is None:
            continue
        drop_relationship(package, rels_part, f"../tables/table{number}.xml")
        if "<Relationship " not in package.read(rels_part).decode("utf-8"):
            package.remove(rels_part)
        sheet = package.read(part).decode("utf-8")
        sheet = re.sub(rf'<tablePart r:id="{found.group(1)}"\s*/>', "", sheet)
        sheet = re.sub(
            r'<tableParts count="(\d+)">',
            lambda match: f'<tableParts count="{max(int(match.group(1)) - 1, 0)}">',
            sheet,
        )
        sheet = sheet.replace('<tableParts count="0"></tableParts>', "")
        if reference is not None:
            sheet = _clear_range(sheet, reference.group(1))
        package.write(part, sheet.encode("utf-8"))
        break
    query_part = f"xl/queryTables/queryTable{number}.xml"
    for part in (table_part, query_part):
        if package.has(part):
            package.remove(part)
            drop_content_type(package, part)
    rels_part = f"xl/tables/_rels/table{number}.xml.rels"
    if package.has(rels_part):
        package.remove(rels_part)
    _drop_connection(package, query)
    workbook = package.read(_WORKBOOK).decode("utf-8")
    workbook = re.sub(rf'<definedName name="ExternalData_{number}"[^>]*>[^<]*</definedName>', "", workbook)
    workbook = workbook.replace("<definedNames></definedNames>", "")
    package.write(_WORKBOOK, workbook.encode("utf-8"))
    return True


def _table_of(package: OpcFile, query: str) -> str | None:
    """The table part a query's query table feeds, if there is one."""
    if not package.has(_CONNECTIONS):
        return None
    raw = package.read(_CONNECTIONS).decode("utf-8")
    for block in re.findall(r"<connection\b.*?</connection>", raw, re.S):
        if f"Location={_escape(query)};" not in block and f"[{_escape(query)}]" not in block:
            continue
        identifier = re.search(r'\bid="(\d+)"', block)
        if identifier is None:
            continue
        for part in package.names():
            if not re.match(r"xl/queryTables/queryTable\d+\.xml$", part):
                continue
            if f'connectionId="{identifier.group(1)}"' in package.read(part).decode("utf-8"):
                number = re.findall(r"(\d+)", part)[-1]
                table = f"xl/tables/table{number}.xml"
                return table if package.has(table) else None
    return None


def _drop_connection(package: OpcFile, query: str) -> None:
    """Take a query's connection out, and the whole part with it when that
    was the last one.

    A connections part holding no connections is not one Excel accepts:
    it opens the workbook with an error rather than with the queries
    (measured).
    """
    if not package.has(_CONNECTIONS):
        return
    raw = package.read(_CONNECTIONS).decode("utf-8")
    kept = [
        block
        for block in re.findall(r"<connection\b.*?</connection>", raw, re.S)
        if f"Location={_escape(query)};" not in block
    ]
    if not kept:
        package.remove(_CONNECTIONS)
        drop_content_type(package, _CONNECTIONS)
        drop_relationship(package, _WORKBOOK_RELS, "connections.xml")
        return
    head = raw[: raw.index("<connections")]
    opening = re.search(r"<connections\b[^>]*>", raw)
    if opening is None:  # pragma: no cover - written by us or by Excel
        return
    package.write(
        _CONNECTIONS,
        (head + opening.group(0) + "".join(kept) + "</connections>").encode("utf-8"),
    )


def _clear_range(sheet: str, reference: str) -> str:
    """Take the cells of a range out of a worksheet, and any row left
    empty with them."""
    parts = [CellRef.parse(part) for part in reference.split(":")]
    first, last = parts[0], parts[-1]

    def keep_cell(match: re.Match[str]) -> str:
        ref = CellRef.parse(match.group(1))
        inside = first.column <= ref.column <= last.column and first.row <= ref.row <= last.row
        return "" if inside else match.group(0)

    sheet = re.sub(r'<c\b[^>]*\br="([A-Z]+\d+)"(?:[^>]*/>|[^>]*>.*?</c>)', keep_cell, sheet, flags=re.S)
    return re.sub(r"<row\b[^>]*>\s*</row>", "", sheet)
