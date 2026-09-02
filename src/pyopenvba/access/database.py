"""``AccessDatabase``: a Jet 4 / ACE database opened as tables of rows.

This is the engine's facade.  It reads the catalog (``MSysObjects``) to
find tables by name, parses their definitions, and walks their owned
pages to yield rows as plain Python values.
"""

from __future__ import annotations

from collections.abc import Iterator
from dataclasses import dataclass
from pathlib import Path

from pyopenvba.access_read import AccessError
from pyopenvba.access._lval import read_long_value
from pyopenvba.access._pages import (
    PAGE_DATA,
    DatabaseHeader,
    PageStore,
    page_owner,
    read_usage_map_ref,
    row_bytes,
    row_pointer,
    row_slots,
    ROW_DELETED,
    ROW_OVERFLOW,
)
from pyopenvba.access._rows import (
    LongValueRef,
    RawRow,
    decode_scalar,
    decode_text,
    split_row,
)
from pyopenvba.access._tdef import (
    TYPE_MEMO,
    ColumnDef,
    TableDefinition,
    parse_table_definition,
)

MSYS_OBJECTS_PAGE = 2

# MSysObjects.Type values.
OBJECT_TABLE = 1
OBJECT_QUERY = 5
OBJECT_LINKED_TABLE = 6
OBJECT_RELATIONSHIP = 8
OBJECT_FORM = -32768
OBJECT_REPORT = -32764
OBJECT_MACRO = -32766
OBJECT_MODULE = -32761

# MSysObjects.Flags bits Access sets on its own objects.
FLAG_SYSTEM = 0x80000000
FLAG_HIDDEN = 0x00000008


@dataclass(frozen=True)
class CatalogEntry:
    """One MSysObjects row."""

    id: int
    parent_id: int
    name: str
    type: int
    flags: int
    owner: bytes | None
    date_create: object
    date_update: object
    page: int
    row: int

    @property
    def is_system(self) -> bool:
        return bool(self.flags & FLAG_SYSTEM) or self.name.startswith("MSys")

    @property
    def is_table(self) -> bool:
        return self.type in (OBJECT_TABLE, OBJECT_LINKED_TABLE)


class Table:
    """A table: its definition plus the rows on its owned pages."""

    def __init__(self, database: AccessDatabase, definition: TableDefinition, name: str) -> None:
        self._db = database
        self.definition = definition
        self.name = name

    @property
    def columns(self) -> list[ColumnDef]:
        return self.definition.columns_by_number()

    @property
    def column_names(self) -> list[str]:
        return [c.name for c in self.columns]

    @property
    def row_count(self) -> int:
        return self.definition.row_count

    def data_pages(self) -> list[int]:
        """Data pages of this table, in owned-page order.  The owned map
        also lists LVAL pages, which carry a different owner tag."""
        store = self._db.store
        owned = read_usage_map_ref(store, self.definition.owned_pages_ref)
        out: list[int] = []
        for page in owned.pages():
            if page >= store.page_count:
                continue
            raw = store.read(page)
            if raw[0] == PAGE_DATA and page_owner(raw) == self.definition.page:
                out.append(page)
        return out

    def raw_rows(self) -> Iterator[tuple[int, int, bytes]]:
        """``(page, slot, row_bytes)`` for every live row, overflow rows
        followed to where they live."""
        store = self._db.store
        for page in self.data_pages():
            raw_page = store.read(page)
            slots = row_slots(raw_page)
            for slot, entry in enumerate(slots):
                if entry & ROW_DELETED:
                    continue
                data = row_bytes(raw_page, slot)
                if data is None:
                    continue
                if entry & ROW_OVERFLOW:
                    target_row, target_page = row_pointer(data)
                    target = row_bytes(
                        store.read(target_page), target_row, overflow_target=True
                    )
                    if target is None:
                        raise AccessError(
                            f"overflow row ({page}, {slot}) points at nothing"
                        )
                    yield target_page, target_row, target
                    continue
                yield page, slot, data

    def rows(self) -> Iterator[dict[str, object]]:
        for _page, _slot, data in self.raw_rows():
            yield self.decode(split_row(self.definition, data))

    def decode(self, raw: RawRow) -> dict[str, object]:
        out: dict[str, object] = {}
        for column in self.columns:
            value = raw.values.get(column.number)
            if value is None:
                if column.type_code == 0x01 and raw.present.get(column.number) is False:
                    out[column.name] = False
                else:
                    out[column.name] = None
                continue
            decoded = decode_scalar(column, value)
            if isinstance(decoded, LongValueRef):
                data = read_long_value(self._db.store, decoded)
                out[column.name] = decode_text(data) if column.type_code == TYPE_MEMO else data
            else:
                out[column.name] = decoded
        return out


class AccessDatabase:
    """Open a ``.accdb`` / Jet 4 ``.mdb`` and read its tables."""

    def __init__(self, source: str | Path | bytes) -> None:
        if isinstance(source, (str, Path)):
            self.path: Path | None = Path(source)
            data = self.path.read_bytes()
        else:
            self.path = None
            data = source
        self.store = PageStore(data)
        self.header = DatabaseHeader.from_page(self.store.read(0))
        self._catalog: list[CatalogEntry] | None = None
        self._definitions: dict[int, TableDefinition] = {}

    def __enter__(self) -> AccessDatabase:
        return self

    def __exit__(self, *_exc: object) -> None:
        return None

    # -- catalog -----------------------------------------------------------

    def definition(self, page: int) -> TableDefinition:
        if page not in self._definitions:
            self._definitions[page] = parse_table_definition(self.store, page)
        return self._definitions[page]

    def catalog(self) -> list[CatalogEntry]:
        if self._catalog is None:
            table = Table(self, self.definition(MSYS_OBJECTS_PAGE), "MSysObjects")
            entries: list[CatalogEntry] = []
            for page, slot, data in table.raw_rows():
                values = table.decode(split_row(table.definition, data))
                owner = values.get("Owner")
                entries.append(
                    CatalogEntry(
                        id=_as_int(values["Id"]),
                        parent_id=_as_int(values["ParentId"]),
                        name=str(values["Name"] or ""),
                        type=_as_int(values["Type"]),
                        flags=_as_int(values["Flags"]) & 0xFFFFFFFF,
                        owner=owner if isinstance(owner, bytes) else None,
                        date_create=values.get("DateCreate"),
                        date_update=values.get("DateUpdate"),
                        page=page,
                        row=slot,
                    )
                )
            self._catalog = entries
        return list(self._catalog)

    def table_entries(self, include_system: bool = False) -> list[CatalogEntry]:
        return [
            e
            for e in self.catalog()
            if e.type == OBJECT_TABLE and (include_system or not e.is_system)
        ]

    def table_names(self, include_system: bool = False) -> list[str]:
        return sorted(e.name for e in self.table_entries(include_system))

    def table(self, name: str) -> Table:
        for entry in self.catalog():
            if entry.type == OBJECT_TABLE and entry.name.lower() == name.lower():
                return Table(self, self.definition(entry.id), entry.name)
        raise AccessError(f"no table named {name!r}")

    def tables(self, include_system: bool = False) -> list[Table]:
        return [self.table(e.name) for e in self.table_entries(include_system)]


def _as_int(value: object) -> int:
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, int):
        return value
    raise AccessError(f"expected an integer catalog field, got {type(value).__name__}")
