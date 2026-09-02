"""``AccessDatabase``: a Jet 4 / ACE database opened as tables of rows.

This is the engine's facade.  It reads the catalog (``MSysObjects``) to
find tables by name, parses their definitions, walks their owned pages
to yield rows as plain Python values, and -- for tables without long
values -- inserts, updates and deletes rows the way the engine does,
indexes and counters included.  ``save()`` writes the pages back.
"""

from __future__ import annotations

import struct
from collections.abc import Iterator
from dataclasses import dataclass
from pathlib import Path

from pyopenvba.access_read import AccessError
from pyopenvba.access._alloc import add_to_map, allocate_page, remove_from_map
from pyopenvba.access._btree import BTree
from pyopenvba.access._datapage import DataPage
from pyopenvba.access._index import decode_key, encode_key, leaf_entries
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
    encode_row,
    encode_scalar,
    split_row,
)
from pyopenvba.access._tdef import (
    INDEX_IGNORE_NULLS,
    OFFSET_INDEX_HEADERS,
    OFFSET_NEXT_AUTONUMBER,
    OFFSET_ROW_COUNT,
    SIZE_REAL_INDEX_HEADER,
    TABLE_TYPE_SYSTEM,
    TYPE_BOOLEAN,
    TYPE_MEMO,
    ColumnDef,
    LogicalIndex,
    RealIndex,
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


@dataclass(frozen=True)
class RowId:
    """A row's home slot: the page and slot number index entries use."""

    page: int
    slot: int


class Index:
    """A named index: its key columns and the B-tree that orders the rows."""

    def __init__(self, table: Table, logical: LogicalIndex, real: RealIndex) -> None:
        self.table = table
        self.logical = logical
        self.real = real
        self.name = logical.name
        self.columns: list[tuple[ColumnDef, bool]] = [
            (table.definition.column_by_number(c.number), c.ascending) for c in real.columns
        ]

    @property
    def column_names(self) -> list[str]:
        return [c.name for c, _ in self.columns]

    @property
    def unique(self) -> bool:
        return self.real.unique

    @property
    def is_primary_key(self) -> bool:
        return self.logical.is_primary_key

    @property
    def ignores_nulls(self) -> bool:
        return bool(self.real.flags & INDEX_IGNORE_NULLS)

    @property
    def distinct_count(self) -> int:
        """The engine's own count of distinct keys, null counting as one."""
        return self.real.entry_count

    def entries(self) -> Iterator[tuple[list[object], int, int]]:
        """``(key values, page, row)`` for every leaf entry in key order.
        Text columns come back as :class:`~pyopenvba.access._index.TextKey`."""
        store = self.table.database.store
        for entry in leaf_entries(store, self.real.root_page):
            yield decode_key(entry.key, self.columns), entry.page, entry.row

    def rows(self) -> Iterator[dict[str, object]]:
        """The table's rows in this index's order."""
        for _key, page, row in self.entries():
            data = self.table.fetch_row(page, row)
            if data is None:
                raise AccessError(f"index {self.name!r} points at dead row ({page}, {row})")
            yield self.table.decode(split_row(self.table.definition, data))

    def key_for(self, values: dict[str, object]) -> bytes | None:
        """The stored key for a row with these values, or ``None`` when
        the index ignores nulls and every key column is null."""
        key_values = [values.get(c.name) for c, _ in self.columns]
        if self.ignores_nulls and all(v is None for v in key_values):
            return None
        return encode_key(key_values, self.columns)


class Table:
    """A table: its definition plus the rows on its owned pages."""

    def __init__(self, database: AccessDatabase, definition: TableDefinition, name: str) -> None:
        self.database = database
        self._db = database
        self.definition = definition
        self.name = name

    # -- structure -------------------------------------------------------------

    @property
    def columns(self) -> list[ColumnDef]:
        return self.definition.columns_by_number()

    @property
    def column_names(self) -> list[str]:
        return [c.name for c in self.columns]

    @property
    def row_count(self) -> int:
        return self.definition.row_count

    @property
    def indexes(self) -> list[Index]:
        d = self.definition
        return [Index(self, li, d.real_indexes[li.real_index]) for li in d.logical_indexes]

    def index(self, name: str) -> Index:
        for index in self.indexes:
            if index.name.lower() == name.lower():
                return index
        raise AccessError(f"table {self.name!r} has no index named {name!r}")

    @property
    def primary_key(self) -> Index | None:
        for index in self.indexes:
            if index.is_primary_key:
                return index
        return None

    def _real_indexes(self) -> list[tuple[int, RealIndex, list[tuple[ColumnDef, bool]]]]:
        d = self.definition
        return [
            (i, real, [(d.column_by_number(c.number), c.ascending) for c in real.columns])
            for i, real in enumerate(d.real_indexes)
        ]

    # -- reading -----------------------------------------------------------------

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

    def fetch_row(self, page: int, slot: int) -> bytes | None:
        """The bytes of the row whose home is ``(page, slot)``, following an
        overflow pointer to where the row now lives; ``None`` if the slot
        is dead.  Index entries name rows by their home slot, which does
        not change when a grown row is moved."""
        store = self._db.store
        raw_page = store.read(page)
        slots = row_slots(raw_page)
        if not 0 <= slot < len(slots):
            raise AccessError(f"page {page} has no slot {slot}")
        entry = slots[slot]
        data = row_bytes(raw_page, slot)
        if data is None:
            return None
        if entry & ROW_OVERFLOW:
            target_row, target_page = row_pointer(data)
            target = row_bytes(store.read(target_page), target_row, overflow_target=True)
            if target is None:
                raise AccessError(f"overflow row ({page}, {slot}) points at nothing")
            return target
        return data

    def raw_rows(self) -> Iterator[tuple[int, int, bytes]]:
        """``(page, slot, row_bytes)`` for every live row, keyed by the row's
        home slot, with overflow rows followed to where they live."""
        store = self._db.store
        for page in self.data_pages():
            slots = row_slots(store.read(page))
            for slot, entry in enumerate(slots):
                if entry & ROW_DELETED:
                    continue
                data = self.fetch_row(page, slot)
                if data is not None:
                    yield page, slot, data

    def rows(self) -> Iterator[dict[str, object]]:
        for _page, _slot, data in self.raw_rows():
            yield self.decode(split_row(self.definition, data))

    def rows_with_ids(self) -> Iterator[tuple[RowId, dict[str, object]]]:
        for page, slot, data in self.raw_rows():
            yield RowId(page, slot), self.decode(split_row(self.definition, data))

    def decode(self, raw: RawRow) -> dict[str, object]:
        out: dict[str, object] = {}
        for column in self.columns:
            value = raw.values.get(column.number)
            if value is None:
                if column.type_code == TYPE_BOOLEAN and raw.present.get(column.number) is False:
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

    # -- writing -----------------------------------------------------------------

    def _encode_values(
        self, values: dict[str, object], keep_raw: RawRow | None = None
    ) -> tuple[dict[int, bytes | None], set[int]]:
        """Encode Python values per column number.  ``keep_raw`` is the
        row's existing bytes: long-value columns not being changed keep
        theirs, since those are not rewritten yet."""
        d = self.definition
        known = {c.name.lower(): c for c in d.columns}
        for name in values:
            if name.lower() not in known:
                raise AccessError(f"table {self.name!r} has no column {name!r}")
        encoded: dict[int, bytes | None] = {}
        booleans: set[int] = set()
        compress_system = d.table_type == TABLE_TYPE_SYSTEM
        for column in d.columns:
            given = column.name in values or any(k.lower() == column.name.lower() for k in values)
            value = values.get(column.name)
            if value is None:
                for key, candidate in values.items():
                    if key.lower() == column.name.lower():
                        value = candidate
            if column.type_code == TYPE_BOOLEAN:
                if value:
                    booleans.add(column.number)
                continue
            if column.is_long_value:
                if keep_raw is not None and not given:
                    encoded[column.number] = keep_raw.values.get(column.number)
                    continue
                if value is None:
                    encoded[column.number] = None
                    continue
                raise AccessError(
                    f"column {column.name!r}: writing Memo/OLE values is not supported yet"
                )
            if value is None:
                encoded[column.number] = None
                continue
            encoded[column.number] = encode_scalar(
                column, value, compress_text=column.compressed_unicode or compress_system
            )
        return encoded, booleans

    def insert_row(self, values: dict[str, object]) -> RowId:
        """Add a row.  AutoNumber columns are assigned; every other column
        not given is null.  Returns the row's home slot."""
        db = self._db
        d = self.definition
        values = dict(values)
        for column in d.columns:
            if column.auto_number and values.get(column.name) is None:
                values[column.name] = d.next_autonumber + 1
                d.next_autonumber += 1
                db.patch_definition(d, OFFSET_NEXT_AUTONUMBER, struct.pack("<I", d.next_autonumber & 0xFFFFFFFF))
        encoded, booleans = self._encode_values(values)
        row = encode_row(d, encoded, booleans)
        page_number = self._page_with_room(len(row))
        page = DataPage(db.store.read(page_number))
        slot = page.add_row(row)
        db.store.write(page_number, page.to_bytes())
        full_values = self.decode(split_row(d, row))
        for i, real, columns in self._real_indexes():
            key = self._key(real, columns, full_values)
            if key is None:
                continue
            distinct = self._btree(i, real).insert(key, page_number, slot)
            if distinct:
                real.entry_count += 1
                db.patch_definition(
                    d,
                    OFFSET_INDEX_HEADERS + i * SIZE_REAL_INDEX_HEADER + 4,
                    struct.pack("<I", real.entry_count),
                )
        d.row_count += 1
        db.patch_definition(d, OFFSET_ROW_COUNT, struct.pack("<I", d.row_count))
        return RowId(page_number, slot)

    def delete_row(self, row_id: RowId) -> None:
        db = self._db
        d = self.definition
        data = self.fetch_row(row_id.page, row_id.slot)
        if data is None:
            raise AccessError(f"row ({row_id.page}, {row_id.slot}) is not live")
        entry = row_slots(db.store.read(row_id.page))[row_id.slot]
        if entry & ROW_OVERFLOW:
            raise AccessError("deleting a row that moved to an overflow page is not supported yet")
        values = self.decode(split_row(d, data))
        for i, real, columns in self._real_indexes():
            key = self._key(real, columns, values)
            if key is not None:
                self._btree(i, real).delete(key, row_id.page, row_id.slot)
        page = DataPage(db.store.read(row_id.page))
        page.remove_row(row_id.slot)
        db.store.write(row_id.page, page.to_bytes())
        d.row_count -= 1
        db.patch_definition(d, OFFSET_ROW_COUNT, struct.pack("<I", d.row_count))

    def update_row(self, row_id: RowId, changes: dict[str, object]) -> None:
        """Change the given columns of one row; the rest keep their values."""
        db = self._db
        d = self.definition
        data = self.fetch_row(row_id.page, row_id.slot)
        if data is None:
            raise AccessError(f"row ({row_id.page}, {row_id.slot}) is not live")
        entry = row_slots(db.store.read(row_id.page))[row_id.slot]
        if entry & ROW_OVERFLOW:
            raise AccessError("updating a row that moved to an overflow page is not supported yet")
        parts = split_row(d, data)
        old_values = self.decode(parts)
        new_values = dict(old_values)
        changed: dict[str, object] = {}
        for name, value in changes.items():
            matched = [c for c in d.columns if c.name.lower() == name.lower()]
            if not matched:
                raise AccessError(f"table {self.name!r} has no column {name!r}")
            new_values[matched[0].name] = value
            changed[matched[0].name] = value
        # Unchanged long values keep their stored bytes; everything else is
        # re-encoded from the decoded values.
        for column in d.columns:
            if column.is_long_value and column.name not in changed:
                new_values.pop(column.name, None)
        encoded, booleans = self._encode_values(new_values, keep_raw=parts)
        row = encode_row(d, encoded, booleans)
        page = DataPage(db.store.read(row_id.page))
        start, end = page.span(row_id.slot)
        if len(row) - (end - start) > page.free_space:
            raise AccessError("the updated row no longer fits its page; overflow rows are not written yet")
        for i, real, columns in self._real_indexes():
            old_key = self._key(real, columns, old_values)
            new_key = self._key(real, columns, new_values)
            if old_key == new_key:
                continue
            tree = self._btree(i, real)
            if old_key is not None:
                tree.delete(old_key, row_id.page, row_id.slot)
            if new_key is not None and tree.insert(new_key, row_id.page, row_id.slot):
                real.entry_count += 1
                db.patch_definition(
                    d, OFFSET_INDEX_HEADERS + i * SIZE_REAL_INDEX_HEADER + 4, struct.pack("<I", real.entry_count)
                )
        page = DataPage(db.store.read(row_id.page))
        page.replace_row(row_id.slot, row)
        db.store.write(row_id.page, page.to_bytes())

    def _key(self, real: RealIndex, columns: list[tuple[ColumnDef, bool]], values: dict[str, object]) -> bytes | None:
        key_values = [values.get(c.name) for c, _ in columns]
        if real.flags & INDEX_IGNORE_NULLS and all(v is None for v in key_values):
            return None
        return encode_key(key_values, columns)

    def _btree(self, position: int, real: RealIndex) -> BTree:
        store = self._db.store

        def allocate() -> int:
            page = allocate_page(store)
            add_to_map(store, real.usage_map_ref, page)
            return page

        return BTree(store, real.root_page, self.definition.page, allocate)

    def _page_with_room(self, row_length: int) -> int:
        """A data page that can take the row, the way the engine picks one:
        the free-space map's pages in order, dropping any that cannot hold
        it, else a fresh page registered with both maps."""
        db = self._db
        d = self.definition
        free_map = read_usage_map_ref(db.store, d.free_space_pages_ref)
        for candidate in free_map.pages():
            if candidate >= db.store.page_count:
                continue
            raw = db.store.read(candidate)
            if raw[0] != PAGE_DATA or page_owner(raw) != d.page:
                continue
            if DataPage(raw).fits(row_length):
                return candidate
            remove_from_map(db.store, d.free_space_pages_ref, candidate)
        page = allocate_page(db.store)
        db.store.write(page, DataPage.new(d.page).to_bytes())
        add_to_map(db.store, d.owned_pages_ref, page)
        add_to_map(db.store, d.free_space_pages_ref, page)
        return page


class AccessDatabase:
    """Open a ``.accdb`` / Jet 4 ``.mdb`` and read or edit its tables."""

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

    # -- persistence -------------------------------------------------------------

    def to_bytes(self) -> bytes:
        return self.store.to_bytes()

    def save(self, path: str | Path | None = None) -> Path:
        target = Path(path) if path is not None else self.path
        if target is None:
            raise AccessError("no path to save to; the database was opened from bytes")
        target.write_bytes(self.to_bytes())
        return target

    def patch_definition(self, definition: TableDefinition, offset: int, data: bytes) -> None:
        """Overwrite bytes of a table definition's header, which always
        lies on its first page."""
        if offset + len(data) > 0x3F + len(definition.real_indexes) * 12:
            raise AccessError("definition patches are limited to the fixed header and index headers")
        raw = bytearray(self.store.read(definition.page))
        raw[offset : offset + len(data)] = data
        self.store.write(definition.page, bytes(raw))

    # -- catalog -----------------------------------------------------------------

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
