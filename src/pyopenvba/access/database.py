"""``AccessDatabase``: a Jet 4 / ACE database opened as tables of rows.

This is the engine's facade.  It reads the catalog (``MSysObjects``) to
find tables by name, parses their definitions, walks their owned pages
to yield rows as plain Python values, and -- for tables without long
values -- inserts, updates and deletes rows the way the engine does,
indexes and counters included.  ``save()`` writes the pages back.
"""

from __future__ import annotations

import struct
from collections.abc import Iterator, Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path

from pyopenvba.access_read import AccessError
from pyopenvba.access._alloc import add_to_map, allocate_page, release_page, remove_from_map, set_usage_bit
from pyopenvba.access._btree import BTree
from pyopenvba.access._datapage import DataPage
from pyopenvba.access._index import decode_key, encode_key, leaf_entries
from pyopenvba.access._lval import (
    free_long_value,
    memo_bytes,
    read_long_value,
    write_long_value,
)
from pyopenvba.access._pages import (
    PAGE_DATA,
    DatabaseHeader,
    PageStore,
    encode_row_pointer,
    page_owner,
    read_usage_map_ref,
    row_bytes,
    row_pointer,
    row_slots,
    ROW_DELETED,
    ROW_OVERFLOW,
)
from pyopenvba.access._schema import (
    DEFAULT_ACM,
    USAGE_MAP_ROW,
    ColumnSpec,
    DefinitionLayout,
    IndexSpec,
    build_definition,
    empty_index_root,
    mark_definition_freed,
    new_index_parts,
    serialize_definition,
    usage_map_page,
)
from pyopenvba.access._rows import (
    LongValueRef,
    RawRow,
    decode_long_value_ref,
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
    TYPE_BOOLEAN,
    TYPE_MEMO,
    TYPE_OLE,
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

    def key_for(self, values: Mapping[str, object]) -> bytes | None:
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
        self, values: Mapping[str, object], keep_raw: RawRow | None = None
    ) -> tuple[dict[int, bytes | None], set[int]]:
        """Encode Python values per column number and store any long
        values.  ``keep_raw`` is the row's existing bytes: long-value
        columns not named in ``values`` keep theirs."""
        d = self.definition
        known = {c.name.lower(): c for c in d.columns}
        for name in values:
            if name.lower() not in known:
                raise AccessError(f"table {self.name!r} has no column {name!r}")
        encoded: dict[int, bytes | None] = {}
        booleans: set[int] = set()
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
                encoded[column.number] = self._store_long_value(column, value)
                continue
            if value is None:
                encoded[column.number] = None
                continue
            encoded[column.number] = encode_scalar(column, value, compress_text=column.compressed_unicode)
        return encoded, booleans

    def _long_value_maps(self, column: ColumnDef) -> tuple[int, int]:
        maps = self.definition.column_usage_maps.get(column.number)
        if maps is None:
            raise AccessError(f"column {column.name!r} has no long-value usage maps")
        return maps

    def _store_long_value(self, column: ColumnDef, value: object) -> bytes:
        if column.type_code == TYPE_MEMO:
            if not isinstance(value, str):
                raise AccessError(f"column {column.name!r}: {value!r} is not text")
            data = memo_bytes(value)
        else:
            if not isinstance(value, (bytes, bytearray)):
                raise AccessError(f"column {column.name!r}: {value!r} is not bytes")
            data = bytes(value)
        if not data:
            raise AccessError(f"column {column.name!r}: an empty long value is stored as null; pass None")
        return write_long_value(self._db.store, self._long_value_maps(column), data, self._db.lval_stamp)

    def _free_long_values(self, parts: RawRow, only: set[str] | None = None) -> None:
        for column in self.definition.columns:
            if not column.is_long_value or (only is not None and column.name not in only):
                continue
            raw = parts.values.get(column.number)
            if raw is None:
                continue
            free_long_value(self._db.store, self._long_value_maps(column), decode_long_value_ref(raw))

    def _check_unique(self, values: Mapping[str, object], exclude: RowId | None) -> None:
        """Refuse a key another row already holds in a unique index; nulls
        count as equal unless the index ignores them, as in the engine."""
        for _i, real, columns in self._real_indexes():
            if not real.unique:
                continue
            key = self._key(real, columns, values)
            if key is None:
                continue
            for entry in leaf_entries(self._db.store, real.root_page):
                if entry.key == key and (exclude is None or (entry.page, entry.row) != (exclude.page, exclude.slot)):
                    names = ", ".join(c.name for c, _ in columns)
                    raise AccessError(f"table {self.name!r}: duplicate value in unique index ({names})")
                if entry.key > key:
                    break

    def insert_row(self, values: Mapping[str, object]) -> RowId:
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
        self._check_unique({c.name: values.get(c.name) for c in d.columns}, exclude=None)
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

    def _moved_to(self, row_id: RowId) -> tuple[int, int] | None:
        """Where a row lives when its home slot is an overflow pointer."""
        raw_page = self._db.store.read(row_id.page)
        entry = row_slots(raw_page)[row_id.slot]
        if not entry & ROW_OVERFLOW or entry & ROW_DELETED:
            return None
        pointer = row_bytes(raw_page, row_id.slot)
        if pointer is None:
            return None
        target_row, target_page = row_pointer(pointer)
        return target_page, target_row

    def delete_row(self, row_id: RowId) -> None:
        db = self._db
        d = self.definition
        data = self.fetch_row(row_id.page, row_id.slot)
        if data is None:
            raise AccessError(f"row ({row_id.page}, {row_id.slot}) is not live")
        parts = split_row(d, data)
        values = self.decode(parts)
        for i, real, columns in self._real_indexes():
            key = self._key(real, columns, values)
            if key is not None:
                self._btree(i, real).delete(key, row_id.page, row_id.slot)
        self._free_long_values(parts)
        moved = self._moved_to(row_id)
        if moved is not None:
            target = DataPage(db.store.read(moved[0]))
            target.remove_row(moved[1], overflow_target=True)
            db.store.write(moved[0], target.to_bytes())
        page = DataPage(db.store.read(row_id.page))
        page.remove_row(row_id.slot)
        db.store.write(row_id.page, page.to_bytes())
        d.row_count -= 1
        db.patch_definition(d, OFFSET_ROW_COUNT, struct.pack("<I", d.row_count))

    def update_row(self, row_id: RowId, changes: Mapping[str, object]) -> None:
        """Change the given columns of one row; the rest keep their values.
        A row that no longer fits its page moves to another page behind an
        overflow pointer, and moves back when it fits again, as the engine
        does."""
        db = self._db
        d = self.definition
        data = self.fetch_row(row_id.page, row_id.slot)
        if data is None:
            raise AccessError(f"row ({row_id.page}, {row_id.slot}) is not live")
        parts = split_row(d, data)
        old_values = self.decode(parts)
        new_values = dict(old_values)
        changed: set[str] = set()
        for name, value in changes.items():
            matched = [c for c in d.columns if c.name.lower() == name.lower()]
            if not matched:
                raise AccessError(f"table {self.name!r} has no column {name!r}")
            new_values[matched[0].name] = value
            changed.add(matched[0].name)
        # Unchanged long values keep their stored bytes; changed ones give
        # back their old storage and are stored afresh.
        self._check_unique(new_values, exclude=row_id)
        for column in d.columns:
            if column.is_long_value and column.name not in changed:
                new_values.pop(column.name, None)
        self._free_long_values(parts, only=changed)
        encoded, booleans = self._encode_values(new_values, keep_raw=parts)
        row = encode_row(d, encoded, booleans)
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
        self._place_row(row_id, row)

    def _place_row(self, row_id: RowId, row: bytes) -> None:
        """Store a row's new bytes: in its home slot when they fit, else on
        another page behind an overflow pointer."""
        db = self._db
        home = DataPage(db.store.read(row_id.page))
        moved = self._moved_to(row_id)
        start, end = home.span(row_id.slot)
        if moved is None:
            # In place when the growth fits the page's free space (measured:
            # a two-byte growth stays with three bytes free and moves with
            # one); otherwise the row moves behind a pointer.
            if len(row) - (end - start) <= home.free_space:
                home.replace_row(row_id.slot, row)
                db.store.write(row_id.page, home.to_bytes())
                return
            # A page that could not take the growth leaves the free-space map.
            remove_from_map(db.store, self.definition.free_space_pages_ref, row_id.page)
            target_page = self._page_with_room(len(row), exclude=row_id.page)
            target = DataPage(db.store.read(target_page))
            target_slot = target.add_row(row, flags=ROW_DELETED)
            db.store.write(target_page, target.to_bytes())
            home = DataPage(db.store.read(row_id.page))
            home.replace_row(row_id.slot, encode_row_pointer(target_page, target_slot), flags=ROW_OVERFLOW)
            db.store.write(row_id.page, home.to_bytes())
            return
        target_page, target_slot = moved
        if len(row) - (end - start) <= home.free_space:
            # It fits at home again (the pointer's bytes count): bring it
            # back and drop the copy.
            home.replace_row(row_id.slot, row, flags=0)
            db.store.write(row_id.page, home.to_bytes())
            target = DataPage(db.store.read(target_page))
            target.remove_row(target_slot, overflow_target=True)
            db.store.write(target_page, target.to_bytes())
            return
        target = DataPage(db.store.read(target_page))
        t_start, t_end = target.span(target_slot)
        if len(row) - (t_end - t_start) <= target.free_space:
            target.replace_row(target_slot, row)
            db.store.write(target_page, target.to_bytes())
            return
        target.remove_row(target_slot, overflow_target=True)
        db.store.write(target_page, target.to_bytes())
        remove_from_map(db.store, self.definition.free_space_pages_ref, target_page)
        new_page = self._page_with_room(len(row), exclude=row_id.page)
        landing = DataPage(db.store.read(new_page))
        new_slot = landing.add_row(row, flags=ROW_DELETED)
        db.store.write(new_page, landing.to_bytes())
        home = DataPage(db.store.read(row_id.page))
        home.replace_row(row_id.slot, encode_row_pointer(new_page, new_slot), flags=ROW_OVERFLOW)
        db.store.write(row_id.page, home.to_bytes())

    def _key(self, real: RealIndex, columns: list[tuple[ColumnDef, bool]], values: Mapping[str, object]) -> bytes | None:
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

    def _page_with_room(self, row_length: int, exclude: int | None = None) -> int:
        """A data page that can take the row, the way the engine picks one:
        the free-space map's pages in order, dropping any that cannot hold
        it, else a fresh page registered with both maps."""
        db = self._db
        d = self.definition
        if row_length + 2 > DataPage.new(d.page).free_space:
            raise AccessError(f"a row of {row_length} bytes cannot fit a page")
        free_map = read_usage_map_ref(db.store, d.free_space_pages_ref)
        for candidate in free_map.pages():
            if candidate >= db.store.page_count or candidate == exclude:
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
        # The engine stamps a chained long value's definition and its first
        # page with one value per session; any value works if both match.
        self.lval_stamp = 0x00500000 | (self.store.page_count & 0xFFFF)

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

    # -- schema --------------------------------------------------------------------

    def _tables_container(self) -> CatalogEntry:
        for entry in self.catalog():
            if entry.type == 3 and entry.name == "Tables":
                return entry
        raise AccessError("the catalog has no Tables container")

    def _default_owner(self) -> bytes:
        owners = [e.owner for e in self.table_entries(include_system=True) if e.owner]
        if not owners:
            raise AccessError("no table row to take an owner SID from")
        return max(set(owners), key=owners.count)

    def _default_aces(self) -> list[dict[str, object]]:
        """The three permission rows the engine gives a new table: the SIDs
        every table in the file carries, at full access."""
        aces = self.table("MSysACEs")
        sids: list[bytes] = []
        for row in aces.rows():
            sid = row["SID"]
            if isinstance(sid, bytes) and row["ACM"] == DEFAULT_ACM and sid not in sids:
                sids.append(sid)
        if len(sids) < 3:
            for row in aces.rows():
                sid = row["SID"]
                if isinstance(sid, bytes) and sid not in sids and len(sid) in (2, 100):
                    sids.append(sid)
        sids = sids[:3]
        if len(sids) != 3:
            raise AccessError("could not find the database's three default SIDs")
        return [{"SID": sid, "ACM": DEFAULT_ACM, "FInheritable": False} for sid in sids]

    def create_table(
        self,
        name: str,
        columns: Sequence[ColumnSpec],
        indexes: Sequence[IndexSpec] | None = None,
        *,
        created: object | None = None,
    ) -> Table:
        """Create a table the way the engine does: a definition page, a page
        of usage maps, an empty root per index, and the catalog rows."""
        import datetime as _dt

        specs = list(columns)
        index_specs = list(indexes or [])
        if any(e.name.lower() == name.lower() for e in self.catalog()):
            raise AccessError(f"an object named {name!r} already exists")
        store = self.store
        tag = self.definition(MSYS_OBJECTS_PAGE).tag
        long_columns = [n for n, c in enumerate(specs) if c.type_code in (TYPE_MEMO, TYPE_OLE)]

        # Pages are taken in the engine's order: the definition, the
        # usage-map page, whatever the catalog rows need, then index roots.
        definition_page = allocate_page(store)
        map_page = allocate_page(store)
        map_rows = 2 + len(index_specs) + 2 * len(long_columns)
        store.write(map_page, usage_map_page(map_rows))
        # The definition page must parse as a table before the catalog can
        # point at it; a placeholder is written first and replaced below.
        placeholder_layout = DefinitionLayout(
            page=definition_page, tag=tag, owned_ref=(map_page << 8), free_ref=(map_page << 8) | 1,
            index_umap_refs=[], index_roots=[],
            column_map_refs=_long_value_map_refs(specs, map_page, len(index_specs)),
        )
        store.write(definition_page, build_definition(specs, [], placeholder_layout))

        when = created if isinstance(created, _dt.datetime) else _dt.datetime.now().replace(microsecond=0)
        objects = self.table("MSysObjects")
        # The engine writes the catalog row in two steps -- first without an
        # owner, then updating it with one -- which decides whether the row
        # stays on its page or moves; the same steps give the same pages.
        catalog_row = objects.insert_row(
            {
                "Id": definition_page,
                "ParentId": self._tables_container().id,
                "Name": name,
                "Type": OBJECT_TABLE,
                "Flags": 0,
                "DateCreate": when,
                "DateUpdate": when,
            }
        )
        objects.update_row(catalog_row, {"Owner": self._default_owner()})
        aces = self.table("MSysACEs")
        for ace in self._default_aces():
            aces.insert_row(dict(ace, ObjectId=definition_page))

        roots = [allocate_page(store) for _ in index_specs]
        for root in roots:
            store.write(root, empty_index_root(definition_page))
        layout = DefinitionLayout(
            page=definition_page,
            tag=tag,
            owned_ref=(map_page << 8) | 0,
            free_ref=(map_page << 8) | 1,
            index_umap_refs=[(map_page << 8) | (2 + i) for i in range(len(index_specs))],
            index_roots=roots,
        )
        layout.column_map_refs = _long_value_map_refs(specs, map_page, len(index_specs))
        store.write(definition_page, build_definition(specs, index_specs, layout))
        for i, root in enumerate(roots):
            add_to_map(store, layout.index_umap_refs[i], root)
        self._catalog = None
        self._definitions.pop(definition_page, None)
        return self.table(name)

    def create_index(self, table_name: str, spec: IndexSpec, *, updated: object | None = None) -> Index:
        """Add an index to an existing table as CREATE INDEX does: an empty
        root page, a usage-map row on the table's map page, the definition
        rewritten with the index appended, existing entries built, and the
        catalog row's DateUpdate stamped."""
        import datetime as _dt

        table = self.table(table_name)
        d = table.definition
        if any(li.name.lower() == spec.name.lower() for li in d.logical_indexes):
            raise AccessError(f"table {table_name!r} already has an index named {spec.name!r}")
        if spec.primary and d.primary_key() is not None:
            raise AccessError(f"table {table_name!r} already has a primary key")
        store = self.store
        map_page = d.owned_pages_ref >> 8
        maps = DataPage(store.read(map_page))
        if not maps.fits(len(USAGE_MAP_ROW)):
            raise AccessError("the table's usage-map page is full; a second map page is not written yet")
        umap_row = maps.add_row(USAGE_MAP_ROW)
        store.write(map_page, maps.to_bytes())
        root = allocate_page(store)
        store.write(root, empty_index_root(d.page))
        umap_ref = (map_page << 8) | umap_row
        real, logical = new_index_parts(spec, d, len(d.real_indexes), umap_ref, root)
        d.real_indexes.append(real)
        d.logical_indexes.append(logical)
        d.real_index_count += 1
        d.logical_index_count += 1
        store.write(d.page, serialize_definition(d))
        add_to_map(store, umap_ref, root)
        self._definitions.pop(d.page, None)
        table = self.table(table_name)
        # Existing rows get their entries, in home-slot order.
        position = len(table.definition.real_indexes) - 1
        real = table.definition.real_indexes[position]
        columns = [(table.definition.column_by_number(c.number), c.ascending) for c in real.columns]
        tree = table._btree(position, real)  # pyright: ignore[reportPrivateUsage]
        distinct = 0
        for row_id, values in table.rows_with_ids():
            key = table._key(real, columns, values)  # pyright: ignore[reportPrivateUsage]
            if key is None:
                continue
            if real.unique:
                table._check_unique(values, exclude=row_id)  # pyright: ignore[reportPrivateUsage]
            if tree.insert(key, row_id.page, row_id.slot):
                distinct += 1
        if distinct:
            real.entry_count = distinct
            self.patch_definition(
                table.definition, OFFSET_INDEX_HEADERS + position * SIZE_REAL_INDEX_HEADER + 4, struct.pack("<I", distinct)
            )
        when = updated if isinstance(updated, _dt.datetime) else _dt.datetime.now().replace(microsecond=0)
        objects = self.table("MSysObjects")
        for rid, row in objects.rows_with_ids():
            if row["Id"] == d.page and row["Type"] == OBJECT_TABLE:
                objects.update_row(rid, {"DateUpdate": when})
                break
        self._catalog = None
        self._definitions.pop(d.page, None)
        return self.table(table_name).index(spec.name)

    def drop_table(self, name: str) -> None:
        """Remove a table and give back everything it held, in the engine's
        order (read off the bytes it leaves behind): the owned-pages row
        goes first, then every index and long-value map has its pages
        released and its bits cleared, then the remaining map rows go,
        then the definition page is marked freed and every page released."""
        table = self.table(name)
        d = table.definition
        if d.is_system or name.startswith("MSys"):
            raise AccessError(f"{name!r} is a system table")
        store = self.store
        released: set[int] = set()

        def clear_map(ref: int) -> None:
            umap = read_usage_map_ref(store, ref)
            for page_number in umap.pages():
                released.add(page_number)
                set_usage_bit(store, umap, page_number, False)

        def kill_row(ref: int) -> None:
            page = DataPage(store.read(ref >> 8))
            page.remove_row(ref & 0xFF)
            store.write(ref >> 8, page.to_bytes())

        clear_map(d.owned_pages_ref)
        kill_row(d.owned_pages_ref)
        for real in d.real_indexes:
            clear_map(real.usage_map_ref)
        for owned_ref, free_ref in d.column_usage_maps.values():
            clear_map(owned_ref)
            clear_map(free_ref)
        kill_row(d.free_space_pages_ref)
        for real in d.real_indexes:
            kill_row(real.usage_map_ref)
        for owned_ref, free_ref in d.column_usage_maps.values():
            kill_row(owned_ref)
            kill_row(free_ref)
        map_pages = {d.owned_pages_ref >> 8, d.free_space_pages_ref >> 8}
        map_pages.update(r.usage_map_ref >> 8 for r in d.real_indexes)
        for owned_ref, free_ref in d.column_usage_maps.values():
            map_pages.update((owned_ref >> 8, free_ref >> 8))
        for map_page in map_pages:
            if all(entry & 0xC000 == 0xC000 for entry in DataPage(store.read(map_page)).slots):
                released.add(map_page)
        for page_number in d.pages:
            store.write(page_number, mark_definition_freed(store.read(page_number)))
            released.add(page_number)
        for page_number in sorted(p for p in released if p < store.page_count):
            release_page(store, page_number)

        objects = self.table("MSysObjects")
        for rid, row in objects.rows_with_ids():
            if row["Id"] == d.page and row["Type"] == OBJECT_TABLE:
                objects.delete_row(rid)
                break
        aces = self.table("MSysACEs")
        for rid, row in list(aces.rows_with_ids()):
            if row["ObjectId"] == d.page:
                aces.delete_row(rid)
        self._catalog = None
        self._definitions.pop(d.page, None)


def _long_value_map_refs(specs: Sequence[ColumnSpec], map_page: int, index_count: int) -> dict[int, tuple[int, int]]:
    """Usage-map references for a table's Memo/OLE columns: two rows each on
    the table's map page, after the owned, free-space and index rows."""
    refs: dict[int, tuple[int, int]] = {}
    next_row = 2 + index_count
    for number, spec in enumerate(specs):
        if spec.type_code in (TYPE_MEMO, TYPE_OLE):
            refs[number] = ((map_page << 8) | next_row, (map_page << 8) | (next_row + 1))
            next_row += 2
    return refs


def _as_int(value: object) -> int:
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, int):
        return value
    raise AccessError(f"expected an integer catalog field, got {type(value).__name__}")
