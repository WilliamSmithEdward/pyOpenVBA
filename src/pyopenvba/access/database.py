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
from pyopenvba.access._index import decode_key, encode_key, leaf_entries, OFFSET_ENTRIES
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
    definition_page_count,
    definition_pages,
    empty_index_root,
    foreign_key_logical,
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
    TYPE_DATETIME,
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

# MSysRelationships.grbit, DAO's RelationAttributeEnum.
RELATION_UNIQUE = 0x1
RELATION_DONT_ENFORCE = 0x2
RELATION_UPDATE_CASCADE = 0x100
RELATION_DELETE_CASCADE = 0x1000
# The three permission rows the engine gives a relationship object.
RELATIONSHIP_ACMS = (0xF00FE, 0xFFFFF, 0xFFFFF)

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
    #: The two stamps as the stored serials (days since 1899-12-30); a
    #: datetime cannot carry the last bit of the double, these can.
    date_create_serial: float | None = None
    date_update_serial: float | None = None

    @property
    def is_system(self) -> bool:
        return bool(self.flags & FLAG_SYSTEM) or self.name.startswith("MSys")

    @property
    def is_table(self) -> bool:
        return self.type in (OBJECT_TABLE, OBJECT_LINKED_TABLE)


@dataclass(frozen=True)
class Relationship:
    """A relationship as MSysRelationships describes it: the referencing
    table's columns against the referenced table's, plus DAO's attribute
    bits (cascades, enforcement, one-to-one)."""

    name: str
    table: str
    columns: tuple[str, ...]
    referenced_table: str
    referenced_columns: tuple[str, ...]
    attributes: int

    @property
    def cascade_updates(self) -> bool:
        return bool(self.attributes & RELATION_UPDATE_CASCADE)

    @property
    def cascade_deletes(self) -> bool:
        return bool(self.attributes & RELATION_DELETE_CASCADE)

    @property
    def enforced(self) -> bool:
        return not self.attributes & RELATION_DONT_ENFORCE


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
        values.  ``keep_raw`` is the row's existing bytes: columns not
        named in ``values`` keep theirs exactly (a stamp re-encoded from
        its datetime could lose its last bit)."""
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
                if (keep_raw.present.get(column.number) if keep_raw is not None and not given else value):
                    booleans.add(column.number)
                continue
            if keep_raw is not None and not given:
                encoded[column.number] = keep_raw.values.get(column.number)
                continue
            if column.is_long_value:
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
        full_values = self._exact_values(split_row(d, row))
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
        """Delete one row with its index entries and long values, settling
        the pages it leaves as the engine does (see :meth:`_row_removed`)."""
        db = self._db
        d = self.definition
        data = self.fetch_row(row_id.page, row_id.slot)
        if data is None:
            raise AccessError(f"row ({row_id.page}, {row_id.slot}) is not live")
        parts = split_row(d, data)
        values = self._exact_values(parts)
        for i, real, columns in self._real_indexes():
            key = self._key(real, columns, values)
            if key is not None:
                self._btree(i, real).delete(key, row_id.page, row_id.slot)
        self._free_long_values(parts)
        moved = self._moved_to(row_id)
        if moved is not None:
            # The page that held the overflow copy is only written back: the
            # engine neither retires it when it empties nor re-lists it.
            target = DataPage(db.store.read(moved[0]))
            target.remove_row(moved[1], overflow_target=True)
            db.store.write(moved[0], target.to_bytes())
        page = DataPage(db.store.read(row_id.page))
        page.remove_row(row_id.slot)
        # A home slot that held only the 4-byte pointer to a moved row is
        # written back without re-listing the page (measured: the engine
        # re-lists after a 15-byte row goes, not after a pointer).
        self._row_removed(row_id.page, page, settle=moved is None)
        d.row_count -= 1
        db.patch_definition(d, OFFSET_ROW_COUNT, struct.pack("<I", d.row_count))

    def _row_removed(self, page_number: int, page: DataPage, *, settle: bool = True) -> None:
        """Write back a page that just lost a row, the way the engine
        settles it: a page with rows left rejoins the free-space map; an
        emptied page is retired (type 0x09, released, out of both maps),
        unless it is the table's first data page, which stays.  Without
        ``settle`` (the row was only a pointer) the page is written back
        and nothing else moves."""
        db = self._db
        d = self.definition
        if not settle:
            db.store.write(page_number, page.to_bytes())
            return
        if page.live_rows == 0 and page_number != min(read_usage_map_ref(db.store, d.owned_pages_ref).pages(), default=page_number):
            page.retire()
            db.store.write(page_number, page.to_bytes())
            release_page(db.store, page_number)
            remove_from_map(db.store, d.owned_pages_ref, page_number)
            remove_from_map(db.store, d.free_space_pages_ref, page_number)
            return
        db.store.write(page_number, page.to_bytes())
        add_to_map(db.store, d.free_space_pages_ref, page_number)

    def truncate(self) -> None:
        """Delete every row the way ``DELETE FROM t`` without a filter does:
        the data and long-value pages are released untouched, the table's
        and columns' usage maps are emptied, each index is reset to an
        empty root with a distinct count of 0, and the AutoNumber counter
        keeps its place."""
        db = self._db
        store = db.store
        d = self.definition
        for page_number in read_usage_map_ref(store, d.owned_pages_ref).pages():
            release_page(store, page_number)
            remove_from_map(store, d.owned_pages_ref, page_number)
            remove_from_map(store, d.free_space_pages_ref, page_number)
        for owned_ref, free_ref in d.column_usage_maps.values():
            for page_number in read_usage_map_ref(store, owned_ref).pages():
                release_page(store, page_number)
                remove_from_map(store, owned_ref, page_number)
                remove_from_map(store, free_ref, page_number)
        for i, real in enumerate(d.real_indexes):
            for page_number in read_usage_map_ref(store, real.usage_map_ref).pages():
                if page_number != real.root_page:
                    release_page(store, page_number)
                    remove_from_map(store, real.usage_map_ref, page_number)
            # The root is reset in place: header and mask as an empty leaf,
            # the old entries' bytes left where they were.
            store.write(real.root_page, empty_index_root(d.page)[:OFFSET_ENTRIES] + store.read(real.root_page)[OFFSET_ENTRIES:])
            real.entry_count = 0
            db.patch_definition(d, OFFSET_INDEX_HEADERS + i * SIZE_REAL_INDEX_HEADER + 4, struct.pack("<I", 0))
        d.row_count = 0
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
        old_values = self._exact_values(parts)
        new_values = dict(old_values)
        given: dict[str, object] = {}
        for name, value in changes.items():
            matched = [c for c in d.columns if c.name.lower() == name.lower()]
            if not matched:
                raise AccessError(f"table {self.name!r} has no column {name!r}")
            new_values[matched[0].name] = value
            given[matched[0].name] = value
        # Untouched columns keep their stored bytes.  A changed long value
        # is stored first and its old storage given back afterwards, the
        # engine's order (measured: the new value went to another page
        # although the old one's page would have had room once freed).
        self._check_unique(new_values, exclude=row_id)
        encoded, booleans = self._encode_values(given, keep_raw=parts)
        self._free_long_values(parts, only=set(given))
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
            # Unlike a delete, a row coming home retires the copy's page
            # when that empties it (measured: type 0x09, released, out of
            # the maps).
            self._row_removed(target_page, target)
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

    def _exact_values(self, parts: RawRow) -> dict[str, object]:
        """The row's values for key building and re-encoding: decoded, except
        that a DateTime is its stored serial (a datetime can lose the last
        bit of the double, and the index key must match the row)."""
        values = self.decode(parts)
        for column in self.columns:
            raw = parts.values.get(column.number)
            if column.type_code == TYPE_DATETIME and isinstance(raw, bytes) and len(raw) == 8:
                values[column.name] = struct.unpack("<d", raw)[0]
        return values

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

    @classmethod
    def create_new(cls, path: str | Path) -> AccessDatabase:
        """Write a blank ``.accdb`` at ``path`` and open it.  The bytes are
        a template Access authored itself (the same one
        :meth:`pyopenvba.AccessReader.create_new` uses, holding one empty
        module), so the engine and Access open the result as their own.
        ``path`` is overwritten if it exists."""
        from pyopenvba._templates import EMPTY_ACCDB_BYTES

        target = Path(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(EMPTY_ACCDB_BYTES)
        return cls(target)

    # -- persistence -------------------------------------------------------------

    def to_bytes(self) -> bytes:
        return self.store.to_bytes()

    def save(self, path: str | Path | None = None) -> Path:
        target = Path(path) if path is not None else self.path
        if target is None:
            raise AccessError("no path to save to; the database was opened from bytes")
        target.write_bytes(self.to_bytes())
        return target

    def _write_definition(self, stream: bytes, first: int, old_chain: Sequence[int]) -> list[int]:
        """Lay a definition stream over ``first`` and a continuation chain.
        As the engine does on every rewrite, the continuation pages are
        allocated afresh in ascending order, chained in reverse, and only
        then are the pages of ``old_chain`` released (their bytes kept).
        Returns the chain written."""
        store = self.store
        fresh = [allocate_page(store) for _ in range(definition_page_count(len(stream)) - 1)]
        chain = fresh[::-1]
        for page, image in zip([first, *chain], definition_pages(stream, chain), strict=True):
            store.write(page, image)
        for page in old_chain:
            release_page(store, page)
        return chain

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
                parts = split_row(table.definition, data)
                values = table.decode(parts)
                owner = values.get("Owner")
                serials = [_stamp_serial(parts, table.definition.column(c).number) for c in ("DateCreate", "DateUpdate")]
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
                        date_create_serial=serials[0],
                        date_update_serial=serials[1],
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
        return self._container("Tables")

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
        updated: object | None = None,
    ) -> Table:
        """Create a table the way the engine does: a definition page, a page
        of usage maps, an empty root per index, and the catalog rows.
        ``created`` and ``updated`` are the catalog row's two timestamps
        (now, and the creation time, by default), as datetimes or as the
        stored serials; the engine stamps DateUpdate when the definition
        is complete, so on a large table it runs a little after
        DateCreate."""
        import datetime as _dt

        specs = list(columns)
        index_specs = list(indexes or [])
        if any(e.name.lower() == name.lower() for e in self.catalog()):
            raise AccessError(f"an object named {name!r} already exists")
        store = self.store
        tag = self.definition(MSYS_OBJECTS_PAGE).tag
        long_columns = [n for n, c in enumerate(specs) if c.type_code in (TYPE_MEMO, TYPE_OLE)]

        # Pages are taken in the engine's order: the definition, the
        # usage-map page, whatever the catalog rows need, index roots, then
        # the definition's continuation pages if it runs past one page.
        definition_page = allocate_page(store)
        map_page = allocate_page(store)
        map_rows = 2 + len(index_specs) + 2 * len(long_columns)
        store.write(map_page, usage_map_page(map_rows))

        when = created if isinstance(created, (_dt.datetime, float)) else _dt.datetime.now().replace(microsecond=0)
        when_updated = updated if isinstance(updated, (_dt.datetime, float)) else when
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
        # The owner arrives with the second step, and so does the final
        # DateUpdate: the first version of the row, whose bytes stay below
        # the slot table after the move, carries DateCreate twice.
        objects.update_row(catalog_row, {"Owner": self._default_owner(), "DateUpdate": when_updated})
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
        self._write_definition(build_definition(specs, index_specs, layout), definition_page, [])
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
        self._write_definition(serialize_definition(d), d.page, d.pages[1:])
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
        when = updated if isinstance(updated, (_dt.datetime, float)) else _dt.datetime.now().replace(microsecond=0)
        objects = self.table("MSysObjects")
        for rid, row in objects.rows_with_ids():
            if row["Id"] == d.page and row["Type"] == OBJECT_TABLE:
                objects.update_row(rid, {"DateUpdate": when})
                break
        self._catalog = None
        self._definitions.pop(d.page, None)
        return self.table(table_name).index(spec.name)

    def relationships(self) -> list[Relationship]:
        """Every relationship in MSysRelationships, columns in pair order."""
        groups: dict[str, list[dict[str, object]]] = {}
        for row in self.table("MSysRelationships").rows():
            groups.setdefault(str(row["szRelationship"]), []).append(row)
        out: list[Relationship] = []
        for name, rows in groups.items():
            rows.sort(key=lambda r: _as_int(r["icolumn"]))
            first = rows[0]
            out.append(
                Relationship(
                    name=name,
                    table=str(first["szObject"]),
                    columns=tuple(str(r["szColumn"]) for r in rows),
                    referenced_table=str(first["szReferencedObject"]),
                    referenced_columns=tuple(str(r["szReferencedColumn"]) for r in rows),
                    attributes=_as_int(first["grbit"]),
                )
            )
        return out

    def create_relationship(
        self,
        name: str,
        table: str,
        columns: Sequence[str],
        referenced_table: str,
        referenced_columns: Sequence[str],
        *,
        cascade_updates: bool = False,
        cascade_deletes: bool = False,
        created: object | None = None,
        table_updated: object | None = None,
        referenced_updated: object | None = None,
    ) -> Relationship:
        """Relate ``table`` to ``referenced_table`` as ``ALTER TABLE ... ADD
        CONSTRAINT ... FOREIGN KEY`` does: a non-unique index named after
        the relationship on the referencing columns, a foreign-key logical
        entry on each side pointing at the other's definition page and
        logical index number (the referenced side's is named ``.r`` plus a
        letter from its index number and shares the unique index the
        foreign key refers to), one MSysRelationships row per column pair,
        a catalog object of type 8 with its three permission rows, and a
        new DateUpdate on both tables.  The three stamps default to now;
        pass datetimes or stored serials."""
        import datetime as _dt

        columns = tuple(columns)
        referenced_columns = tuple(referenced_columns)
        if not columns or len(columns) != len(referenced_columns):
            raise AccessError("a relationship pairs one or more columns with as many referenced columns")
        if any(r.name.lower() == name.lower() for r in self.relationships()):
            raise AccessError(f"a relationship named {name!r} already exists")
        if any(e.name.lower() == name.lower() for e in self.catalog()):
            raise AccessError(f"an object named {name!r} already exists")
        child = self.table(table)
        parent = self.table(referenced_table)
        cd, pd = child.definition, parent.definition
        for col in columns:
            if not any(c.name.lower() == col.lower() for c in cd.columns):
                raise AccessError(f"table {table!r} has no column {col!r}")
        wanted = [pd.column(c).number for c in referenced_columns]
        parent_real = next(
            (i for i, real in enumerate(pd.real_indexes) if real.unique and [c.number for c in real.columns] == wanted),
            None,
        )
        if parent_real is None:
            raise AccessError(f"table {referenced_table!r} has no unique index on {', '.join(referenced_columns)}")
        if any(li.name.lower() == name.lower() for li in cd.logical_indexes):
            raise AccessError(f"table {table!r} already has an index named {name!r}")
        attributes = (RELATION_UPDATE_CASCADE if cascade_updates else 0) | (RELATION_DELETE_CASCADE if cascade_deletes else 0)
        now = _dt.datetime.now().replace(microsecond=0)
        when = created if isinstance(created, (_dt.datetime, float)) else now
        child_when = table_updated if isinstance(table_updated, (_dt.datetime, float)) else when
        parent_when = referenced_updated if isinstance(referenced_updated, (_dt.datetime, float)) else when
        store = self.store
        child_logical_number = len(cd.logical_indexes)
        parent_logical_number = len(pd.logical_indexes)

        # The foreign-key index on the referencing table, built like CREATE INDEX.
        map_page = cd.owned_pages_ref >> 8
        maps = DataPage(store.read(map_page))
        if not maps.fits(len(USAGE_MAP_ROW)):
            raise AccessError("the table's usage-map page is full; a second map page is not written yet")
        umap_row = maps.add_row(USAGE_MAP_ROW)
        store.write(map_page, maps.to_bytes())
        root = allocate_page(store)
        store.write(root, empty_index_root(cd.page))
        umap_ref = (map_page << 8) | umap_row
        real, _plain = new_index_parts(IndexSpec(name, columns), cd, len(cd.real_indexes), umap_ref, root)
        cd.real_indexes.append(real)
        cd.logical_indexes.append(
            foreign_key_logical(
                cd.tag, name, child_logical_number, len(cd.real_indexes) - 1,
                referencing=True, other_logical=parent_logical_number, other_page=pd.page,
                cascade_updates=cascade_updates, cascade_deletes=cascade_deletes,
            )
        )
        cd.real_index_count += 1
        cd.logical_index_count += 1
        self._write_definition(serialize_definition(cd), cd.page, cd.pages[1:])
        add_to_map(store, umap_ref, root)
        self._definitions.pop(cd.page, None)
        child = self.table(table)
        position = len(child.definition.real_indexes) - 1
        real = child.definition.real_indexes[position]
        key_columns = [(child.definition.column_by_number(c.number), c.ascending) for c in real.columns]
        tree = child._btree(position, real)  # pyright: ignore[reportPrivateUsage]
        distinct = 0
        for row_id, values in child.rows_with_ids():
            key = child._key(real, key_columns, values)  # pyright: ignore[reportPrivateUsage]
            if key is not None and tree.insert(key, row_id.page, row_id.slot):
                distinct += 1
        if distinct:
            real.entry_count = distinct
            self.patch_definition(child.definition, OFFSET_INDEX_HEADERS + position * SIZE_REAL_INDEX_HEADER + 4, struct.pack("<I", distinct))

        # The referenced side: a ``.r<letter>`` entry sharing the unique index.
        pd.logical_indexes.append(
            foreign_key_logical(
                pd.tag, ".r" + chr(0x41 + parent_logical_number), parent_logical_number, parent_real,
                referencing=False, other_logical=child_logical_number, other_page=cd.page,
                cascade_updates=cascade_updates, cascade_deletes=cascade_deletes,
            )
        )
        pd.logical_index_count += 1
        self._write_definition(serialize_definition(pd), pd.page, pd.pages[1:])
        self._definitions.pop(pd.page, None)

        # The relationship rows, the catalog object, its permissions, the stamps.
        relationships = self.table("MSysRelationships")
        for i, (col, ref) in enumerate(zip(columns, referenced_columns, strict=True)):
            relationships.insert_row(
                {
                    "szRelationship": name, "grbit": attributes, "ccolumn": len(columns), "icolumn": i,
                    "szObject": table, "szColumn": col, "szReferencedObject": referenced_table, "szReferencedColumn": ref,
                }
            )
        objects = self.table("MSysObjects")
        object_id = self._next_object_id()
        catalog_row = objects.insert_row(
            {
                "Id": object_id,
                "ParentId": self._container("Relationships").id,
                "Name": name,
                "Type": OBJECT_RELATIONSHIP,
                "Flags": 0,
                "DateCreate": when,
                "DateUpdate": when,
            }
        )
        objects.update_row(catalog_row, {"Owner": self._default_owner()})
        aces = self.table("MSysACEs")
        for ace, acm in zip(self._default_aces(), RELATIONSHIP_ACMS, strict=True):
            aces.insert_row({"SID": ace["SID"], "ACM": acm, "FInheritable": False, "ObjectId": object_id})
        for definition_page, stamp in ((cd.page, child_when), (pd.page, parent_when)):
            for rid, row in objects.rows_with_ids():
                if row["Id"] == definition_page and row["Type"] == OBJECT_TABLE:
                    objects.update_row(rid, {"DateUpdate": stamp})
                    break
        self._catalog = None
        return Relationship(name, table, columns, referenced_table, referenced_columns, attributes)

    def _next_object_id(self) -> int:
        """The id the engine gives the next Access-layer object: one past the
        highest of the negative ids in the catalog."""
        ids = [e.id for e in self.catalog() if e.id < 0]
        return max(ids) + 1 if ids else -0x80000000

    def _container(self, name: str) -> CatalogEntry:
        for entry in self.catalog():
            if entry.type == 3 and entry.name == name:
                return entry
        raise AccessError(f"the catalog has no {name} container")

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
        # Only the first definition page is marked freed; continuation
        # pages keep their bytes and just go back to the free map.
        store.write(d.page, mark_definition_freed(store.read(d.page)))
        released.update(d.pages)
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


def _stamp_serial(parts: RawRow, column_number: int) -> float | None:
    """A DateTime column's stored double, untouched by datetime rounding."""
    raw = parts.values.get(column_number)
    return struct.unpack("<d", raw)[0] if isinstance(raw, bytes) and len(raw) == 8 else None


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
