"""``AccessDatabase``: a Jet 4 / ACE database opened as tables of rows.

This is the engine's facade.  It reads the catalog (``MSysObjects``) to
find tables by name, parses their definitions, walks their owned pages
to yield rows as plain Python values, and -- for tables without long
values -- inserts, updates and deletes rows the way the engine does,
indexes and counters included.  ``save()`` writes the pages back.
"""

from __future__ import annotations

import datetime as dt
import random
import struct
from collections.abc import Callable, Generator, Iterator, Mapping, Sequence
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path

from pyopenvba.access._alloc import (
    GLOBAL_USAGE_MAP_PAGE,
    GLOBAL_USAGE_MAP_ROW,
    add_to_map,
    allocate_page,
    read_usage_map,
    release_page,
    remove_from_map,
    set_usage_bit,
)
from pyopenvba.access._btree import BTree
from pyopenvba.access._complex import (
    ATTACHMENT_COLUMNS,
    ATTACHMENT_TYPE,
    FLAT_TABLE_FLAGS,
    HAS_COMPLEX_COLUMN,
    MISC_ATTACHMENT,
    MISC_ELEMENT_ID,
    MISC_KEY,
    MISC_SCALAR,
    PRIMARY_INDEX,
    SCALAR_INDEX,
    SCALAR_TYPES,
    TYPE_TABLES,
    Attachment,
    ComplexColumn,
    decode_file_data,
    encode_file_data,
    flat_table_name,
    index_name,
    patch_column_header,
)
from pyopenvba.access._datapage import DataPage
from pyopenvba.access._index import OFFSET_ENTRIES, decode_key, encode_key, leaf_entries
from pyopenvba.access._lval import (
    free_long_value,
    memo_bytes,
    read_long_value,
    write_long_value,
)
from pyopenvba.access._designs import (
    set_property,
    CATALOG_CONTAINERS,
    add_control,
    CONTAINERS,
    NAV_TYPES,
    OBJECT_TYPES,
    AccessDesign,
    parse_design,
    template,
    with_guid,
)
from pyopenvba.access._macros import (
    NAV_MACRO_TYPE,
    OBJECT_MACRO,
    Macro,
    MacroAction,
    build_macro,
    parse_macro,
)
from pyopenvba.access._pages import (
    PAGE_DATA,
    PAGE_USAGE_BITMAP,
    PAGE_SIZE,
    ROW_DELETED,
    ROW_OVERFLOW,
    DatabaseHeader,
    PageStore,
    encode_row_pointer,
    page_owner,
    read_usage_map_ref,
    row_bytes,
    row_pointer,
    row_slots,
)
from pyopenvba.access._vba import (
    CRLF as VBA_CRLF,
    Reference,
    add_reference,
    make_libid,
    references,
    remove_reference,
    PROP_DATA_HAS_MODULE,
    TYPE_INFO_CLSID,
    add_to_project_documents,
    document_attributes,
    MODULETYPE,
    NAV_MODULE_GROUP,
    NAV_MODULE_TYPE,
    OBJECT_ID_STEP,
    PROP_DATA,
    STORAGE_TABLE,
    VBAModule,
    add_to_dir,
    add_to_dir_data,
    add_to_folder_list,
    add_to_project,
    add_to_project_wm,
    attribute_lines,
    dir_block,
    invalidate_cache,
    module_blocks,
    module_offset_at,
    module_stream,
    next_folder,
    read_source,
    remove_from_dir,
    remove_from_dir_data,
    remove_from_folder_list,
    remove_from_project,
    remove_from_project_wm,
    rename_attribute,
    rename_dir_data,
    rename_in_dir,
    rename_project,
    rename_project_wm,
    set_module_offset,
    split_source,
    stream_name_of,
    stream_row_name,
)
from pyopenvba.access._storage import (
    DIR_DATA,
    TYPE_FOLDER,
    TYPE_VALUE,
    dir_data_entries,
)
from pyopenvba.access._props import (
    BLOCK_COLUMN,
    BLOCK_OBJECT,
    DB_BOOLEAN,
    DB_INTEGER,
    DB_LONG,
    DB_TEXT,
    ENGINE_PROPERTIES,
    PropertyBlob,
    PropertyValue,
    dao_type_for,
    encode_property_value,
    parse_property_blob,
    serialize_property_blob,
)
from pyopenvba.access._queries import (
    ATTR_END,
    ATTR_START,
    ATTR_TYPE,
    QUERY_PASSTHROUGH,
    QueryRow,
    SavedQuery,
    rows_from_sql,
)
from pyopenvba.access._rows import (
    LongValueRef,
    RawRow,
    decode_long_value_ref,
    decode_scalar,
    decode_column_text,
    encode_row,
    encode_scalar,
    split_row,
)
from pyopenvba.access._schema import (
    DEFAULT_ACM,
    FIXED_SIZES,
    USAGE_MAP_ROW,
    ColumnSpec,
    DefinitionLayout,
    IndexSpec,
    build_definition,
    column_header,
    definition_page_count,
    definition_pages,
    empty_index_root,
    foreign_key_logical,
    mark_definition_freed,
    new_index_parts,
    serialize_definition,
    usage_map_page,
)
from pyopenvba.access._tdef import (
    OFFSET_LAST_COMPLEX_ID,
    SIZE_COLUMN_HEADER,
    FIXED_LENGTH_TYPES,
    INDEX_IGNORE_NULLS,
    INDEX_REQUIRED,
    INDEX_KIND_FOREIGN,
    OFFSET_INDEX_HEADERS,
    OFFSET_NEXT_AUTONUMBER,
    OFFSET_ROW_COUNT,
    SIZE_REAL_INDEX_HEADER,
    TYPE_BIGINT,
    TYPE_BINARY,
    TYPE_NUMERIC,
    TYPE_TEXT,
    TYPE_BOOLEAN,
    TYPE_COMPLEX,
    TYPE_DATETIME,
    TYPE_MEMO,
    TYPE_OLE,
    ColumnDef,
    LogicalIndex,
    RealIndex,
    TableDefinition,
    parse_column_header,
    parse_table_definition,
)
from pyopenvba.access._validate import Rules, apply_defaults
from pyopenvba.access._validate import check as check_rules
from pyopenvba.access._validate import read as read_rules
from pyopenvba.access_read import AccessError
from pyopenvba.vba import compress, decompress, parse_project_stream

MSYS_OBJECTS_PAGE = 2

# MSysObjects.Type values.
OBJECT_TABLE = 1
OBJECT_QUERY = 5
OBJECT_LINKED_TABLE = 6
OBJECT_RELATIONSHIP = 8
OBJECT_FORM = -32768
OBJECT_REPORT = -32764
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
# Every linked table carries this, whatever it is linked to; a source
# that is not another Access file carries the second bit as well
# (measured: a Jet link had 0x200000, a text link 0xA00000).
FLAG_LINKED = 0x00200000
FLAG_LINKED_FOREIGN = 0x00800000


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
    #: Where a linked table's rows really are; empty on a local object.
    connect: str = ""
    database: str = ""
    foreign_name: str = ""

    @property
    def is_system(self) -> bool:
        return bool(self.flags & FLAG_SYSTEM) or self.name.startswith("MSys")

    @property
    def is_table(self) -> bool:
        return self.type in (OBJECT_TABLE, OBJECT_LINKED_TABLE)


@dataclass(frozen=True)
class LinkedTable:
    """A table this database only points at: the catalog knows its name
    and where its rows are, and nothing else about it is stored here.
    Following the link means opening the other file, which is the
    caller's decision -- the path comes out of the database, so it is not
    one this library takes on its own."""

    name: str
    database: str
    source: str
    connect: str
    id: int
    flags: int
    date_create: object
    date_update: object

    @property
    def is_jet(self) -> bool:
        """True when the rows are in another Access file, which is the
        link with no connect string of its own."""
        return not self.connect


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
        self._rules: Rules | None = None

    # --- attachment and multi-valued columns --------------------------------

    def complex_columns(self) -> list[ComplexColumn]:
        """The attachment and multi-valued columns on this table."""
        return self._db.complex_columns(self.name)

    def add_complex_column(
        self, name: str, kind: str = "attachment", *, updated: object | None = None
    ) -> ComplexColumn:
        """Add an attachment or multi-valued column.

        `kind` is `"attachment"` or one of the scalar kinds Access offers
        (`"Text"`, `"Long"`, ...).  Everything a complex column costs goes
        in: the flat table that holds the values with its three indexes,
        the column on this table, the unique index over it, the
        `MSysComplexColumns` row that pairs them, and a complex id for
        every row already here.
        """
        if kind not in TYPE_TABLES:
            raise AccessError(f"kind must be one of {', '.join(sorted(TYPE_TABLES))}, not {kind!r}")
        if any(c.name.lower() == name.lower() for c in self.definition.columns):
            raise AccessError(f"table {self.name!r} already has a column named {name!r}")
        db = self._db
        type_table = next((e for e in db.catalog() if e.name == TYPE_TABLES[kind]), None)
        if type_table is None:
            raise AccessError(f"this database has no {TYPE_TABLES[kind]}, so it cannot hold one")

        rng = random.Random()
        guid = "".join(rng.choice("0123456789ABCDEF") for _ in range(32))
        key_column, id_column = "_" + name, f"{self.name}_{name}"

        # The flat table: the type's own columns, then the element id,
        # then the key, which is the order Access numbers them.
        if kind == "attachment":
            specs = [ColumnSpec(c, t, size=s) for c, t, s in ATTACHMENT_COLUMNS]
            scalar, misc = "FileName", {c: MISC_ATTACHMENT for c, _t, _s in ATTACHMENT_COLUMNS}
        else:
            value_type, value_size = SCALAR_TYPES[kind]
            specs = [ColumnSpec("Value", value_type, size=value_size)]
            scalar, misc = "Value", {"Value": MISC_SCALAR}
        # Access keeps both among the variable columns, not in the fixed
        # block, even though a Long would normally sit there.
        specs += [
            ColumnSpec(id_column, "Long", autonumber=True, variable=True),
            ColumnSpec(key_column, "Long", variable=True),
        ]
        misc[id_column], misc[key_column] = MISC_ELEMENT_ID, MISC_KEY

        flat = db.create_table(
            flat_table_name(guid, name),
            specs,
            [
                IndexSpec(key_column, (key_column,)),
                IndexSpec(SCALAR_INDEX, (key_column, scalar), unique=True),
                IndexSpec(PRIMARY_INDEX, (id_column,), primary=True),
            ],
            updated=updated,
        )
        for position, column in enumerate(flat.definition.columns):
            db.patch_column_header(
                flat.definition,
                position,
                patch_column_header(column.raw, misc_flags=misc[column.name]),
            )
        objects = db.table("MSysObjects")
        for rid, row in objects.rows_with_ids():
            if row["Id"] == flat.definition.page and row["Type"] == OBJECT_TABLE:
                objects.update_row(rid, {"Flags": FLAT_TABLE_FLAGS - (1 << 32)})
                break

        existing = [c.name for c in self.definition.columns if c.type_code == TYPE_COMPLEX]
        self.add_column(ColumnSpec(name, "Complex", autonumber=True), updated=updated)

        # Every row already here needs an id, and it needs one before the
        # unique index goes on: they would all be null, and nulls count as
        # equal.  The id belongs to the row, not the column, so a second
        # complex column takes the id the row already carries.
        counter = self.definition.last_complex_id
        for row_id, row in list(self.rows_with_ids()):
            shared = next((row[c] for c in existing if isinstance(row.get(c), int)), None)
            if shared is None:
                counter += 1
                shared = counter
            self.update_row(row_id, {name: shared})
        db.create_index(
            self.name, IndexSpec(index_name(name, guid), (name,), unique=True), updated=updated
        )

        # The catalog row of a table that has a complex column carries a
        # flag of its own, and no other table does.
        for rid, row in objects.rows_with_ids():
            if row["Id"] == self.definition.page and row["Type"] == OBJECT_TABLE:
                objects.update_row(rid, {"Flags": _as_int(row["Flags"] or 0) | HAS_COMPLEX_COLUMN})
                break

        pairing = db.table("MSysComplexColumns")
        rid = pairing.insert_row(
            {
                "ColumnName": name,
                "ComplexTypeObjectID": type_table.id,
                "ConceptualTableID": self.definition.page,
                "FlatTableID": flat.definition.page,
            }
        )
        complex_id = next(_as_int(r["ComplexID"]) for at, r in pairing.rows_with_ids() if at == rid)

        # The ComplexID lives in the column header's collation slot.
        self.definition = db.definition(self.definition.page)
        position = next(
            i for i, c in enumerate(self.definition.columns) if c.name.lower() == name.lower()
        )
        db.patch_column_header(
            self.definition,
            position,
            patch_column_header(self.definition.columns[position].raw, sort_order=complex_id),
        )
        # Last, because creating the index rebuilt the definition and a
        # rebuild writes the counter from the object it was handed.
        self.definition = db.definition(self.definition.page)
        self.definition.last_complex_id = counter
        db.patch_definition(self.definition, OFFSET_LAST_COMPLEX_ID, struct.pack("<I", counter))
        db.forget_catalog()
        return next(c for c in self.complex_columns() if c.column.lower() == name.lower())


    def _complex(self, column: str, attachment: bool | None = None) -> ComplexColumn:
        for found in self.complex_columns():
            if found.column.lower() != column.lower():
                continue
            if attachment is not None and found.is_attachment is not attachment:
                wanted = "attachments" if attachment else "scalar values"
                raise AccessError(
                    f"column {column!r} holds {found.kind} values, not {wanted}"
                )
            return found
        raise AccessError(f"table {self.name!r} has no complex column named {column!r}")

    def _elements(self, spec: ComplexColumn, key: int) -> list[dict[str, object]]:
        """A complex value's rows, ordered as the engine orders them.

        The flat table carries an index on `(<key column>, FileName)` --
        `Value` for a scalar column -- and that is the order DAO walks, so
        elements come back sorted by their name or value rather than by
        when they were added.  The element's own id breaks ties.
        """
        flat = self._db.table(spec.flat_table)
        scalar = "FileName" if spec.is_attachment else "Value"
        rows = [row for row in flat.rows() if _as_int(row[spec.key_column]) == key]
        return sorted(rows, key=lambda row: (str(row.get(scalar) or ""), _as_int(row[spec.id_column])))

    def _replace_elements(self, spec: ComplexColumn, key: int) -> Table:
        flat = self._db.table(spec.flat_table)
        for row_id, row in list(flat.rows_with_ids()):
            if _as_int(row[spec.key_column]) == key:
                flat.delete_row(row_id)
        return flat

    def attachments(self, column: str, key: int) -> list[Attachment]:
        """The files an attachment column holds for one row.

        `key` is the value the row itself carries in that column -- an id
        shared by every complex column in the row, not the row's own key.
        """
        out: list[Attachment] = []
        for row in self._elements(self._complex(column, attachment=True), key):
            blob = row.get("FileData")
            extension, data = decode_file_data(blob) if isinstance(blob, bytes) else ("", b"")
            flags = row.get("FileFlags")
            stamp = row.get("FileTimeStamp")
            out.append(
                Attachment(
                    name=str(row.get("FileName") or ""),
                    data=data,
                    type=str(row.get("FileType") or extension),
                    flags=flags if isinstance(flags, int) else None,
                    timestamp=stamp if isinstance(stamp, dt.datetime) else None,
                )
            )
        return out

    def multi_values(self, column: str, key: int) -> list[object]:
        """The values a multi-valued column holds for one row."""
        spec = self._complex(column, attachment=False)
        return [row.get("Value") for row in self._elements(spec, key)]

    def set_attachments(self, column: str, key: int, files: Sequence[Attachment]) -> None:
        """Replace the files an attachment column holds for one row."""
        spec = self._complex(column, attachment=True)
        flat = self._replace_elements(spec, key)
        for item in files:
            flat.insert_row(
                {
                    spec.key_column: key,
                    "FileData": encode_file_data(item.type, item.data),
                    "FileFlags": item.flags,
                    "FileName": item.name,
                    "FileTimeStamp": item.timestamp,
                    "FileType": item.type,
                    "FileURL": None,
                }
            )

    def set_multi_values(self, column: str, key: int, values: Sequence[object]) -> None:
        """Replace the values a multi-valued column holds for one row."""
        spec = self._complex(column, attachment=False)
        flat = self._replace_elements(spec, key)
        for value in values:
            flat.insert_row({spec.key_column: key, "Value": value})


    def rules(self) -> Rules:
        """What this table's properties say about the rows it will take:
        defaults, Required columns and validation rules (see
        :mod:`pyopenvba.access._validate`).  Read once and kept; a system
        table has none, which keeps every catalog write off the blob."""
        if self._rules is None:
            self._rules = (
                Rules()
                if self.definition.is_system or self.name.startswith("MSys")
                else read_rules(self)
            )
        return self._rules

    # -- structure -------------------------------------------------------------

    @property
    def columns(self) -> list[ColumnDef]:
        """The columns in the order the definition holds them, which is the
        order Access shows: column numbers only, until ALTER COLUMN gives a
        replacement column a new number in its predecessor's place."""
        return list(self.definition.columns)

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
                out[column.name] = (
                    decode_column_text(column, data) if column.type_code == TYPE_MEMO else data
                )
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
        """Add a row.  AutoNumber columns are assigned, a column with a
        DefaultValue takes it when the row does not name it, and every
        other column not given is null.  A row the table's own rules
        refuse (see :meth:`rules`) is not written.  Returns the row's home
        slot."""
        db = self._db
        d = self.definition
        rules = self.rules()
        values = apply_defaults(rules, values) if rules.defaults else dict(values)
        for column in d.columns:
            # A complex column is flagged AutoNumber too, but its id comes
            # from the complex counter below, not from this one.
            if column.type_code == TYPE_COMPLEX:
                continue
            if column.auto_number and values.get(column.name) is None:
                values[column.name] = d.next_autonumber + 1
                d.next_autonumber += 1
                db.patch_definition(d, OFFSET_NEXT_AUTONUMBER, struct.pack("<I", d.next_autonumber & 0xFFFFFFFF))
        # Every complex column in a row shares one id, handed out from the
        # counter at 0x1C.  A row with no elements still takes one, and an
        # id is not reused after a delete.
        complex_columns = [c for c in d.columns if c.type_code == TYPE_COMPLEX]
        if complex_columns and all(values.get(c.name) is None for c in complex_columns):
            d.last_complex_id += 1
            for column in complex_columns:
                values[column.name] = d.last_complex_id
            db.patch_definition(d, OFFSET_LAST_COMPLEX_ID, struct.pack("<I", d.last_complex_id))
        if rules:
            check_rules(self.name, rules, {c.name: values.get(c.name) for c in d.columns})
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
        store = self._db.store
        raw_page = store.read(row_id.page)
        entry = row_slots(raw_page)[row_id.slot]
        if not entry & ROW_OVERFLOW or entry & ROW_DELETED:
            return None
        pointer = row_bytes(raw_page, row_id.slot)
        if pointer is None:
            return None
        target_row, target_page = row_pointer(pointer)
        return target_page, target_row

    def delete_row(self, row_id: RowId, *, retire_empty: bool = True) -> None:
        """Delete one row with its index entries and long values, settling
        the pages it leaves as the engine does (see :meth:`_row_removed`).
        ``retire_empty=False`` leaves an emptied page alive and owned,
        which is what the engine does when a catalog row goes."""
        db = self._db
        d = self.definition
        data = self.fetch_row(row_id.page, row_id.slot)
        if data is None:
            raise AccessError(f"row ({row_id.page}, {row_id.slot}) is not live")
        parts = split_row(d, data)
        values = self._exact_values(parts)
        for i, real, columns in self._real_indexes():
            key = self._key(real, columns, values)
            if key is None:
                continue
            self._btree(i, real).delete(key, row_id.page, row_id.slot)
            self._row_left_index(i, real)
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
        self._row_removed(row_id.page, page, settle=moved is None, retire=retire_empty)
        d.row_count -= 1
        db.patch_definition(d, OFFSET_ROW_COUNT, struct.pack("<I", d.row_count))

    def _row_left_index(self, position: int, real: RealIndex) -> None:
        """One row stopped being counted by an index: a delete, or an update
        that wrote one of its columns.  An index built over existing rows
        counts them in its header; each such row takes one off, the count
        stops at zero, and the distinct-key count is capped at what is
        left.  An index made before its rows counts none and never
        changes (measured on a 30-row table and on two six-row ones)."""
        if not real.row_count:
            return
        real.row_count -= 1
        real.entry_count = min(real.entry_count, real.row_count)
        self._db.patch_definition(
            self.definition,
            OFFSET_INDEX_HEADERS + position * SIZE_REAL_INDEX_HEADER,
            struct.pack("<II", real.row_count, real.entry_count),
        )

    def _row_removed(self, page_number: int, page: DataPage, *, settle: bool = True, retire: bool = True) -> None:
        """Write back a page that just lost a row, the way the engine
        settles it: a page with rows left rejoins the free-space map; an
        emptied page is retired (type 0x09, released, out of both maps),
        unless it is the table's first data page, which stays.  Without
        ``settle`` (the row was only a pointer) the page is written back
        and nothing else moves; without ``retire`` an emptied page stays
        alive and owned, as it does when a catalog row goes."""
        db = self._db
        d = self.definition
        if not settle:
            db.store.write(page_number, page.to_bytes())
            return
        if retire and page.live_rows == 0 and page_number != min(read_usage_map_ref(db.store, d.owned_pages_ref).pages(), default=page_number):
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
            real.row_count = 0
            db.patch_definition(d, OFFSET_INDEX_HEADERS + i * SIZE_REAL_INDEX_HEADER, struct.pack("<II", 0, 0))
        d.row_count = 0
        db.patch_definition(d, OFFSET_ROW_COUNT, struct.pack("<I", d.row_count))

    # -- columns ---------------------------------------------------------------

    def add_column(self, spec: ColumnSpec, *, updated: object | None = None) -> ColumnDef:
        """Add a column as ``ALTER TABLE ... ADD COLUMN`` does: a header and
        name appended to the definition with the next column number, a
        fixed column placed just past the highest fixed column, a variable
        one given the next variable index; the rows are left as they are
        (they read back with the new column null), and the catalog row's
        DateUpdate is stamped."""
        import datetime as _dt

        d = self.definition
        if spec.autonumber:
            # A complex column is flagged AutoNumber but takes its id
            # from the complex counter, and Access does let you add one.
            if spec.type_code != TYPE_COMPLEX:
                raise AccessError("an AutoNumber column cannot be added to an existing table")
        if any(c.name.lower() == spec.name.lower() for c in d.columns):
            raise AccessError(f"table {self.name!r} already has a column named {spec.name!r}")
        code = spec.type_code
        is_fixed = code == TYPE_BOOLEAN or code in FIXED_SIZES or code == TYPE_BINARY
        number = d.max_columns
        if is_fixed:
            fixed_end = max((c.fixed_offset + c.length for c in d.columns if c.is_fixed and c.type_code != TYPE_BOOLEAN), default=0)
            # ALTER TABLE writes the current variable-column count into a
            # fixed column's variable-index field (CREATE TABLE writes 0).
            header = column_header(spec, number, d.var_column_count, 0 if code == TYPE_BOOLEAN else fixed_end, d.tag)
        else:
            header = column_header(spec, number, d.var_column_count, 0, d.tag)
        column = parse_column_header(header, spec.name)
        if column.is_long_value:
            # A Memo or OLE column brings two usage-map rows (owned and
            # free-space pages of its values) on the table's map page and a
            # map pair in the definition.
            owned_ref, free_ref = self._db._new_map_rows(d, 2)  # pyright: ignore[reportPrivateUsage]
            d.column_usage_maps[number] = (owned_ref, free_ref)
        d.columns.append(column)
        d.max_columns += 1
        if not is_fixed:
            d.var_column_count += 1
        # Header bytes 9..10 hold the column's position among the columns
        # present; an add recomputes them for every column, a drop leaves
        # them alone.
        for position, existing in enumerate(d.columns):
            raw = bytearray(existing.raw)
            struct.pack_into("<H", raw, 9, position)
            existing.raw = bytes(raw)
        self._db._write_definition(serialize_definition(d), d.page, d.pages[1:], keep_tail=True)  # pyright: ignore[reportPrivateUsage]
        self._stamp_catalog(updated if isinstance(updated, (_dt.datetime, float)) else _dt.datetime.now().replace(microsecond=0))
        self._db._definitions.pop(d.page, None)  # pyright: ignore[reportPrivateUsage]
        _write_column_properties(self._db.table(self.name), [spec], updated)
        self._rules = None
        return column

    def drop_column(self, name: str, *, updated: object | None = None) -> None:
        """Drop a column as ``ALTER TABLE ... DROP COLUMN`` does: its header
        and name leave the definition while the other columns keep their
        numbers, offsets and variable indexes and the maximum column count
        stays; rows are not rewritten; the catalog row's DateUpdate is
        stamped.  A Memo or OLE column also gives back its long-value
        pages and its two map rows.  A column an index uses, or the last
        column, is refused."""
        import datetime as _dt

        d = self.definition
        column = d.column(name)
        if len(d.columns) == 1:
            raise AccessError("a table keeps at least one column")
        for real in d.real_indexes:
            if any(c.number == column.number for c in real.columns):
                raise AccessError(f"column {column.name!r} is used by an index; drop the index first")
        if column.is_long_value and column.number in d.column_usage_maps:
            # The column's long-value pages go back untouched and its two
            # map rows are killed, as when a table is dropped.
            store = self._db.store
            owned_ref, free_ref = d.column_usage_maps.pop(column.number)
            released: list[int] = []
            for ref in (owned_ref, free_ref):
                umap = read_usage_map_ref(store, ref)
                for page_number in umap.pages():
                    if ref == owned_ref:
                        released.append(page_number)
                    set_usage_bit(store, umap, page_number, False)
            for ref in (owned_ref, free_ref):
                maps = DataPage(store.read(ref >> 8))
                maps.remove_row(ref & 0xFF)
                store.write(ref >> 8, maps.to_bytes())
            for page_number in released:
                if page_number < store.page_count:
                    release_page(store, page_number)
        d.columns.remove(column)
        self._db._write_definition(serialize_definition(d), d.page, d.pages[1:], keep_tail=True)  # pyright: ignore[reportPrivateUsage]
        self._stamp_catalog(updated if isinstance(updated, (_dt.datetime, float)) else _dt.datetime.now().replace(microsecond=0))
        self._db._definitions.pop(d.page, None)  # pyright: ignore[reportPrivateUsage]

    def alter_column(self, name: str, spec: ColumnSpec, *, updated: object | None = None) -> ColumnDef:
        """Retype or resize a column as ``ALTER TABLE ... ALTER COLUMN`` does:
        a new column with the new type takes the old one's name, place and
        position field under the next column number; every row is
        re-encoded with its value copied (or converted between numeric
        types) into the new column while the old column's bytes stay in the
        row as a phantom; the old header leaves the definition, which is
        written once; the catalog row is stamped.  Indexed, Memo and OLE
        columns are refused."""
        import datetime as _dt

        d = self.definition
        old = d.column(name)
        if spec.name.lower() != old.name.lower():
            raise AccessError("alter_column keeps the column's name; use rename_column to change it")
        if spec.autonumber or old.flags & 0x04:
            raise AccessError("an AutoNumber column cannot be altered")
        if old.is_long_value or spec.type_code in (TYPE_MEMO, TYPE_OLE):
            raise AccessError("altering to or from a Memo or OLE column is not written yet")
        for real in d.real_indexes:
            if any(c.number == old.number for c in real.columns):
                raise AccessError(f"column {old.name!r} is used by an index; drop the index first")
        code = spec.type_code
        is_fixed = code == TYPE_BOOLEAN or code in FIXED_SIZES or code == TYPE_BINARY
        number = d.max_columns
        fixed_end = max((c.fixed_offset + c.length for c in d.columns if c.is_fixed and c.type_code != TYPE_BOOLEAN), default=0)
        header = bytearray(column_header(spec, number, d.var_column_count, 0 if not is_fixed or code == TYPE_BOOLEAN else fixed_end, d.tag))
        header[9:11] = old.raw[9:11]  # the replacement keeps the old column's position
        new = parse_column_header(bytes(header), old.name)
        position = d.columns.index(old)
        d.columns.insert(position + 1, new)
        d.max_columns += 1
        var_count_before = d.var_column_count
        if not is_fixed:
            d.var_column_count += 1

        # Every row re-encoded under the definition holding both columns.
        convert = _converter(old.type_code, new.type_code)
        for page, slot, raw in list(self.raw_rows()):
            parts = split_row(d, raw)
            encoded: dict[int, bytes | None] = {c.number: parts.values.get(c.number) for c in d.columns if c.type_code != TYPE_BOOLEAN and c.number != new.number}
            booleans = {c.number for c in d.columns if c.type_code == TYPE_BOOLEAN and parts.present.get(c.number)}
            raw_value = parts.values.get(old.number)
            value: object = decode_scalar(old, raw_value) if isinstance(raw_value, bytes) else None
            if old.type_code == TYPE_BOOLEAN:
                value = bool(parts.present.get(old.number))
            if value is None:
                encoded[new.number] = None
            elif new.type_code == TYPE_BOOLEAN:
                if convert(value):
                    booleans.add(new.number)
            else:
                encoded[new.number] = encode_scalar(new, convert(value), compress_text=new.compressed_unicode)
            self._place_row(RowId(page, slot), encode_row(d, encoded, booleans, template=parts, template_var_count=var_count_before))

        d.columns.remove(old)
        self._db._write_definition(serialize_definition(d), d.page, d.pages[1:], keep_tail=True)  # pyright: ignore[reportPrivateUsage]
        self._stamp_catalog(updated if isinstance(updated, (_dt.datetime, float)) else _dt.datetime.now().replace(microsecond=0))
        self._db._definitions.pop(d.page, None)  # pyright: ignore[reportPrivateUsage]
        return new

    def rename_column(self, name: str, new_name: str, *, updated: object | None = None) -> None:
        """Rename a column as setting a Field's Name through DAO does: the
        name in the definition changes (the header does not), the column's
        property block in the catalog row's blob follows, every relationship
        row naming the column follows, and the catalog row's DateUpdate is
        stamped.  Indexes refer to columns by number and are untouched."""
        import datetime as _dt

        d = self.definition
        column = d.column(name)
        if not new_name or len(new_name) > 64:
            raise AccessError("a column name is 1 to 64 characters")
        if new_name.lower() != column.name.lower() and any(c.name.lower() == new_name.lower() for c in d.columns):
            raise AccessError(f"table {self.name!r} already has a column named {new_name!r}")
        when = updated if isinstance(updated, (_dt.datetime, float)) else _dt.datetime.now().replace(microsecond=0)
        old_name = column.name
        column.name = new_name
        self._db._write_definition(serialize_definition(d), d.page, d.pages[1:], keep_tail=True)  # pyright: ignore[reportPrivateUsage]
        relationships = self._db.table("MSysRelationships")
        for rid, row in list(relationships.rows_with_ids()):
            changes: dict[str, object] = {}
            if str(row["szObject"]).lower() == self.name.lower() and str(row["szColumn"]).lower() == old_name.lower():
                changes["szColumn"] = new_name
            if str(row["szReferencedObject"]).lower() == self.name.lower() and str(row["szReferencedColumn"]).lower() == old_name.lower():
                changes["szReferencedColumn"] = new_name
            if changes:
                relationships.update_row(rid, changes)
        blob = self.property_blob()
        catalog_changes: dict[str, object] = {"DateUpdate": when}
        if old_name in blob.column_properties:
            blob.column_properties = {new_name if key == old_name else key: value for key, value in blob.column_properties.items()}
            blob.block_order = [(kind, new_name if kind == BLOCK_COLUMN and target == old_name else target) for kind, target in blob.block_order]
            catalog_changes["LvProp"] = serialize_property_blob(blob)
        rid, _row = self._catalog_row()
        self._db.table("MSysObjects").update_row(rid, catalog_changes)
        self._db._catalog = None  # pyright: ignore[reportPrivateUsage]
        self._db._definitions.pop(d.page, None)  # pyright: ignore[reportPrivateUsage]

    def _stamp_catalog(self, when: object) -> None:
        rid, _row = self._catalog_row()
        self._db.table("MSysObjects").update_row(rid, {"DateUpdate": when})
        self._db._catalog = None  # pyright: ignore[reportPrivateUsage]

    # -- properties ------------------------------------------------------------

    def property_blob(self) -> PropertyBlob:
        """The table's ``MR2`` property blob from its catalog row, parsed;
        empty when the row carries none."""
        _rid, row = self._catalog_row()
        lv = row.get("LvProp")
        return parse_property_blob(lv) if isinstance(lv, bytes) and lv else PropertyBlob()

    def properties(self) -> dict[str, object]:
        """The table's own properties (Description, Orientation, ...) decoded."""
        return self.property_blob().decoded()

    def column_properties(self, column: str) -> dict[str, object]:
        """One column's properties (Caption, Format, DecimalPlaces, ...) decoded."""
        name = self.definition.column(column).name
        return self.property_blob().decoded_column(name)

    def set_properties(
        self,
        values: Mapping[str, object],
        *,
        column: str | None = None,
        updated: object | None = None,
    ) -> None:
        """Add or replace properties on the table, or on ``column``, the way
        DAO's ``Properties.Append`` does: the blob is rebuilt with the new
        records appended in the order given (an existing property keeps its
        type, flags and place), stored as the catalog row's LvProp, and the
        row's stamps are left alone.  A value may be a
        :class:`~pyopenvba.access._props.PropertyValue` to control the DAO
        type and flags; otherwise the type follows the Python value, except
        for the five the engine keeps for itself (Required,
        AllowZeroLength, DefaultValue, ValidationRule, ValidationText),
        which take the engine's own type and flags.  A column's
        ValidationRule ends in the NUL DAO's setter leaves there.

        Writing one of those five stamps the catalog row's DateUpdate with
        ``updated`` (now by default), because it goes through the engine's
        own field definition; appending a client property (Caption,
        Description) leaves the stamps alone.  Both measured."""
        import datetime as _dt

        blob = self.property_blob()
        if column is None:
            records = blob.object_properties
            if (BLOCK_OBJECT, "") not in blob.block_order:
                blob.block_order.append((BLOCK_OBJECT, ""))
        else:
            name = self.definition.column(column).name
            records = blob.column_properties.setdefault(name, {})
            if (BLOCK_COLUMN, name) not in blob.block_order:
                blob.block_order.append((BLOCK_COLUMN, name))
        for prop, value in values.items():
            if isinstance(value, PropertyValue):
                records[prop] = value
                continue
            existing = records.get(prop)
            engine = ENGINE_PROPERTIES.get(prop)
            if existing is not None:
                dao_type, flags = existing.type, existing.flags
            elif engine is not None:
                dao_type, flags = engine
            else:
                dao_type, flags = dao_type_for(value), 0
            if column is not None and prop == "ValidationRule" and isinstance(value, str) and not value.endswith(chr(0)):
                value = value + chr(0)
            records[prop] = PropertyValue(dao_type, flags, encode_property_value(dao_type, value))
        rid, _row = self._catalog_row()
        changes: dict[str, object] = {"LvProp": serialize_property_blob(blob)}
        if any(name in ENGINE_PROPERTIES for name in values):
            changes["DateUpdate"] = (
                updated if isinstance(updated, (_dt.datetime, float)) else _dt.datetime.now().replace(microsecond=0)
            )
        self._db.table("MSysObjects").update_row(rid, changes)
        self._db._catalog = None  # pyright: ignore[reportPrivateUsage]
        self._rules = None

    def _catalog_row(self) -> tuple[RowId, dict[str, object]]:
        objects = self._db.table("MSysObjects")
        for rid, row in objects.rows_with_ids():
            if row["Id"] == self.definition.page and row["Type"] in (OBJECT_TABLE, OBJECT_LINKED_TABLE):
                return rid, row
        raise AccessError(f"table {self.name!r} has no catalog row")

    def update_row(self, row_id: RowId, changes: Mapping[str, object]) -> None:
        """Change the given columns of one row; the rest keep their values.
        A row that no longer fits its page moves to another page behind an
        overflow pointer, and moves back when it fits again, as the engine
        does."""
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
        rules = self.rules()
        if rules:
            check_rules(self.name, rules, new_values, columns=set(given))
        self._check_unique(new_values, exclude=row_id)
        encoded, booleans = self._encode_values(given, keep_raw=parts)
        self._free_long_values(parts, only=set(given))
        row = encode_row(d, encoded, booleans, template=parts)
        for i, real, columns in self._real_indexes():
            old_key = self._key(real, columns, old_values)
            new_key = self._key(real, columns, new_values)
            if old_key is not None and any(c.name in given for c, _ in columns):
                # A row written through an index costs the index one of the
                # rows it counts, even when the value does not change
                # (measured: SET M = M dropped the counter all the same).
                self._row_left_index(i, real)
            if old_key == new_key:
                continue
            tree = self._btree(i, real)
            if old_key is not None:
                tree.delete(old_key, row_id.page, row_id.slot)
            if new_key is not None:
                # The distinct-key count grows on inserts only: an update
                # that moves a row to a new key leaves it alone (measured on
                # a table rename through the catalog's name index).
                tree.insert(new_key, row_id.page, row_id.slot)
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
        self._vba_changed = False
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

    # -- transactions ------------------------------------------------------------

    @contextmanager
    def transaction(self) -> Generator[AccessDatabase]:
        """Group writes so they all land or none do.  Leaving the block
        normally keeps them; an exception puts the database back exactly
        as it was, pages and session state alike.  The engine writes the
        same bytes either way -- a DAO transaction changes nothing about
        where its pages land (measured) -- so this is a way of undoing
        work, not a different way of writing it."""
        state = self.store.snapshot()
        try:
            yield self
        except BaseException:
            # The pages go back, and everything read from them is thrown
            # away: a definition object was edited in place on the way in.
            self.store.restore(state)
            self._catalog = None
            self._definitions.clear()
            raise

    # -- SQL ---------------------------------------------------------------------

    def execute(
        self,
        sql: str,
        parameters: Mapping[str, object] | None = None,
        *,
        created: object | None = None,
        updated: object | None = None,
        referenced_updated: object | None = None,
        owner_updated: object | None = None,
    ) -> list[dict[str, object]] | int:
        """Run one SQL statement.  SELECT returns its rows as dicts keyed by
        the output column names; INSERT, UPDATE and DELETE return the
        number of rows affected; CREATE, DROP and ALTER return 0, as DAO
        does.  ``parameters`` supplies ``[Name]`` references that are not
        columns, and ``created``/``updated`` the catalog timestamps a DDL
        statement stamps, with ``referenced_updated`` for the other table
        of a FOREIGN KEY and ``owner_updated`` for the middle write of a
        CREATE TABLE whose columns carry properties.  See
        :mod:`pyopenvba.access._sql` and :mod:`pyopenvba.access._ddl` for
        the grammar."""
        from pyopenvba.access._sql import execute

        return execute(
            self,
            sql,
            parameters,
            created=created,
            updated=updated,
            referenced_updated=referenced_updated,
            owner_updated=owner_updated,
        )

    # -- persistence -------------------------------------------------------------

    def table_specs(self, name: str) -> tuple[list[ColumnSpec], list[IndexSpec]]:
        """The table described as the specs that would create it again.

        Sizes come back in the units `ColumnSpec` takes -- characters for
        Text, bytes for Binary -- not the bytes the header holds.
        """
        table = self.table(name)
        definition = table.definition
        properties = {c.name: table.column_properties(c.name) for c in definition.columns}
        columns: list[ColumnSpec] = []
        for column in definition.columns_by_number():
            own = properties.get(column.name, {})
            size: int | tuple[int, int] | None = None
            if column.type_code == TYPE_TEXT:
                size = column.length // 2
            elif column.type_code == TYPE_BINARY:
                size = column.length
            elif column.type_code == TYPE_NUMERIC:
                size = (column.precision, column.scale)
            zero = own.get("AllowZeroLength")
            columns.append(
                ColumnSpec(
                    column.name,
                    column.type_name,
                    size=size,
                    autonumber=column.auto_number,
                    compressed=column.compressed_unicode,
                    variable=not column.is_fixed
                    and column.type_code in FIXED_LENGTH_TYPES,
                    required=bool(own.get("Required")),
                    default=_as_text(own.get("DefaultValue")),
                    allow_zero_length=bool(zero) if zero is not None else None,
                    validation_rule=_as_text(own.get("ValidationRule")),
                    validation_text=_as_text(own.get("ValidationText")),
                )
            )
        indexes: list[IndexSpec] = []
        for logical in definition.logical_indexes:
            if logical.kind == INDEX_KIND_FOREIGN:
                continue
            real = definition.real_indexes[logical.real_index]
            indexes.append(
                IndexSpec(
                    logical.name,
                    tuple(
                        (definition.column_by_number(c.number).name, c.ascending)
                        for c in real.columns
                    ),
                    unique=real.unique,
                    primary=logical.is_primary_key,
                    ignore_nulls=bool(real.flags & INDEX_IGNORE_NULLS),
                    required=bool(real.flags & INDEX_REQUIRED),
                )
            )
        return columns, indexes

    def rebuild_table(self, name: str) -> None:
        """Write a table's rows out again so they land on as few pages as
        they need.

        Deleting rows does not shrink a table -- not here and not in
        Access, which is what Compact and Repair is for.  The rows are
        read out, the table is dropped and made again from the same
        definition, and the rows are written back; the pages the old copy
        held come free, and :meth:`compact` can then give them back.

        Both halves are the writers the engine was measured against, so
        this adds no new way to lay a table out.  It refuses a table it
        cannot carry across whole -- one with a complex column, a link, or
        a relationship naming it -- and it compares the rows before and
        after, putting the whole database back if they differ.  A rebuild
        either round-trips or leaves nothing changed.
        """
        table = self.table(name)
        definition = table.definition
        if any(c.type_code == TYPE_COMPLEX for c in definition.columns):
            raise AccessError(
                f"{name!r} has a complex column, whose values live in another "
                f"table keyed by an id this cannot carry across"
            )
        if any(link.name.lower() == name.lower() for link in self.links()):
            raise AccessError(f"{name!r} is a link, so it holds no rows to pack")
        naming = [
            r.name
            for r in self.relationships()
            if name.lower() in (r.table.lower(), r.referenced_table.lower())
        ]
        if naming:
            raise AccessError(
                f"{name!r} is named by the relationship {naming[0]!r}, which a "
                f"rebuild would leave pointing at a table that no longer exists"
            )

        before = list(table.rows())
        columns, indexes = self.table_specs(name)
        table_properties = dict(table.properties())
        counter = definition.next_autonumber
        state = self.store.snapshot()
        try:
            self.drop_table(name)
            # The engine keeps a dropped table's pages back until the
            # database is reopened, so the rows would otherwise land past
            # the end of the file and the rebuild would gain nothing.
            self.store.reopen()
            rebuilt = self.create_table(name, columns, indexes)
            if table_properties:
                rebuilt.set_properties(table_properties)
            for row in before:
                rebuilt.insert_row(row)
            if counter:
                fresh = self.table(name).definition
                fresh.next_autonumber = counter
                self.patch_definition(
                    fresh, OFFSET_NEXT_AUTONUMBER, struct.pack("<I", counter & 0xFFFFFFFF)
                )
            after = list(self.table(name).rows())
        except Exception:
            self.store.restore(state)
            self.forget_catalog()
            raise
        if after != before:
            self.store.restore(state)
            self.forget_catalog()
            raise AccessError(
                f"rebuilding {name!r} did not round-trip its rows, so nothing "
                f"was changed"
            )


    def compact(self, rebuild: bool = False) -> int:
        """Give back the free pages at the end of the file, and say how
        many went.

        By itself this reclaims the run of free pages the file ends with,
        which is what a dropped table leaves behind, and touches nothing
        else: no page moves and no object is rewritten, so nothing can be
        lost.  A file whose free pages are only in the middle keeps its
        size, and deleting rows leaves exactly that -- a table does not
        give its pages back as it empties.

        `rebuild=True` writes each table's rows out again first, so they
        land on as few pages as they need and the rest of the file's free
        pages join the run at the end; see :meth:`rebuild_table`, whose
        refusals apply, and which leaves a table it will not touch alone.
        A 3000-row table cut to 300 came back from 229 pages to 93.

        This is still not Access's Compact and Repair: that renumbers
        every page, and resets each AutoNumber to one past its largest
        value where a rebuild here keeps the counter as it was.
        """
        if rebuild:
            for name in self.table_names():
                try:
                    self.rebuild_table(name)
                except AccessError:
                    # A table a rebuild will not carry across is left as
                    # it is; the rest of the file still compacts.
                    continue
        store = self.store
        free = read_usage_map(store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW)
        spare = set(free.pages())
        keep = store.page_count
        while keep > 2:
            page = keep - 1
            # A usage-bitmap page is what records the free ones, so it
            # stays even when nothing points at it any more.
            if page not in spare or store.page_type(page) == PAGE_USAGE_BITMAP:
                break
            keep -= 1
        if keep >= store.page_count:
            return 0
        for page in range(keep, store.page_count):
            set_usage_bit(store, free, page, False)
        return store.truncate(keep)


    def to_bytes(self) -> bytes:
        return self.store.to_bytes()

    def save(self, path: str | Path | None = None, *, allow_protected: bool = False) -> Path:
        """Write the database out.

        A VBA change made to a **password-protected** project is refused
        here rather than written, which is what the other hosts do on
        their own `save`.  Pass `allow_protected=True` to write it anyway:
        the protection bytes are kept as they are, so the result still
        wants the original password, and what was written may or may not
        be what the project's owner intended.
        """
        if self._vba_changed and not allow_protected and self.vba_is_protected():
            raise AccessError(
                "this database's VBA project is password-protected; pass "
                "allow_protected=True to save a change to it anyway"
            )
        target = Path(path) if path is not None else self.path
        if target is None:
            raise AccessError("no path to save to; the database was opened from bytes")
        target.write_bytes(self.to_bytes())
        return target

    def vba_is_protected(self) -> bool:
        """Whether the VBA project carries a password.

        Read from the `PROJECT` stream's `DPB` record, the same way the
        other hosts read it.
        """
        try:
            _modules, project_id, _streams = self._vba_storage_ids()
        except AccessError:
            return False
        for _rid, row in self.table(STORAGE_TABLE).rows_with_ids():
            payload = row.get("Lv")
            if (
                str(row["Name"]) == "PROJECT"
                and _as_int(row["ParentId"]) == project_id
                and isinstance(payload, bytes)
            ):
                return parse_project_stream(payload).protection.has_password
        return False

    def _write_definition(self, stream: bytes, first: int, old_chain: Sequence[int], *, keep_tail: bool = False) -> list[int]:
        """Lay a definition stream over ``first`` and a continuation chain.
        As the engine does on every rewrite, the continuation pages are
        allocated afresh in ascending order, chained in reverse, and only
        then are the pages of ``old_chain`` released (their bytes kept).
        Returns the chain written."""
        store = self.store
        fresh = [allocate_page(store) for _ in range(definition_page_count(len(stream)) - 1)]
        chain = fresh[::-1]
        images = definition_pages(stream, chain)
        if keep_tail:
            # A rewritten definition is written up to its new length plus the
            # eight reserved bytes, zeroed; when it shrank, the dropped
            # entries' bytes beyond that stay on the first page (measured on
            # DROP CONSTRAINT).  A new table's page is filled fresh.
            kept_from = min(len(stream) + 8, PAGE_SIZE)
            images[0] = images[0][:kept_from] + store.read(first)[kept_from:]
        for page, image in zip([first, *chain], images, strict=True):
            store.write(page, image)
        for page in old_chain:
            release_page(store, page, kind="rewrite")
        return chain

    def patch_definition(self, definition: TableDefinition, offset: int, data: bytes) -> None:
        """Overwrite bytes of a table definition's header, which always
        lies on its first page.

        This does not drop the parsed definition the database has cached
        -- inserting a row patches the AutoNumber counter for every row
        written, and re-parsing each time would cost more than it saves.
        A caller has to keep the object it holds in step, as the row
        writer does.
        """
        if offset + len(data) > 0x3F + len(definition.real_indexes) * 12:
            raise AccessError("definition patches are limited to the fixed header and index headers")
        raw = bytearray(self.store.read(definition.page))
        raw[offset : offset + len(data)] = data
        self.store.write(definition.page, bytes(raw))

    # -- catalog -----------------------------------------------------------------

    def forget_catalog(self) -> None:
        """Drop the cached catalog, so the next read picks up a change made
        through a table's own writers."""
        self._catalog = None

    def patch_column_header(self, definition: TableDefinition, position: int, raw: bytes) -> None:
        """Overwrite one column's 25-byte header.

        `patch_definition` stops at the index headers; the column headers
        follow them, in the order the definition lists the columns.  Only
        a complex column needs this: three of its header fields have no
        `ColumnSpec` to carry them.
        """
        if len(raw) != SIZE_COLUMN_HEADER:
            raise AccessError(f"a column header is {SIZE_COLUMN_HEADER} bytes, not {len(raw)}")
        start = (
            OFFSET_INDEX_HEADERS
            + len(definition.real_indexes) * SIZE_REAL_INDEX_HEADER
            + position * SIZE_COLUMN_HEADER
        )
        if start + SIZE_COLUMN_HEADER > PAGE_SIZE:
            raise AccessError("this table's column headers do not fit its first definition page")
        page = bytearray(self.store.read(definition.page))
        page[start : start + SIZE_COLUMN_HEADER] = raw
        self.store.write(definition.page, bytes(page))
        self._definitions.pop(definition.page, None)


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
                        connect=str(values.get("Connect") or ""),
                        database=str(values.get("Database") or ""),
                        foreign_name=str(values.get("ForeignName") or ""),
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
        owner_updated: object | None = None,
    ) -> Table:
        """Create a table the way the engine does: a definition page, a page
        of usage maps, an empty root per index, and the catalog rows.
        ``created`` and ``updated`` are the catalog row's two timestamps
        (now, and the creation time, by default), as datetimes or as the
        stored serials; the engine stamps DateUpdate when the definition
        is complete, so on a large table it runs a little after
        DateCreate.  ``owner_updated`` is the DateUpdate written with the
        owner: the last one when no column asks for a property, an
        intermediate one when a column does.  It defaults to
        ``updated``."""
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
        with_owner = owner_updated if isinstance(owner_updated, (_dt.datetime, float)) else when_updated
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
        objects.update_row(catalog_row, {"Owner": self._default_owner(), "DateUpdate": with_owner})
        if any(spec.type_code == TYPE_BIGINT for spec in specs):
            # A BigInt column is newer than the file format, so the engine
            # records the versions needed to read, write and design the
            # table (measured: only BigInt does this, of the types that can
            # be created).  They arrive one at a time, each rewriting the
            # whole blob, as every property append does.
            for count in range(1, len(VERSION_PROPERTIES) + 1):
                objects.update_row(catalog_row, {"LvProp": serialize_property_blob(_version_properties(count))})
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
        table = self.table(name)
        _write_column_properties(table, columns, updated)
        return table

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
        umap_ref = self._new_map_rows(d, 1)[0]
        root = allocate_page(store)
        store.write(root, empty_index_root(d.page))
        real, logical = new_index_parts(spec, d, len(d.real_indexes), umap_ref, root)
        d.real_indexes.append(real)
        d.logical_indexes.append(logical)
        d.real_index_count += 1
        d.logical_index_count += 1
        self._write_definition(serialize_definition(d), d.page, d.pages[1:], keep_tail=True)
        add_to_map(store, umap_ref, root)
        self._definitions.pop(d.page, None)
        table = self.table(table_name)
        # Existing rows get their entries, in home-slot order.
        position = len(table.definition.real_indexes) - 1
        real = table.definition.real_indexes[position]
        columns = [(table.definition.column_by_number(c.number), c.ascending) for c in real.columns]
        tree = table._btree(position, real)  # pyright: ignore[reportPrivateUsage]
        distinct = 0
        entries = 0
        # The engine builds a new index in key order, so its leaves fill and
        # split at the end rather than in the middle (measured: 111/111/111/67
        # entries over 400 text keys, where row order gives 92/92/92/124).
        keyed: list[tuple[bytes, RowId, dict[str, object]]] = []
        for row_id, values in table.rows_with_ids():
            key = table._key(real, columns, values)  # pyright: ignore[reportPrivateUsage]
            if key is not None:
                keyed.append((key, row_id, dict(values)))
        keyed.sort(key=lambda item: (item[0], item[1].page, item[1].slot))
        for key, row_id, values in keyed:
            if real.unique:
                table._check_unique(values, exclude=row_id)  # pyright: ignore[reportPrivateUsage]
            entries += 1
            if tree.insert(key, row_id.page, row_id.slot):
                distinct += 1
        if entries:
            # An index built over rows records how many it holds and how
            # many distinct keys; deletes later take the first down.
            real.row_count = entries
            real.entry_count = distinct
            self.patch_definition(
                table.definition, OFFSET_INDEX_HEADERS + position * SIZE_REAL_INDEX_HEADER, struct.pack("<II", entries, distinct)
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

    def complex_columns(self, table: str | None = None) -> list[ComplexColumn]:
        """Every attachment and multi-valued column, with the flat table
        that holds its values.

        A `Complex` column keeps nothing in the row but a Long; the values
        live one per row in `f_<GUID>_<Column>`, joined on that Long.
        `MSysComplexColumns` names the pairing.
        """
        names = {entry.id: entry.name for entry in self.catalog()}
        out: list[ComplexColumn] = []
        for row in self.table("MSysComplexColumns").rows():
            parent = names.get(_as_int(row["ConceptualTableID"]))
            flat = names.get(_as_int(row["FlatTableID"]))
            if parent is None or flat is None:
                continue
            if table is not None and parent.lower() != table.lower():
                continue
            type_name = names.get(_as_int(row["ComplexTypeObjectID"]), "")
            column = str(row["ColumnName"])
            out.append(
                ComplexColumn(
                    table=parent,
                    column=column,
                    flat_table=flat,
                    kind="attachment"
                    if type_name == ATTACHMENT_TYPE
                    else type_name.removeprefix("MSysComplexType_"),
                    complex_id=_as_int(row["ComplexID"]),
                    key_column="_" + column,
                    id_column=f"{parent}_{column}",
                )
            )
        return out


    # --- macros -------------------------------------------------------------
    # A macro is a binary blob under `Scripts`, laid out the same way a
    # module is under `Modules`; see `pyopenvba.access._macros`.

    def _scripts_id(self) -> int:
        rows = [row for _rid, row in self.table(STORAGE_TABLE).rows_with_ids()]
        root = next(
            (_as_int(r["Id"]) for r in rows if str(r["Name"]) == "MSysAccessStorage_ROOT"), None
        )
        if root is None:
            raise AccessError("this database has no object storage")
        for row in rows:
            if (
                _as_int(row["ParentId"]) == root
                and str(row["Name"]) == "Scripts"
                and _as_int(row["Type"]) == TYPE_FOLDER
            ):
                return _as_int(row["Id"])
        raise AccessError("MSysAccessStorage has no 'Scripts' folder")

    def _macro_blob(self, folder: str) -> tuple[RowId, bytes] | None:
        scripts = self._scripts_id()
        storage = self.table(STORAGE_TABLE)
        folders = {
            _as_int(r["Id"]): str(r["Name"])
            for _rid, r in storage.rows_with_ids()
            if _as_int(r["ParentId"]) == scripts and _as_int(r["Type"]) == TYPE_FOLDER
        }
        wanted = next((i for i, name in folders.items() if name == folder), None)
        if wanted is None:
            return None
        for rid, row in storage.rows_with_ids():
            payload = row.get("Lv")
            if (
                _as_int(row["ParentId"]) == wanted
                and str(row["Name"]) == "Blob"
                and isinstance(payload, bytes)
            ):
                return rid, payload
        return None

    # --- the project's references -------------------------------------------

    def references(self) -> list[Reference]:
        """Every library the VBA project points at.

        The two a project always has -- VBA itself and Access -- are not
        in the file, so they are not here either.
        """
        return references(self._vba_dir()[1])

    def add_reference(
        self,
        name: str,
        guid: str,
        major: int = 1,
        minor: int = 0,
        *,
        path: str = "",
        description: str = "",
        lcid: int = 0,
    ) -> Reference:
        """Point the project at a type library.

        `guid` is the library's, with or without its braces, and the
        version is written in hex the way Access writes it -- DAO 12.0 as
        `c.0`.
        """
        if any(found.name.lower() == name.lower() for found in self.references()):
            raise AccessError(f"the project already references {name!r}")
        libid = make_libid(guid, major, minor, path, description, lcid)
        rid, dir_stream = self._vba_dir()
        self.table(STORAGE_TABLE).update_row(
            rid, {"Lv": compress(add_reference(dir_stream, name, libid))}
        )
        self._invalidate_vba_cache()
        self._drop_srp()
        return next(found for found in self.references() if found.name == name)

    def drop_reference(self, name: str) -> None:
        """Stop pointing at a library."""
        rid, dir_stream = self._vba_dir()
        self.table(STORAGE_TABLE).update_row(
            rid, {"Lv": compress(remove_reference(dir_stream, name))}
        )
        self._invalidate_vba_cache()
        self._drop_srp()


    def macros(self) -> list[Macro]:
        """Every macro in the database, in the order `Scripts` lists them."""
        scripts = self._scripts_id()
        listing = next(
            (
                row.get("Lv")
                for _rid, row in self.table(STORAGE_TABLE).rows_with_ids()
                if _as_int(row["ParentId"]) == scripts and str(row["Name"]) == DIR_DATA
            ),
            None,
        )
        if not isinstance(listing, bytes):
            return []
        out: list[Macro] = []
        for name, folder in dir_data_entries(listing):
            found = self._macro_blob(folder)
            out.append(Macro(name, parse_macro(found[1]) if found else ()))
        return out

    def macro(self, name: str) -> Macro:
        for found in self.macros():
            if found.name.lower() == name.lower():
                return found
        raise AccessError(f"this database has no macro named {name!r}")

    def create_macro(
        self,
        name: str,
        actions: Sequence[MacroAction],
        *,
        updated: object | None = None,
    ) -> Macro:
        """Add a macro that runs `actions`."""
        if not name or len(name) > 64:
            raise AccessError("a macro name is 1 to 64 characters")
        if any(found.name.lower() == name.lower() for found in self.macros()):
            raise AccessError(f"a macro named {name!r} already exists")
        blob = build_macro(tuple(actions))

        scripts = self._scripts_id()
        storage = self.table(STORAGE_TABLE)
        rows = [row for _rid, row in storage.rows_with_ids()]
        folders = {
            str(r["Name"])
            for r in rows
            if _as_int(r["ParentId"]) == scripts and _as_int(r["Type"]) == TYPE_FOLDER
        }
        folder = next_folder("Scripts", folders)
        when = (
            updated
            if isinstance(updated, (dt.datetime, float))
            else dt.datetime.now().replace(microsecond=0)
        )

        folder_rid = storage.insert_row(
            {"ParentId": scripts, "Name": folder, "Type": TYPE_FOLDER,
             "DateCreate": when, "DateUpdate": when}
        )
        folder_id = next(_as_int(r["Id"]) for rid, r in storage.rows_with_ids() if rid == folder_rid)
        storage.insert_row(
            {"ParentId": folder_id, "Name": "Blob", "Type": TYPE_VALUE, "Lv": blob,
             "DateCreate": when, "DateUpdate": when}
        )
        listing = next(
            (
                (rid, row.get("Lv"))
                for rid, row in storage.rows_with_ids()
                if _as_int(row["ParentId"]) == scripts and str(row["Name"]) == DIR_DATA
            ),
            None,
        )
        if listing is None:
            # A database's first macro brings the listing with it.
            storage.insert_row(
                {"ParentId": scripts, "Name": DIR_DATA, "Type": TYPE_VALUE,
                 "Lv": add_to_dir_data(bytes(4), name, folder),
                 "DateCreate": when, "DateUpdate": when}
            )
        else:
            rid, payload = listing
            storage.update_row(
                rid,
                {"Lv": add_to_dir_data(payload if isinstance(payload, bytes) else bytes(4), name, folder)},
            )

        objects = self.table("MSysObjects")
        container = next(e.id for e in self.catalog() if e.name == "Scripts" and e.type == 3)
        owner = next((e.owner for e in self.catalog() if e.owner), None)
        # A macro's object id steps by one; a module's steps by four.
        object_id = max((e.id for e in self.catalog() if e.id < 0), default=-(2**31)) + 1
        objects.insert_row(
            {"Id": object_id, "ParentId": container, "Name": name, "Type": OBJECT_MACRO,
             "Flags": 0, "Owner": owner, "LvProp": _macro_properties(),
             "DateCreate": when, "DateUpdate": when}
        )
        self.table("MSysNavPaneObjectIDs").insert_row(
            {"Id": object_id, "Name": name, "Type": NAV_MACRO_TYPE}
        )
        self._catalog = None
        return self.macro(name)

    def delete_macro(self, name: str) -> None:
        """Remove a macro and every structure it occupies."""
        found = self.macro(name)
        scripts = self._scripts_id()
        storage = self.table(STORAGE_TABLE)
        listing_rid, listing_payload = next(
            (rid, payload)
            for rid, row in storage.rows_with_ids()
            if _as_int(row["ParentId"]) == scripts
            and str(row["Name"]) == DIR_DATA
            and isinstance(payload := row.get("Lv"), bytes)
        )
        folder = dict(dir_data_entries(listing_payload))[found.name]
        folders = {
            _as_int(r["Id"]): str(r["Name"])
            for _rid, r in storage.rows_with_ids()
            if _as_int(r["ParentId"]) == scripts and _as_int(r["Type"]) == TYPE_FOLDER
        }
        folder_id = next(i for i, folder_name in folders.items() if folder_name == folder)
        doomed = {folder_id}
        for rid, row in list(storage.rows_with_ids()):
            if _as_int(row["Id"]) in doomed or _as_int(row["ParentId"]) in doomed:
                storage.delete_row(rid, retire_empty=False)
        storage.update_row(
            listing_rid, {"Lv": remove_from_dir_data(listing_payload, found.name)}
        )

        objects = self.table("MSysObjects")
        for rid, row in list(objects.rows_with_ids()):
            if row["Type"] == OBJECT_MACRO and str(row["Name"]) == found.name:
                object_id = _as_int(row["Id"])
                objects.delete_row(rid, retire_empty=False)
                nav = self.table("MSysNavPaneObjectIDs")
                for nav_rid, nav_row in list(nav.rows_with_ids()):
                    if _as_int(nav_row["Id"]) == object_id:
                        nav.delete_row(nav_rid, retire_empty=False)
                break
        self._catalog = None


    # --- forms and reports --------------------------------------------------
    # Both live under `MSysAccessStorage` the way a module does, with the
    # design in a `Blob`; see `pyopenvba.access._designs`.

    def _design_container(self, kind: str) -> int:
        if kind not in CONTAINERS:
            raise AccessError(f"kind must be 'form' or 'report', not {kind!r}")
        rows = [row for _rid, row in self.table(STORAGE_TABLE).rows_with_ids()]
        root = next(
            (_as_int(r["Id"]) for r in rows if str(r["Name"]) == "MSysAccessStorage_ROOT"), None
        )
        if root is None:
            raise AccessError("this database has no object storage")
        for row in rows:
            if (
                _as_int(row["ParentId"]) == root
                and str(row["Name"]) == CONTAINERS[kind]
                and _as_int(row["Type"]) == TYPE_FOLDER
            ):
                return _as_int(row["Id"])
        raise AccessError(f"MSysAccessStorage has no {CONTAINERS[kind]!r} folder")

    def _design_blob(self, container: int, folder: str) -> bytes | None:
        storage = self.table(STORAGE_TABLE)
        wanted = next(
            (
                _as_int(r["Id"])
                for _rid, r in storage.rows_with_ids()
                if _as_int(r["ParentId"]) == container
                and _as_int(r["Type"]) == TYPE_FOLDER
                and str(r["Name"]) == folder
            ),
            None,
        )
        if wanted is None:
            return None
        for _rid, row in storage.rows_with_ids():
            payload = row.get("Lv")
            if (
                _as_int(row["ParentId"]) == wanted
                and str(row["Name"]) == "Blob"
                and isinstance(payload, bytes)
            ):
                return payload
        return None

    def _designs(self, kind: str) -> list[AccessDesign]:
        container = self._design_container(kind)
        listing = next(
            (
                row.get("Lv")
                for _rid, row in self.table(STORAGE_TABLE).rows_with_ids()
                if _as_int(row["ParentId"]) == container and str(row["Name"]) == DIR_DATA
            ),
            None,
        )
        if not isinstance(listing, bytes):
            return []
        out: list[AccessDesign] = []
        for name, folder in dir_data_entries(listing):
            blob = self._design_blob(container, folder)
            objects = parse_design(blob)[1] if blob else ()
            out.append(AccessDesign(name, kind, objects))
        return out

    def forms(self) -> list[AccessDesign]:
        """Every form, with its sections and controls."""
        return self._designs("form")

    def reports(self) -> list[AccessDesign]:
        """Every report, with its sections and controls."""
        return self._designs("report")

    def form(self, name: str) -> AccessDesign:
        return self._design("form", name)

    def report(self, name: str) -> AccessDesign:
        return self._design("report", name)

    def _design(self, kind: str, name: str) -> AccessDesign:
        for found in self._designs(kind):
            if found.name.lower() == name.lower():
                return found
        raise AccessError(f"this database has no {kind} named {name!r}")

    def create_form(self, name: str, *, updated: object | None = None) -> AccessDesign:
        """Add an empty form."""
        return self._create_design("form", name, updated=updated)

    def create_report(self, name: str, *, updated: object | None = None) -> AccessDesign:
        """Add an empty report, with its page header, detail and page
        footer sections."""
        return self._create_design("report", name, updated=updated)

    def _create_design(self, kind: str, name: str, *, updated: object | None) -> AccessDesign:
        """The design itself comes from a captured template -- an empty one
        as Access writes it -- with a GUID of its own patched in, since the
        catalog row repeats it and two objects sharing one is not something
        Access writes."""
        if not name or len(name) > 64:
            raise AccessError(f"a {kind} name is 1 to 64 characters")
        if any(found.name.lower() == name.lower() for found in self._designs(kind)):
            raise AccessError(f"a {kind} named {name!r} already exists")

        container = self._design_container(kind)
        storage = self.table(STORAGE_TABLE)
        rows = [row for _rid, row in storage.rows_with_ids()]
        folders = {
            str(r["Name"])
            for r in rows
            if _as_int(r["ParentId"]) == container and _as_int(r["Type"]) == TYPE_FOLDER
        }
        folder = next_folder(CONTAINERS[kind], folders)
        when = (
            updated
            if isinstance(updated, (dt.datetime, float))
            else dt.datetime.now().replace(microsecond=0)
        )
        guid = random.Random().randbytes(16)

        folder_rid = storage.insert_row(
            {"ParentId": container, "Name": folder, "Type": TYPE_FOLDER,
             "DateCreate": when, "DateUpdate": when}
        )
        folder_id = next(_as_int(r["Id"]) for rid, r in storage.rows_with_ids() if rid == folder_rid)
        for stream, payload in (
            ("Blob", with_guid(template(kind, "blob"), guid)),
            ("TypeInfo", template(kind, "typeinfo")),
            ("BlobDelta", None),
            ("PropData", template(kind, "propdata")),
        ):
            values: dict[str, object] = {
                "ParentId": folder_id, "Name": stream, "Type": TYPE_VALUE,
                "DateCreate": when, "DateUpdate": when,
            }
            if payload is not None:
                values["Lv"] = payload
            storage.insert_row(values)

        adders: tuple[tuple[str, Callable[[bytes], bytes]], ...] = (
            (DIR_DATA, lambda payload: add_to_dir_data(payload, name, folder)),
            ("PropData", lambda payload: add_to_folder_list(payload, folder)),
        )
        for stream, add in adders:
            found = next(
                (
                    (rid, row.get("Lv"))
                    for rid, row in storage.rows_with_ids()
                    if _as_int(row["ParentId"]) == container and str(row["Name"]) == stream
                ),
                None,
            )
            if found is None:
                storage.insert_row(
                    {"ParentId": container, "Name": stream, "Type": TYPE_VALUE,
                     "Lv": add(bytes(4)), "DateCreate": when, "DateUpdate": when}
                )
            else:
                rid, payload = found
                storage.update_row(
                    rid, {"Lv": add(payload if isinstance(payload, bytes) else bytes(4))}
                )

        objects = self.table("MSysObjects")
        parent = next(e.id for e in self.catalog() if e.name == CATALOG_CONTAINERS[kind] and e.type == 3)
        owner = next((e.owner for e in self.catalog() if e.owner), None)
        object_id = max((e.id for e in self.catalog() if e.id < 0), default=-(2**31)) + 1
        objects.insert_row(
            {"Id": object_id, "ParentId": parent, "Name": name, "Type": OBJECT_TYPES[kind],
             "Flags": 0, "Owner": owner, "LvProp": _design_properties(kind, guid),
             "DateCreate": when, "DateUpdate": when}
        )
        self.table("MSysNavPaneObjectIDs").insert_row(
            {"Id": object_id, "Name": name, "Type": NAV_TYPES[kind]}
        )
        self.forget_catalog()
        return self._design(kind, name)

    def add_control(
        self,
        design: str,
        control_type: str,
        name: str,
        *,
        kind: str = "form",
        section: str = "Detail",
        parent: str | None = None,
        left: int = 0,
        top: int = 0,
        width: int = 1440,
        height: int = 240,
        caption: str | None = None,
    ) -> AccessDesign:
        """Put a control on a form or report.

        `control_type` is one of the twenty-one whose slots were read
        back from a control Access itself made.  A record's id is its slot
        in the object's own schema and the schema differs by control type,
        so a type whose slots have not been measured is refused rather
        than guessed at, and so is one that carries records this project
        cannot name -- a chart, a navigation control, an Edge browser.

        `caption` is the label's or button's text and the control source
        of anything that binds to data, and `section` says which band the
        control belongs to.  `parent` names a control that holds controls
        of its own: a tab control holds pages, and a page is the one type
        that must have a parent.  Sizes are in twips, as Access keeps them; a
        page break takes only its top, which is all Access writes for one.
        """
        return self._rewrite_design(
            kind,
            design,
            lambda blob: add_control(
                blob,
                control_type,
                name,
                random.Random().randbytes(16),
                section=section,
                parent=parent,
                left=left,
                top=top,
                width=width,
                height=height,
                caption=caption,
            ),
        )

    def _rewrite_design(
        self, kind: str, design: str, change: Callable[[bytes], bytes]
    ) -> AccessDesign:
        """Put `change` through the design's blob, which lives in one row
        of the storage table under the design's own folder."""
        found = self._design(kind, design)
        container = self._design_container(kind)
        listing = next(
            payload
            for _rid, row in self.table(STORAGE_TABLE).rows_with_ids()
            if _as_int(row["ParentId"]) == container
            and str(row["Name"]) == DIR_DATA
            and isinstance(payload := row.get("Lv"), bytes)
        )
        folder = dict(dir_data_entries(listing))[found.name]
        folder_id = next(
            _as_int(r["Id"])
            for _rid, r in self.table(STORAGE_TABLE).rows_with_ids()
            if _as_int(r["ParentId"]) == container
            and _as_int(r["Type"]) == TYPE_FOLDER
            and str(r["Name"]) == folder
        )
        storage = self.table(STORAGE_TABLE)
        for rid, row in list(storage.rows_with_ids()):
            payload = row.get("Lv")
            if (
                _as_int(row["ParentId"]) == folder_id
                and str(row["Name"]) == "Blob"
                and isinstance(payload, bytes)
            ):
                storage.update_row(rid, {"Lv": change(payload)})
                break
        else:  # pragma: no cover - a design without its own blob
            raise AccessError(f"the {kind} {design!r} has no design blob")
        return self._design(kind, design)

    def set_control_property(
        self,
        design: str,
        control: str,
        name: str,
        value: object,
        *,
        kind: str = "form",
    ) -> AccessDesign:
        """Change one property of one control, or of one section.

        `name` is the property's own name -- `Caption`, `FontSize`,
        `BackColor` -- and what it may hold follows from how the design
        stores it: text for a caption, a number for a colour or a size, a
        `bool` for a flag.  A property the control already carries keeps
        its record where it is; one it does not is written at the id its
        control type's schema gives it.

        Which properties a type has is `PROPERTY_SLOTS`, measured off
        controls Access itself wrote; a name that is not among them is
        refused rather than written somewhere it does not belong.
        """
        return self._rewrite_design(
            kind, design, lambda blob: set_property(blob, control, name, value)
        )

    def set_design_property(
        self, design: str, name: str, value: object, *, kind: str = "form"
    ) -> AccessDesign:
        """Change one property of the form or report itself."""
        return self._rewrite_design(
            kind, design, lambda blob: set_property(blob, None, name, value)
        )

    #: What a form's or report's own module is called.
    DESIGN_MODULE_PREFIX = {"form": "Form_", "report": "Report_"}

    def set_design_code(self, design: str, code: str, *, kind: str = "form") -> VBAModule:
        """Put code behind a form or report.

        The module is named after the design -- `Form_Summary` for a form,
        `Report_Monthly` for a report -- and is created if the design has
        none.  It belongs to the design rather than to `Modules`, so it
        gets no storage folder and no catalog row of its own; what makes
        the design answer to it is a `DocClass=` line in `PROJECT` and a
        CLSID shared between the design's `TypeInfo` and the module's
        `VB_Base`.
        """
        found = self._design(kind, design)
        name = self.DESIGN_MODULE_PREFIX[kind] + found.name
        if any(module.name.lower() == name.lower() for module in self.modules()):
            self.set_module_source(name, code)
            return self.module(name)

        import uuid as _uuid

        clsid = _uuid.uuid4()
        rng = random.Random()
        storage = self.table(STORAGE_TABLE)
        _modules, project_id, streams_id = self._vba_storage_ids()
        dir_rid, dir_stream = self._vba_dir()
        row_name = stream_row_name(
            rng, {str(r["Name"]) for _rid, r in storage.rows_with_ids()}
        )
        when = dt.datetime.now().replace(microsecond=0)
        storage.insert_row(
            {
                "ParentId": streams_id,
                "Name": row_name,
                "Type": TYPE_VALUE,
                "Lv": module_stream(document_attributes(name, str(clsid).upper()), code),
                "DateCreate": when,
                "DateUpdate": when,
            }
        )
        storage.update_row(
            dir_rid,
            {
                "Lv": compress(
                    add_to_dir(dir_stream, dir_block(name, row_name, rng.randbytes(2), "class"))
                )
            },
        )
        for rid, row in list(storage.rows_with_ids()):
            payload = row.get("Lv")
            if not isinstance(payload, bytes) or not payload:
                continue
            label, parent = str(row["Name"]), _as_int(row["ParentId"])
            if label == "_VBA_PROJECT":
                storage.update_row(rid, {"Lv": invalidate_cache(payload)})
            elif label == "PROJECTwm" and parent == project_id:
                storage.update_row(rid, {"Lv": add_to_project_wm(payload, name)})
            elif label == "PROJECT" and parent == project_id:
                storage.update_row(
                    rid,
                    {
                        "Lv": add_to_project_documents(
                            payload.decode("latin-1"), name
                        ).encode("latin-1")
                    },
                )

        # the design's own two: the CLSID it shares with the module, and
        # the byte that says it has one
        container = self._design_container(kind)
        listing = next(
            payload
            for _rid, row in storage.rows_with_ids()
            if _as_int(row["ParentId"]) == container
            and str(row["Name"]) == DIR_DATA
            and isinstance(payload := row.get("Lv"), bytes)
        )
        folder = dict(dir_data_entries(listing))[found.name]
        folder_id = next(
            _as_int(r["Id"])
            for _rid, r in storage.rows_with_ids()
            if _as_int(r["ParentId"]) == container
            and _as_int(r["Type"]) == TYPE_FOLDER
            and str(r["Name"]) == folder
        )
        for rid, row in list(storage.rows_with_ids()):
            payload = row.get("Lv")
            if _as_int(row["ParentId"]) != folder_id or not isinstance(payload, bytes):
                continue
            if str(row["Name"]) == "TypeInfo":
                raw = bytearray(payload)
                raw[TYPE_INFO_CLSID : TYPE_INFO_CLSID + 16] = clsid.bytes_le
                storage.update_row(rid, {"Lv": bytes(raw)})
            elif str(row["Name"]) == "PropData":
                raw = bytearray(payload)
                raw[PROP_DATA_HAS_MODULE] = 1
                storage.update_row(rid, {"Lv": bytes(raw)})
        self._drop_srp()
        return self.module(name)


    def delete_form(self, name: str) -> None:
        """Remove a form and every structure it occupies."""
        self._delete_design("form", name)

    def delete_report(self, name: str) -> None:
        """Remove a report and every structure it occupies."""
        self._delete_design("report", name)

    def _delete_design(self, kind: str, name: str) -> None:
        found = self._design(kind, name)
        container = self._design_container(kind)
        storage = self.table(STORAGE_TABLE)
        listing_rid, listing_payload = next(
            (rid, payload)
            for rid, row in storage.rows_with_ids()
            if _as_int(row["ParentId"]) == container
            and str(row["Name"]) == DIR_DATA
            and isinstance(payload := row.get("Lv"), bytes)
        )
        folder = dict(dir_data_entries(listing_payload))[found.name]
        folder_id = next(
            _as_int(r["Id"])
            for _rid, r in storage.rows_with_ids()
            if _as_int(r["ParentId"]) == container
            and _as_int(r["Type"]) == TYPE_FOLDER
            and str(r["Name"]) == folder
        )
        for rid, row in list(storage.rows_with_ids()):
            if _as_int(row["Id"]) == folder_id or _as_int(row["ParentId"]) == folder_id:
                storage.delete_row(rid, retire_empty=False)
        storage.update_row(
            listing_rid, {"Lv": remove_from_dir_data(listing_payload, found.name)}
        )
        for rid, row in list(storage.rows_with_ids()):
            payload = row.get("Lv")
            if (
                _as_int(row["ParentId"]) == container
                and str(row["Name"]) == "PropData"
                and isinstance(payload, bytes)
            ):
                storage.update_row(rid, {"Lv": remove_from_folder_list(payload, folder)})

        objects = self.table("MSysObjects")
        for rid, row in list(objects.rows_with_ids()):
            if row["Type"] == OBJECT_TYPES[kind] and str(row["Name"]) == found.name:
                object_id = _as_int(row["Id"])
                objects.delete_row(rid, retire_empty=False)
                nav = self.table("MSysNavPaneObjectIDs")
                for nav_rid, nav_row in list(nav.rows_with_ids()):
                    if _as_int(nav_row["Id"]) == object_id:
                        nav.delete_row(nav_rid, retire_empty=False)
                break
        self.forget_catalog()


    def database_properties(self) -> dict[str, object]:
        """The database's own settings, from the MSysDb row's property blob."""
        for row in self.table("MSysObjects").rows():
            if row["Name"] == "MSysDb":
                lv = row.get("LvProp")
                return parse_property_blob(lv).decoded() if isinstance(lv, bytes) and lv else {}
        return {}

    def queries(self) -> list[SavedQuery]:
        """Every saved query: its catalog name and MSysQueries rows."""
        names = {e.id: e.name for e in self.catalog() if e.type == OBJECT_QUERY}
        by_id: dict[int, list[QueryRow]] = {oid: [] for oid in names}
        for row in self.table("MSysQueries").rows():
            oid = _as_int(row["ObjectId"])
            if oid not in by_id:
                continue
            order = row["Order"]
            by_id[oid].append(
                QueryRow(
                    attribute=_as_int(row["Attribute"]),
                    order=int.from_bytes(order, "big") if isinstance(order, bytes) else _as_int(order or 0),
                    name1=None if row["Name1"] is None else str(row["Name1"]),
                    name2=None if row["Name2"] is None else str(row["Name2"]),
                    expression=None if row["Expression"] is None else str(row["Expression"]),
                    flag=None if row["Flag"] is None else _as_int(row["Flag"]),
                )
            )
        return [SavedQuery(names[oid], rows) for oid, rows in by_id.items()]

    def query(self, name: str) -> SavedQuery:
        for saved in self.queries():
            if saved.name.lower() == name.lower():
                return saved
        raise AccessError(f"no query named {name!r}")

    def create_query(
        self,
        name: str,
        sql: str,
        *,
        created: object | None = None,
        updated: object | None = None,
        owner_updated: object | None = None,
        connect: str | None = None,
    ) -> SavedQuery:
        """Save a query as DAO's ``CreateQueryDef`` does: the MSysQueries
        rows for ``sql`` (see :mod:`pyopenvba.access._queries` for the
        subset covered), a catalog object of type 5 under the Tables
        container carrying the two properties DAO gives every query, three
        permission rows, and DAO's query type in Flags (0 for a select).
        DAO stamps the row three times: ``created`` on the insert,
        ``owner_updated`` when it sets the owner (that version's bytes
        outlive the final one on the page) and ``updated`` with the last
        write; each defaults to the previous.

        ``connect`` makes it a pass-through query instead: ``sql`` is then
        whatever the server should run, kept as written, and the rows are
        the two a query always has plus one type row carrying the connect
        string.  DAO gets there by making an empty query, setting Connect
        and then SQL, which leaves a dead row behind; this does the
        same."""
        import datetime as _dt

        if any(e.name.lower() == name.lower() for e in self.catalog()):
            raise AccessError(f"an object named {name!r} already exists")
        rows = _passthrough_rows(connect, sql) if connect is not None else rows_from_sql(sql)
        now = _dt.datetime.now().replace(microsecond=0)
        when = created if isinstance(created, (_dt.datetime, float)) else now
        when_owner = owner_updated if isinstance(owner_updated, (_dt.datetime, float)) else when
        when_updated = updated if isinstance(updated, (_dt.datetime, float)) else when_owner
        object_id = self._next_object_id()
        saved = SavedQuery(name, rows)
        objects = self.table("MSysObjects")
        catalog_row = objects.insert_row(
            {
                "Id": object_id,
                "ParentId": self._container("Tables").id,
                "Name": name,
                "Type": OBJECT_QUERY,
                "Flags": 0,
                "DateCreate": when,
                "DateUpdate": when,
            }
        )
        objects.update_row(catalog_row, {"Owner": self._default_owner(), "DateUpdate": when_owner})
        # DAO appends the two properties one at a time: the first blob is
        # short enough to sit inline in the row, the second moves it to a
        # long-value page, and the row's earlier, longer version leaves
        # its bytes above the final one -- carrying the owner step's stamp
        # and Flags 0, so the final DateUpdate and an action query's type
        # arrive with the last write.
        blob = PropertyBlob(block_order=[(BLOCK_OBJECT, "")])
        blob.object_properties["ODBCTimeout"] = PropertyValue(DB_INTEGER, 1, struct.pack("<h", 60))
        objects.update_row(catalog_row, {"LvProp": serialize_property_blob(blob)})
        blob.object_properties["MaxRecords"] = PropertyValue(DB_LONG, 1, struct.pack("<i", 0))
        final: dict[str, object] = {"LvProp": serialize_property_blob(blob), "DateUpdate": when_updated}
        if saved.catalog_flags:
            final["Flags"] = saved.catalog_flags
        objects.update_row(catalog_row, final)
        aces = self.table("MSysACEs")
        for ace in self._default_aces():
            aces.insert_row(dict(ace, ObjectId=object_id))
        queries = self.table("MSysQueries")

        def write(row: QueryRow) -> RowId:
            return queries.insert_row(
                {
                    "ObjectId": object_id,
                    "Attribute": row.attribute,
                    "Order": row.order.to_bytes(4, "big"),
                    "Name1": row.name1,
                    "Name2": row.name2,
                    "Expression": row.expression,
                    "Flag": row.flag,
                }
            )

        for row in rows:
            if connect is not None and row.attribute == ATTR_TYPE:
                # Setting Connect writes this row with no SQL in it; setting
                # SQL then deletes that row and writes another.
                queries.delete_row(write(QueryRow(row.attribute, row.order, name1=row.name1, flag=row.flag)))
            write(row)
        self._catalog = None
        return saved

    # -- linked tables -------------------------------------------------------------

    def links(self) -> list[LinkedTable]:
        """Every table this database only points at, in catalog order."""
        return [
            LinkedTable(
                name=e.name,
                database=e.database,
                source=e.foreign_name,
                connect=e.connect,
                id=e.id,
                flags=e.flags,
                date_create=e.date_create,
                date_update=e.date_update,
            )
            for e in self.catalog()
            if e.type == OBJECT_LINKED_TABLE
        ]

    def link(self, name: str) -> LinkedTable:
        for linked in self.links():
            if linked.name.lower() == name.lower():
                return linked
        raise AccessError(f"no linked table named {name!r}")

    def link_table(
        self,
        name: str,
        database: str,
        source: str,
        *,
        connect: str = "",
        created: object | None = None,
        updated: object | None = None,
    ) -> LinkedTable:
        """Point at a table in another file the way ``TableDefs.Append``
        does with a Connect string: one catalog row of type 6 under the
        Tables container, carrying the file in Database, the table's name
        over there in ForeignName and the linked flag, plus the three
        permission rows every object gets.  There is no definition page
        and no data: nothing of the table itself is stored here.

        A link to another Access file leaves ``connect`` empty, which is
        what DAO stores for ``;DATABASE=<path>``; another source keeps its
        prefix there (``Text;``, ``ODBC;...``) and its row carries the
        second linked flag.  The catalog id is the
        next one up from the lowest in use, as the engine allocates the
        ids of objects that have no definition page."""
        import datetime as _dt

        if any(e.name.lower() == name.lower() for e in self.catalog()):
            raise AccessError(f"an object named {name!r} already exists")
        when = created if isinstance(created, (_dt.datetime, float)) else _dt.datetime.now().replace(microsecond=0)
        when_updated = updated if isinstance(updated, (_dt.datetime, float)) else when
        object_id = max((e.id for e in self.catalog() if e.id < 0), default=-(2**31)) + 1
        objects = self.table("MSysObjects")
        row = objects.insert_row(
            {
                "Id": object_id,
                "ParentId": self._tables_container().id,
                "Name": name,
                "Type": OBJECT_LINKED_TABLE,
                "Flags": FLAG_LINKED | (FLAG_LINKED_FOREIGN if connect else 0),
                "DateCreate": when,
                "DateUpdate": when,
            }
        )
        # Where the rows really are arrives with the owner, on the second
        # write, so the row takes its page before the long values take
        # theirs (measured: the engine's row and its Database value went
        # on to two freed pages in that order).  An empty connect string
        # is stored as null, which is what a link to another Access file
        # has.
        objects.update_row(
            row,
            {
                "Owner": self._default_owner(),
                "DateUpdate": when_updated,
                "Connect": connect or None,
                "Database": database,
                "ForeignName": source,
            },
        )
        aces = self.table("MSysACEs")
        for ace in self._default_aces():
            aces.insert_row(dict(ace, ObjectId=object_id))
        self._catalog = None
        return self.link(name)

    def drop_link(self, name: str) -> None:
        """Forget a linked table: its catalog row and its three permission
        rows.  The file it pointed at is not touched."""
        linked = self.link(name)
        objects = self.table("MSysObjects")
        for rid, row in list(objects.rows_with_ids()):
            if row["Id"] == linked.id and row["Type"] == OBJECT_LINKED_TABLE:
                # The page the row was alone on stays alive and owned
                # where a filtered DELETE would have retired it (measured
                # against TableDefs.Delete).
                objects.delete_row(rid, retire_empty=False)
                break
        aces = self.table("MSysACEs")
        for rid, row in list(aces.rows_with_ids()):
            if row["ObjectId"] == linked.id:
                aces.delete_row(rid)
        self._catalog = None

    def drop_query(self, name: str) -> None:
        """Remove a saved query as ``QueryDefs.Delete`` does: its MSysQueries
        rows, its catalog row (freeing the property blob) and its three
        permission rows."""
        saved = self.query(name)
        entry = next(e for e in self.catalog() if e.type == OBJECT_QUERY and e.name == saved.name)
        queries = self.table("MSysQueries")
        # The rows go in the order of the (ObjectId, Attribute, Order) index
        # the engine finds them through, which decides the bytes the page
        # keeps after compaction.
        by_index = queries.index("ObjectIdAttribute")
        doomed = [RowId(page, row) for key, page, row in by_index.entries() if key[0] == entry.id]
        for rid in doomed:
            queries.delete_row(rid)
        objects = self.table("MSysObjects")
        for rid, row in list(objects.rows_with_ids()):
            if row["Id"] == entry.id and row["Type"] == OBJECT_QUERY:
                objects.delete_row(rid)
                break
        aces = self.table("MSysACEs")
        for rid, row in list(aces.rows_with_ids()):
            if row["ObjectId"] == entry.id:
                aces.delete_row(rid)
        self._catalog = None

    def rename_table(self, name: str, new_name: str, *, updated: object | None = None) -> None:
        """Rename a table as setting a TableDef's Name does: the catalog
        row's Name and DateUpdate change, and every MSysRelationships row
        naming the table as its object or referenced object follows.  The
        definition is not touched (it does not carry the name)."""
        import datetime as _dt

        table = self.table(name)
        if new_name.lower() != name.lower() and any(e.name.lower() == new_name.lower() for e in self.catalog()):
            raise AccessError(f"an object named {new_name!r} already exists")
        if not new_name or len(new_name) > 64:
            raise AccessError("a table name is 1 to 64 characters")
        when = updated if isinstance(updated, (_dt.datetime, float)) else _dt.datetime.now().replace(microsecond=0)
        relationships = self.table("MSysRelationships")
        for rid, row in list(relationships.rows_with_ids()):
            changes: dict[str, object] = {}
            if str(row["szObject"]).lower() == name.lower():
                changes["szObject"] = new_name
            if str(row["szReferencedObject"]).lower() == name.lower():
                changes["szReferencedObject"] = new_name
            if changes:
                relationships.update_row(rid, changes)
        objects = self.table("MSysObjects")
        for rid, row in objects.rows_with_ids():
            if row["Id"] == table.definition.page and row["Type"] in (OBJECT_TABLE, OBJECT_LINKED_TABLE):
                objects.update_row(rid, {"Name": new_name, "DateUpdate": when})
                break
        self._catalog = None

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
        umap_ref = self._new_map_rows(cd, 1)[0]
        root = allocate_page(store)
        store.write(root, empty_index_root(cd.page))
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
        self._write_definition(serialize_definition(cd), cd.page, cd.pages[1:], keep_tail=True)
        add_to_map(store, umap_ref, root)
        self._definitions.pop(cd.page, None)
        child = self.table(table)
        position = len(child.definition.real_indexes) - 1
        real = child.definition.real_indexes[position]
        key_columns = [(child.definition.column_by_number(c.number), c.ascending) for c in real.columns]
        tree = child._btree(position, real)  # pyright: ignore[reportPrivateUsage]
        distinct = entries = 0
        for row_id, values in child.rows_with_ids():
            key = child._key(real, key_columns, values)  # pyright: ignore[reportPrivateUsage]
            if key is None:
                continue
            entries += 1
            if tree.insert(key, row_id.page, row_id.slot):
                distinct += 1
        if entries:
            real.row_count = entries
            real.entry_count = distinct
            self.patch_definition(child.definition, OFFSET_INDEX_HEADERS + position * SIZE_REAL_INDEX_HEADER, struct.pack("<II", entries, distinct))

        # The referenced side: a ``.r<letter>`` entry sharing the unique index.
        pd.logical_indexes.append(
            foreign_key_logical(
                pd.tag, ".r" + chr(0x41 + parent_logical_number), parent_logical_number, parent_real,
                referencing=False, other_logical=child_logical_number, other_page=cd.page,
                cascade_updates=cascade_updates, cascade_deletes=cascade_deletes,
            )
        )
        pd.logical_index_count += 1
        self._write_definition(serialize_definition(pd), pd.page, pd.pages[1:], keep_tail=True)
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

    def drop_relationship(
        self,
        name: str,
        *,
        table_updated: object | None = None,
        referenced_updated: object | None = None,
    ) -> None:
        """Remove a relationship as ``ALTER TABLE ... DROP CONSTRAINT`` does:
        the foreign-key index's map bits are cleared, its map row killed and
        its pages released untouched, the logical entries on both sides go
        (the others keep their numbers), the MSysRelationships rows, the
        catalog object and its permission rows are deleted, and both
        tables' DateUpdate is stamped."""
        import datetime as _dt

        rel = next((r for r in self.relationships() if r.name.lower() == name.lower()), None)
        if rel is None:
            raise AccessError(f"no relationship named {name!r}")
        child = self.table(rel.table)
        parent = self.table(rel.referenced_table)
        cd, pd = child.definition, parent.definition
        fk_position = next((i for i, li in enumerate(cd.logical_indexes) if li.kind == INDEX_KIND_FOREIGN and li.name.lower() == rel.name.lower()), None)
        if fk_position is None:
            raise AccessError(f"table {rel.table!r} has no foreign-key index named {rel.name!r}")
        fk = cd.logical_indexes[fk_position]
        back_position = next(
            (
                i for i, li in enumerate(pd.logical_indexes)
                if li.kind == INDEX_KIND_FOREIGN and li.relationship_table_page == cd.page and li.relationship_index == fk.number
            ),
            None,
        )
        if back_position is None:
            raise AccessError(f"table {rel.referenced_table!r} carries no entry for relationship {rel.name!r}")
        now = _dt.datetime.now().replace(microsecond=0)
        child_when = table_updated if isinstance(table_updated, (_dt.datetime, float)) else now
        parent_when = referenced_updated if isinstance(referenced_updated, (_dt.datetime, float)) else child_when
        store = self.store

        # The foreign-key index goes: map bits cleared, map row killed,
        # pages released untouched; the definition loses the real index and
        # its logical entry, later real indexes shifting down by one.
        real = cd.real_indexes[fk.real_index]
        umap = read_usage_map_ref(store, real.usage_map_ref)
        released = list(umap.pages())
        for page_number in released:
            set_usage_bit(store, umap, page_number, False)
        map_page = DataPage(store.read(real.usage_map_ref >> 8))
        map_page.remove_row(real.usage_map_ref & 0xFF)
        store.write(real.usage_map_ref >> 8, map_page.to_bytes())
        del cd.real_indexes[fk.real_index]
        del cd.logical_indexes[fk_position]
        for li in cd.logical_indexes:
            if li.real_index > fk.real_index:
                li.real_index -= 1
                struct.pack_into("<I", raw := bytearray(li.raw), 8, li.real_index)
                li.raw = bytes(raw)
        cd.real_index_count -= 1
        cd.logical_index_count -= 1
        self._write_definition(serialize_definition(cd), cd.page, cd.pages[1:], keep_tail=True)
        for page_number in released:
            if page_number < store.page_count:
                release_page(store, page_number)
        self._definitions.pop(cd.page, None)

        del pd.logical_indexes[back_position]
        pd.logical_index_count -= 1
        self._write_definition(serialize_definition(pd), pd.page, pd.pages[1:], keep_tail=True)
        self._definitions.pop(pd.page, None)

        relationships = self.table("MSysRelationships")
        for rid, row in list(relationships.rows_with_ids()):
            if str(row["szRelationship"]).lower() == rel.name.lower():
                relationships.delete_row(rid)
        objects = self.table("MSysObjects")
        object_id: int | None = None
        for rid, row in list(objects.rows_with_ids()):
            if row["Type"] == OBJECT_RELATIONSHIP and str(row["Name"]).lower() == rel.name.lower():
                object_id = _as_int(row["Id"])
                objects.delete_row(rid)
                break
        if object_id is not None:
            aces = self.table("MSysACEs")
            for rid, row in list(aces.rows_with_ids()):
                if row["ObjectId"] == object_id:
                    aces.delete_row(rid)
        for definition_page, stamp in ((cd.page, child_when), (pd.page, parent_when)):
            for rid, row in objects.rows_with_ids():
                if row["Id"] == definition_page and row["Type"] == OBJECT_TABLE:
                    objects.update_row(rid, {"DateUpdate": stamp})
                    break
        self._catalog = None

    def _new_map_rows(self, definition: TableDefinition, count: int) -> list[int]:
        """References to ``count`` fresh 69-byte usage-map rows for a table:
        on the first of its map pages with room (the page its own maps live
        on, then any page a later map spilled onto), else on a new map page
        the engine allocates the moment the current one is full (measured
        with a 58th map row: it went to row 0 of a fresh page)."""
        store = self.store
        candidates: list[int] = []
        refs = [definition.owned_pages_ref, definition.free_space_pages_ref]
        refs += [r.usage_map_ref for r in definition.real_indexes]
        for owned, free in definition.column_usage_maps.values():
            refs += [owned, free]
        for ref in refs:
            if ref >> 8 not in candidates:
                candidates.append(ref >> 8)
        out: list[int] = []
        while len(out) < count:
            page_number = next((p for p in candidates if DataPage(store.read(p)).fits(len(USAGE_MAP_ROW))), None)
            if page_number is None:
                page_number = allocate_page(store)
                store.write(page_number, usage_map_page(0))
                candidates.append(page_number)
            maps = DataPage(store.read(page_number))
            out.append((page_number << 8) | maps.add_row(USAGE_MAP_ROW))
            store.write(page_number, maps.to_bytes())
        return out

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

    def drop_index(self, table_name: str, name: str, *, updated: object | None = None) -> None:
        """Remove an index as DROP INDEX does: every page it holds released
        with its bytes left alone, its usage-map row deleted, the
        definition rewritten without its three records (the indexes after
        it move up and every logical index pointing past it follows), and
        the catalog row's DateUpdate stamped.  The primary key and an
        index a relationship rests on cannot be dropped, as the engine
        refuses them."""
        import datetime as _dt

        table = self.table(table_name)
        d = table.definition
        logical = next((li for li in d.logical_indexes if li.name.lower() == name.lower()), None)
        if logical is None:
            raise AccessError(f"table {table_name!r} has no index named {name!r}")
        if logical.is_primary_key:
            raise AccessError(f"index {name!r} is the primary key of {table_name!r}")
        if logical.relationship_kind:
            raise AccessError(f"index {name!r} belongs to a relationship; drop that first")
        position = logical.real_index
        if any(li is not logical and li.real_index == position for li in d.logical_indexes):
            raise AccessError(f"index {name!r} shares its B-tree with another index")
        store = self.store
        real = d.real_indexes[position]

        umap = read_usage_map_ref(store, real.usage_map_ref)
        pages = sorted(umap.pages())
        for page_number in pages:
            set_usage_bit(store, umap, page_number, False)
        map_page = DataPage(store.read(real.usage_map_ref >> 8))
        map_page.remove_row(real.usage_map_ref & 0xFF)
        store.write(real.usage_map_ref >> 8, map_page.to_bytes())
        for page_number in pages:
            release_page(store, page_number)

        d.real_indexes.pop(position)
        d.logical_indexes.remove(logical)
        d.real_index_count -= 1
        d.logical_index_count -= 1
        for other in d.logical_indexes:
            if other.real_index > position:
                # A logical index names its B-tree by position, so the ones
                # after the hole move down with it; its own number stays,
                # because relationships elsewhere point at that.
                other.real_index -= 1
                raw = bytearray(other.raw)
                struct.pack_into("<I", raw, 8, other.real_index)
                other.raw = bytes(raw)
        self._write_definition(serialize_definition(d), d.page, d.pages[1:], keep_tail=True)
        table._stamp_catalog(updated if isinstance(updated, (_dt.datetime, float)) else _dt.datetime.now().replace(microsecond=0))  # pyright: ignore[reportPrivateUsage]
        self._catalog = None
        self._definitions.pop(d.page, None)

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


    # --- the VBA project ----------------------------------------------------
    # A module lives in five storage streams and three catalog tables at
    # once.  See `pyopenvba.access._vba` for why writing here marks the
    # compiled cache stale rather than rebuilding it.

    def _vba_storage_ids(self) -> tuple[int, int, int]:
        """The ``Modules``, ``VBAProject`` and module-stream storage ids,
        walked by name from the root rather than assumed."""
        rows = [row for _rid, row in self.table(STORAGE_TABLE).rows_with_ids()]

        def child(parent: int, name: str) -> int:
            for row in rows:
                if (
                    _as_int(row["ParentId"]) == parent
                    and str(row["Name"]) == name
                    and _as_int(row["Type"]) == 1
                    and _as_int(row["Id"]) != parent
                ):
                    return _as_int(row["Id"])
            raise AccessError(f"MSysAccessStorage has no {name!r} folder")

        root = next(
            (_as_int(row["Id"]) for row in rows if str(row["Name"]) == "MSysAccessStorage_ROOT"),
            None,
        )
        if root is None:
            raise AccessError("this database has no VBA project")
        project = child(child(root, "VBA"), "VBAProject")
        return child(root, "Modules"), project, child(project, "VBA")

    def _vba_row(self, name: str, parent: int | None = None) -> tuple[RowId, dict[str, object]]:
        for rid, row in self.table(STORAGE_TABLE).rows_with_ids():
            if str(row["Name"]) == name and (parent is None or _as_int(row["ParentId"]) == parent):
                return rid, row
        raise AccessError(f"MSysAccessStorage has no {name!r} row")

    def _vba_dir(self) -> tuple[RowId, bytes]:
        """The decompressed dir stream and the row it came from."""
        rid, row = self._vba_row("dir")
        payload = row.get("Lv")
        if not isinstance(payload, bytes):
            raise AccessError("the dir stream row holds no value")
        return rid, decompress(payload)

    def _module_stream_row(self, dir_stream: bytes, name: str) -> tuple[RowId, bytes, int]:
        """A module's storage row, its bytes and its MODULEOFFSET."""
        _modules, _project, streams = self._vba_storage_ids()
        rid, row = self._vba_row(stream_name_of(dir_stream, name), streams)
        payload = row.get("Lv")
        if not isinstance(payload, bytes):
            raise AccessError(f"the stream row for {name!r} holds no value")
        at = module_offset_at(dir_stream, name)
        return rid, payload, int.from_bytes(dir_stream[at : at + 4], "little")

    def _invalidate_vba_cache(self) -> None:
        storage = self.table(STORAGE_TABLE)
        self._vba_changed = True
        for rid, row in list(storage.rows_with_ids()):
            payload = row.get("Lv")
            if str(row["Name"]) == "_VBA_PROJECT" and isinstance(payload, bytes) and payload:
                storage.update_row(rid, {"Lv": invalidate_cache(payload)})
                return
        raise AccessError("this database has no _VBA_PROJECT stream")

    def _drop_srp(self) -> int:
        """Retire the ``__SRP_*`` compiled cache rows, which is what Access
        runs in preference to the canonical p-code."""
        storage = self.table(STORAGE_TABLE)
        self._vba_changed = True
        doomed = [
            rid for rid, row in storage.rows_with_ids() if str(row["Name"]).startswith("__SRP_")
        ]
        for rid in doomed:
            storage.delete_row(rid, retire_empty=False)
        return len(doomed)

    def modules(self) -> list[VBAModule]:
        """Every module in the database's VBA project, in dir-stream order.

        ``source`` is the body without the leading ``Attribute`` block, so
        it reads as the VBE shows it.
        """
        _dir_rid, dir_stream = self._vba_dir()
        _modules, _project, streams = self._vba_storage_ids()
        payloads = {
            str(row["Name"]): row.get("Lv")
            for _rid, row in self.table(STORAGE_TABLE).rows_with_ids()
            if _as_int(row["ParentId"]) == streams
        }
        out: list[VBAModule] = []
        for name, stream_name, kind in module_blocks(dir_stream):
            payload = payloads.get(stream_name)
            at = module_offset_at(dir_stream, name)
            offset = int.from_bytes(dir_stream[at : at + 4], "little")
            text = read_source(payload, offset) if isinstance(payload, bytes) else ""
            _attributes, body = split_source(text)
            out.append(VBAModule(name, kind, stream_name, VBA_CRLF.join(body)))
        return out

    def module(self, name: str) -> VBAModule:
        """One module by name, case-insensitively as VBA compares names."""
        for module in self.modules():
            if module.name.lower() == name.lower():
                return module
        raise AccessError(f"the VBA project has no module named {name!r}")

    def create_module(
        self,
        name: str,
        code: str = "Option Compare Database",
        *,
        kind: str = "module",
        updated: object | None = None,
    ) -> VBAModule:
        """Add a module holding ``code``.

        ``kind`` is ``"module"`` for a standard module or ``"class"`` for a
        class module, which carries seven more source attributes and a
        different MODULETYPE.  The project is marked for recompilation, so
        the code has to compile when Access next opens the database.
        """
        if kind not in MODULETYPE:
            raise AccessError(f"kind must be 'module' or 'class', not {kind!r}")
        if not name or len(name) > 64:
            raise AccessError("a module name is 1 to 64 characters")
        if any(module.name.lower() == name.lower() for module in self.modules()):
            raise AccessError(f"a module named {name!r} already exists")

        rng = random.Random()
        storage = self.table(STORAGE_TABLE)
        modules_id, project_id, streams_id = self._vba_storage_ids()
        dir_rid, dir_stream = self._vba_dir()
        rows = [row for _rid, row in storage.rows_with_ids()]
        stream_name = stream_row_name(
            rng, {str(r["Name"]) for r in rows if _as_int(r["ParentId"]) == streams_id}
        )
        # Every module carries its own MODULEEND2 word; two sharing one is
        # not something Access writes.
        cookie = rng.randbytes(2)
        folders = {
            str(r["Name"])
            for r in rows
            if _as_int(r["ParentId"]) == modules_id and _as_int(r["Type"]) == 1
        }
        folder = next_folder("Modules", folders)
        when = (
            updated
            if isinstance(updated, (dt.datetime, float))
            else dt.datetime.now().replace(microsecond=0)
        )

        # Ids come from the table's own AutoNumber, not from max + 1: every
        # database Access wrote has the counter equal to its highest id, and
        # leaving it behind makes Access's own next insert collide.
        folder_rid = storage.insert_row(
            {"ParentId": modules_id, "Name": folder, "Type": 1, "DateCreate": when, "DateUpdate": when}
        )
        folder_id = next(_as_int(r["Id"]) for rid, r in storage.rows_with_ids() if rid == folder_rid)
        storage.insert_row(
            {"ParentId": folder_id, "Name": "PropData", "Type": 2, "Lv": PROP_DATA,
             "DateCreate": when, "DateUpdate": when}
        )
        storage.insert_row(
            {"ParentId": streams_id, "Name": stream_name, "Type": 2,
             "Lv": module_stream(attribute_lines(name, kind), code),
             "DateCreate": when, "DateUpdate": when}
        )

        for rid, row in list(storage.rows_with_ids()):
            payload = row.get("Lv")
            if not isinstance(payload, bytes) or not payload:
                continue
            row_name, parent = str(row["Name"]), _as_int(row["ParentId"])
            if row_name == "_VBA_PROJECT":
                storage.update_row(rid, {"Lv": invalidate_cache(payload)})
            elif row_name == "\x03DirData" and parent == modules_id:
                # The four bytes an entry ends with name the object's
                # storage folder, not a terminator.
                storage.update_row(rid, {"Lv": add_to_dir_data(payload, name, folder)})
            elif row_name == "PropData" and parent == modules_id:
                storage.update_row(rid, {"Lv": add_to_folder_list(payload, folder)})
            elif row_name == "PROJECTwm" and parent == project_id:
                storage.update_row(rid, {"Lv": add_to_project_wm(payload, name)})
            elif row_name == "PROJECT" and parent == project_id:
                storage.update_row(
                    rid,
                    {"Lv": add_to_project(payload.decode("latin-1"), name, kind).encode("latin-1")},
                )
        storage.update_row(
            dir_rid,
            {"Lv": compress(add_to_dir(dir_stream, dir_block(name, stream_name, cookie, kind)))},
        )

        objects = self.table("MSysObjects")
        container = next(e.id for e in self.catalog() if e.name == "Modules" and e.type == 3)
        owner = next((e.owner for e in self.catalog() if e.type == OBJECT_MODULE and e.owner), None)
        object_id = max((e.id for e in self.catalog() if e.id < 0), default=-(2**31)) + OBJECT_ID_STEP
        objects.insert_row(
            {"Id": object_id, "ParentId": container, "Name": name, "Type": OBJECT_MODULE,
             "Flags": 0, "Owner": owner, "DateCreate": when, "DateUpdate": when}
        )
        self.table("MSysNavPaneObjectIDs").insert_row(
            {"Id": object_id, "Name": name, "Type": NAV_MODULE_TYPE}
        )
        groups = self.table("MSysNavPaneGroupToObjects")
        peers = [
            r for _rid, r in groups.rows_with_ids() if _as_int(r["GroupID"]) == NAV_MODULE_GROUP
        ]
        groups.insert_row(
            {"GroupID": NAV_MODULE_GROUP, "ObjectID": object_id, "Flags": 0, "Icon": 0,
             "Position": max((_as_int(r["Position"]) for r in peers), default=-1) + 1}
        )

        self._drop_srp()
        self._catalog = None
        return self.module(name)

    def set_module_source(self, name: str, code: str) -> None:
        """Replace a module's body, keeping the attribute block it has.

        The stream becomes the source alone and the project is marked for
        recompilation, which is what makes the declaration forms a p-code
        writer cannot emit -- ``Const``, arrays, ``Static``, fixed-length
        strings, a whole new procedure -- reachable.
        """
        dir_rid, dir_stream = self._vba_dir()
        rid, payload, offset = self._module_stream_row(dir_stream, name)
        attributes, _body = split_source(read_source(payload, offset))
        storage = self.table(STORAGE_TABLE)
        storage.update_row(rid, {"Lv": module_stream(attributes, code)})
        storage.update_row(dir_rid, {"Lv": compress(set_module_offset(dir_stream, name, 0))})
        self._invalidate_vba_cache()
        self._drop_srp()

    def rename_module(self, name: str, new_name: str) -> None:
        """Rename a module in all eight places its name lives."""
        if not new_name or len(new_name) > 64:
            raise AccessError("a module name is 1 to 64 characters")
        module = self.module(name)
        if new_name.lower() != name.lower() and any(
            other.name.lower() == new_name.lower() for other in self.modules()
        ):
            raise AccessError(f"a module named {new_name!r} already exists")

        storage = self.table(STORAGE_TABLE)
        dir_rid, dir_stream = self._vba_dir()
        rid, payload, offset = self._module_stream_row(dir_stream, module.name)
        attributes, body = split_source(read_source(payload, offset))
        renamed = rename_attribute(VBA_CRLF.join(attributes), module.name, new_name)
        storage.update_row(
            rid, {"Lv": module_stream(renamed.split(VBA_CRLF), VBA_CRLF.join(body))}
        )
        stream = set_module_offset(rename_in_dir(dir_stream, module.name, new_name), new_name, 0)
        storage.update_row(dir_rid, {"Lv": compress(stream)})

        modules_id, project_id, _streams = self._vba_storage_ids()
        for row_rid, row in list(storage.rows_with_ids()):
            value = row.get("Lv")
            if not isinstance(value, bytes) or not value:
                continue
            row_name, parent = str(row["Name"]), _as_int(row["ParentId"])
            if row_name == "_VBA_PROJECT":
                storage.update_row(row_rid, {"Lv": invalidate_cache(value)})
            elif row_name == "\x03DirData" and parent == modules_id:
                storage.update_row(row_rid, {"Lv": rename_dir_data(value, module.name, new_name)})
            elif row_name == "PROJECTwm" and parent == project_id:
                storage.update_row(row_rid, {"Lv": rename_project_wm(value, module.name, new_name)})
            elif row_name == "PROJECT" and parent == project_id:
                text = value.decode("latin-1")
                fixed = rename_project(text, module.name, new_name)
                if fixed != text:
                    storage.update_row(row_rid, {"Lv": fixed.encode("latin-1")})

        objects = self.table("MSysObjects")
        for row_rid, row in objects.rows_with_ids():
            if row["Type"] == OBJECT_MODULE and str(row["Name"]) == module.name:
                objects.update_row(row_rid, {"Name": new_name})
                break
        nav = self.table("MSysNavPaneObjectIDs")
        for row_rid, row in nav.rows_with_ids():
            if str(row["Name"]) == module.name:
                nav.update_row(row_rid, {"Name": new_name})
                break
        self._drop_srp()
        self._catalog = None

    def delete_module(self, name: str) -> None:
        """Remove a module and every structure it occupies."""
        module = self.module(name)
        index = [m.name for m in self.modules()].index(module.name)
        storage = self.table(STORAGE_TABLE)
        modules_id, project_id, streams_id = self._vba_storage_ids()

        # A module's storage folder is linked to it by position, not by
        # name: deleting the second of three modules, Access dropped the
        # second folder and left the others named as they were.
        folders = sorted(
            (
                (rid, row)
                for rid, row in storage.rows_with_ids()
                if _as_int(row["ParentId"]) == modules_id and _as_int(row["Type"]) == 1
            ),
            key=lambda pair: _as_int(pair[1]["Id"]),
        )
        folder_rid, folder_row = folders[index]
        folder_name, folder_id = str(folder_row["Name"]), _as_int(folder_row["Id"])
        doomed = [folder_rid] + [
            rid for rid, row in storage.rows_with_ids() if _as_int(row["ParentId"]) == folder_id
        ]

        for rid, row in list(storage.rows_with_ids()):
            value = row.get("Lv")
            row_name, parent = str(row["Name"]), _as_int(row["ParentId"])
            if rid in doomed or (parent == streams_id and row_name == module.stream_name):
                storage.delete_row(rid, retire_empty=False)
                continue
            if not isinstance(value, bytes) or not value:
                continue
            if row_name == "dir":
                storage.update_row(
                    rid, {"Lv": compress(remove_from_dir(decompress(value), module.name))}
                )
            elif row_name == "_VBA_PROJECT":
                storage.update_row(rid, {"Lv": invalidate_cache(value)})
            elif row_name == "\x03DirData" and parent == modules_id:
                storage.update_row(rid, {"Lv": remove_from_dir_data(value, module.name)})
            elif row_name == "PropData" and parent == modules_id:
                storage.update_row(rid, {"Lv": remove_from_folder_list(value, folder_name)})
            elif row_name == "PROJECTwm" and parent == project_id:
                storage.update_row(rid, {"Lv": remove_from_project_wm(value, module.name)})
            elif row_name == "PROJECT" and parent == project_id:
                text = value.decode("latin-1")
                fixed = remove_from_project(text, module.name)
                if fixed != text:
                    storage.update_row(rid, {"Lv": fixed.encode("latin-1")})

        objects = self.table("MSysObjects")
        for rid, row in list(objects.rows_with_ids()):
            if row["Type"] == OBJECT_MODULE and str(row["Name"]) == module.name:
                object_id = _as_int(row["Id"])
                objects.delete_row(rid, retire_empty=False)
                nav = self.table("MSysNavPaneObjectIDs")
                for nav_rid, nav_row in list(nav.rows_with_ids()):
                    if _as_int(nav_row["Id"]) == object_id:
                        nav.delete_row(nav_rid, retire_empty=False)
                groups = self.table("MSysNavPaneGroupToObjects")
                for group_rid, group_row in list(groups.rows_with_ids()):
                    if _as_int(group_row["ObjectID"]) == object_id:
                        groups.delete_row(group_rid, retire_empty=False)
                break
        self._drop_srp()
        self._catalog = None


def _macro_properties() -> bytes:
    """The property blob a macro's catalog row carries: `PublishToWeb`,
    true, which is what Access writes and nothing else."""
    return serialize_property_blob(
        PropertyBlob(
            names=["PublishToWeb"],
            object_properties={"PublishToWeb": PropertyValue(type=DB_BOOLEAN, flags=0, raw=b"\x01")},
            column_properties={},
            block_order=[(BLOCK_OBJECT, "")],
        )
    )


def _design_properties(kind: str, guid: bytes) -> bytes:
    """A form's or report's catalog properties: the captured blob with a
    GUID of its own, which the design blob repeats."""
    blob = parse_property_blob(template(kind, "lvprop"))
    existing = blob.object_properties["GUID"]
    blob.object_properties["GUID"] = PropertyValue(type=existing.type, flags=existing.flags, raw=guid)
    return serialize_property_blob(blob)


def _converter(old_code: int, new_code: int) -> Callable[[object], object]:
    """How a value moves between column types in ALTER COLUMN: unchanged
    within a type family, numbers converted between integer and floating
    kinds, everything else refused."""
    from decimal import Decimal

    integers = {TYPE_BOOLEAN, 0x02, 0x03, 0x04, 0x13}
    floats = {0x06, 0x07}
    if old_code == new_code:
        return lambda value: value
    if old_code in integers and new_code in integers:
        return lambda value: int(value)  # pyright: ignore[reportArgumentType]
    if old_code in integers and new_code in floats:
        return lambda value: float(value)  # pyright: ignore[reportArgumentType]
    if old_code in floats and new_code in integers:
        return lambda value: round(float(value))  # pyright: ignore[reportArgumentType]
    if old_code in floats and new_code in floats:
        return lambda value: float(value)  # pyright: ignore[reportArgumentType]
    if old_code in integers | floats and new_code == 0x05:
        return lambda value: Decimal(str(value))
    if old_code == 0x05 and new_code in floats:
        return lambda value: float(value)  # pyright: ignore[reportArgumentType]
    raise AccessError(f"converting column type {old_code:#04x} to {new_code:#04x} is not written")


#: What the engine records as the least Access build that can handle a
#: table holding a BigInt column.
BIGINT_MIN_VERSION = "16.0.7124.1000"
VERSION_PROPERTIES = ("FCMinReadVer", "FCMinWriteVer", "FCMinDesignVer")


def _version_properties(count: int = len(VERSION_PROPERTIES)) -> PropertyBlob:
    """The version properties a BigInt column brings with it, the first
    ``count`` of them."""
    blob = PropertyBlob()
    for name in VERSION_PROPERTIES[:count]:
        blob.names.append(name)
        blob.object_properties[name] = PropertyValue(type=DB_TEXT, flags=0, raw=BIGINT_MIN_VERSION.encode("utf-16-le"))
    blob.block_order.append((0, ""))
    return blob


def _write_column_properties(table: Table, columns: Sequence[ColumnSpec], updated: object | None) -> None:
    """Write the properties the specs ask for, one blob write per column in
    column order, which is the engine's own order for a CREATE TABLE whose
    columns are NOT NULL (measured).  Each write carries the stamp the
    statement already put on the catalog row, so it lands where the
    engine left it."""
    for spec in columns:
        properties = spec.properties()
        if properties:
            table.set_properties(properties, column=spec.name, updated=updated)


def _passthrough_rows(connect: str, sql: str) -> list[QueryRow]:
    """The three rows a pass-through query has: the two every query has and
    a type row of flag 8 holding the connect string and the SQL the server
    is sent, stored exactly as given."""
    return [
        QueryRow(ATTR_START, 1, flag=0),
        QueryRow(ATTR_END, 1),
        QueryRow(ATTR_TYPE, 1, name1=connect, expression=sql, flag=QUERY_PASSTHROUGH),
    ]


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


def _as_text(value: object) -> str | None:
    """A column property as the spec wants it: text, or nothing."""
    return None if value is None else str(value)


def _as_int(value: object) -> int:
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, int):
        return value
    raise AccessError(f"expected an integer catalog field, got {type(value).__name__}")
