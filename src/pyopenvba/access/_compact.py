"""Compact and Repair, done the way DAO's ``CompactDatabase`` does it.

The engine does not rewrite a file in place.  It creates a bare database
(what ``DBEngine.CreateDatabase`` makes, before that routine writes its
permission rows) and copies every object into it, so the result is laid
out by the copy and not by the source's history: pages are numbered as
they are taken, rows sit on as few pages as they need, and nothing dead
is carried across.  Measured against ACE 16 on a dozen databases, the
copy runs in these phases, and each phase here is one of the writers the
engine was already measured against:

1. The Access containers (``Forms``, ``Modules``, ...) and then each
   container's objects, containers in the order of the catalog's Name
   index, objects likewise; the ``Relationships`` container waits for
   the end.  Forms, reports, macros and modules are catalog rows plus
   their streams in ``MSysAccessStorage``, which is copied as a table.
2. The ``Tables`` container: tables, queries and links together in Name
   order.  A table is created as ``CREATE TABLE`` would create it from
   its definition (no indexes), its rows are written in primary-key
   order (a table without a primary key keeps its stored order), then
   each index is created over the rows in the order the definition
   lists them.  The AutoNumber counter comes out at the largest value
   present, the complex-id counter one past the source's.  A query is
   its catalog row and its ``MSysQueries`` rows, the type and end rows
   first and the rest by attribute; a link is its catalog row.
3. The ``MSysComplexColumns`` rows, with the table ids translated.
4. A pass over every object in the same order as 1 and 2, containers
   first, that writes each object's permission rows and then its
   catalog row's flags, stamps, owner and property blob -- which is why
   a row written in 1 or 2 carries Flags 0, the compaction time in both
   date columns, no owner and no blob.
5. The relationships, in Name order, each as ``ADD CONSTRAINT`` adds one,
   stamping both tables.

The one thing not derived from the source is the encoding of the owner
and permission SIDs, which the engine keys to the file's creation date:
the destination keeps the source's creation date, so those bytes copy
across unchanged and the file stays consistent with itself.
"""

from __future__ import annotations

import math
import struct
from collections.abc import Callable, Iterable
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import TYPE_CHECKING

from pyopenvba.access._index import encode_key
from pyopenvba.access._lval import read_long_value, write_long_value
from pyopenvba.access._pages import PAGE_SIZE, FileGrowth
from pyopenvba.access._rows import LongValueRef, decode_long_value_ref, split_row
from pyopenvba.access._tdef import (
    OFFSET_INDEX_HEADERS,
    OFFSET_LAST_COMPLEX_ID,
    OFFSET_NEXT_AUTONUMBER,
    SIZE_COLUMN_HEADER,
    SIZE_REAL_INDEX_HEADER,
    TYPE_BOOLEAN,
    TYPE_COMPLEX,
    ColumnDef,
)
from pyopenvba.access_read import AccessError

if TYPE_CHECKING:
    from pyopenvba.access.database import AccessDatabase, CatalogEntry, RowId, Table

#: The bare database the engine copies into: ``CreateDatabase`` with the
#: permission rows it writes last taken back out (their data page, their
#: index root, the counters and the map bits), as measured -- one for the
#: ACE format and one for Jet 4.
_TEMPLATES = Path(__file__).resolve().parents[1] / "_templates" / "blank_files"
SKELETON = _TEMPLATES / "engine_skeleton.accdb"
SKELETON_JET4 = _TEMPLATES / "engine_skeleton.mdb"

#: MSysObjects.Type of the object kinds a catalog row can hold.
_TYPE_TABLE = 1
_TYPE_CONTAINER = 3
_TYPE_QUERY = 5
_TYPE_LINK = 6
_TYPE_RELATIONSHIP = 8

#: The long-value catalog columns copied with an object's row; the owner
#: and the property blob are written in the last pass instead.
_ROW_EXTRAS = ("Connect", "Database", "ForeignName", "Lv", "LvExtra", "LvModule", "RmtInfoLong", "RmtInfoShort")

#: A query's type row and end row lead its MSysQueries rows; the rest
#: follow by attribute, then by order (measured on seven query shapes).
_ATTR_TYPE = 0
_ATTR_END = 255


@dataclass
class _Source:
    """What the compaction reads from the source once."""

    entries: list[CatalogEntry]
    rows: dict[int, dict[str, object]] = field(default_factory=lambda: {})
    aces: dict[int, list[dict[str, object]]] = field(default_factory=lambda: {})


def compact_and_repair(
    source: AccessDatabase, *, clock: Callable[[], float] | None = None, lval_stamp: int | None = None
) -> AccessDatabase:
    """Copy ``source`` into a fresh engine skeleton as the engine's Compact
    and Repair does; see the module docstring for the phases.  ``clock``
    supplies the DateUpdate the relationship pass stamps on each table
    (the engine reads its clock once per stamp); it defaults to now."""
    from pyopenvba.access.database import AccessDatabase as _Database

    skeleton = _Database((SKELETON if source.header.is_ace else SKELETON_JET4).read_bytes())
    same = (source.header.code_page, source.header.sort_order, source.header.sort_version)
    wanted = (skeleton.header.code_page, skeleton.header.sort_order, skeleton.header.sort_version)
    if same != wanted:
        raise AccessError(
            f"compaction is measured for the General sort order and code page 1252 only; "
            f"this database has code page {same[0]} and sort order {same[1]}/{same[2]}"
        )
    # The header page is carried across whole: sort order, password and
    # the creation date the SID encoding is keyed to.
    skeleton.store.write(0, source.store.read(0))
    skeleton.header = source.header
    # The destination grows, and lets rewritten pages back into use, the
    # way the engine's does: sized 64 pages ahead in the ACE format and 32
    # in Jet 4 (measured on wide tables in bare databases of each), the
    # bare file's own pages and the permission page it took back counting
    # as taken.
    skeleton.store.growth = FileGrowth(
        physical=64 if source.header.is_ace else 32, since_release=skeleton.store.page_count + 1
    )
    if lval_stamp is not None:
        skeleton.lval_stamp = lval_stamp
    run = _Compaction(source, skeleton, clock)
    run.copy_containers()
    run.copy_container_objects()
    run.copy_tables_container()
    run.copy_complex_columns()
    run.write_owners_and_permissions()
    run.copy_relationships()
    return skeleton


class _Compaction:
    def __init__(self, source: AccessDatabase, dest: AccessDatabase, clock: Callable[[], float] | None) -> None:
        import datetime as _dt

        self.source = source
        self.dest = dest
        self.clock: Callable[[], float] = clock or (lambda: _serial(_dt.datetime.now()))
        #: Source object id -> destination object id.
        self.ids: dict[int, int] = {}
        self.src = _Source(list(source.catalog()))
        objects = source.table("MSysObjects")
        for _rid, row in objects.rows_with_ids():
            self.src.rows[_as_int(row["Id"])] = row
        for row in source.table("MSysACEs").rows():
            self.src.aces.setdefault(_as_int(row["ObjectId"]), []).append(row)
        name_column = dest.table("MSysObjects").definition.column("Name")
        self._name_column = name_column
        # Containers keep their ids; so do the objects the skeleton has.
        for entry in self.src.entries:
            if entry.type == _TYPE_CONTAINER or self._in_skeleton(entry) is not None:
                self.ids[entry.id] = entry.id

    # -- ordering ---------------------------------------------------------

    def _name_key(self, entry: CatalogEntry) -> bytes:
        """The engine walks the catalog's Name index, so its order is the
        index key's."""
        return encode_key([entry.name], [(self._name_column, True)])

    def _sorted(self, entries: Iterable[CatalogEntry]) -> list[CatalogEntry]:
        return sorted(entries, key=self._name_key)

    def _containers(self) -> list[CatalogEntry]:
        return self._sorted(e for e in self.src.entries if e.type == _TYPE_CONTAINER)

    def _children(self, container: CatalogEntry) -> list[CatalogEntry]:
        return self._sorted(e for e in self.src.entries if e.parent_id == container.id and e.type != _TYPE_CONTAINER)

    def _in_skeleton(self, entry: CatalogEntry) -> CatalogEntry | None:
        for candidate in self.dest.catalog():
            if candidate.type == entry.type and candidate.name.lower() == entry.name.lower() and candidate.id == entry.id:
                return candidate
        return None

    # -- phase 1: containers and their objects ----------------------------

    def copy_containers(self) -> None:
        objects = self.dest.table("MSysObjects")
        for entry in self._containers():
            if self._in_skeleton(entry) is not None:
                continue
            objects.insert_row(self._row_values(entry, entry.id))
        self.dest.forget_catalog()

    def _creation_stamp(self) -> float:
        """The stamp a row is first written with, in both date columns: the
        compaction time.  The source's own stamps replace it in the last
        pass; see :meth:`_copy_table` for what an index does to it."""
        return self.clock()

    def copy_container_objects(self) -> None:
        objects = self.dest.table("MSysObjects")
        for container in self._containers():
            if container.name in ("Relationships", "Tables"):
                continue
            for entry in self._children(container):
                if self._in_skeleton(entry) is not None:
                    continue  # MSysDb: the skeleton's row, restamped in the last pass
                new_id = self.dest._next_object_id()  # pyright: ignore[reportPrivateUsage]
                self.ids[entry.id] = new_id
                objects.insert_row(self._row_values(entry, new_id))
                self.dest.forget_catalog()

    def _row_values(self, entry: CatalogEntry, new_id: int) -> dict[str, object]:
        """A catalog row as the copy first writes it: everything but the
        owner and the property blob."""
        source_row = self.src.rows[entry.id]
        now = self._creation_stamp()
        values: dict[str, object] = {
            "Id": new_id,
            "ParentId": self.ids.get(entry.parent_id, entry.parent_id),
            "Name": entry.name,
            "Type": entry.type,
            "Flags": 0,
            "DateCreate": now,
            "DateUpdate": now,
        }
        for column in _ROW_EXTRAS:
            value = source_row.get(column)
            if value is not None:
                values[column] = value
        return values

    def _dest_row(self, object_id: int) -> RowId:
        objects = self.dest.table("MSysObjects")
        for rid, row in objects.rows_with_ids():
            if _as_int(row["Id"]) == object_id:
                return rid
        raise AccessError(f"no catalog row for object {object_id}")

    # -- phase 2: the Tables container ------------------------------------

    def copy_tables_container(self) -> None:
        tables = next(e for e in self.src.entries if e.type == _TYPE_CONTAINER and e.name == "Tables")
        for entry in self._children(tables):
            if entry.type == _TYPE_TABLE:
                if self._in_skeleton(entry) is not None:
                    continue
                self._copy_table(entry)
            elif entry.type in (_TYPE_QUERY, _TYPE_LINK):
                self._copy_catalog_object(entry)
                if entry.type == _TYPE_QUERY:
                    self._copy_query_rows(entry)
            else:
                raise AccessError(f"the Tables container holds an object of type {entry.type} ({entry.name!r}), which this cannot copy")

    def _copy_catalog_object(self, entry: CatalogEntry) -> None:
        new_id = self.dest._next_object_id()  # pyright: ignore[reportPrivateUsage]
        self.ids[entry.id] = new_id
        self.dest.table("MSysObjects").insert_row(self._row_values(entry, new_id))
        self.dest.forget_catalog()

    def _copy_query_rows(self, entry: CatalogEntry) -> None:
        source_rows = [row for row in self.source.table("MSysQueries").rows() if _as_int(row["ObjectId"]) == entry.id]

        def rank(row: dict[str, object]) -> tuple[int, int, bytes]:
            attribute = _as_int(row["Attribute"])
            lead = 0 if attribute == _ATTR_TYPE else 1 if attribute == _ATTR_END else 2
            order = row.get("Order")
            return (lead, attribute, order if isinstance(order, bytes) else b"")

        queries = self.dest.table("MSysQueries")
        for row in sorted(source_rows, key=rank):
            values = {k: v for k, v in row.items() if v is not None}
            values["ObjectId"] = self.ids[entry.id]
            queries.insert_row(values)

    def _copy_table(self, entry: CatalogEntry) -> None:
        source_table = self.source.table(entry.name)
        definition = source_table.definition
        specs, indexes = self.source.table_specs(entry.name)
        # The copy numbers the columns densely in the order the definition
        # lists them, which need not be their old number order (an
        # attachment's flat table lists its key column first under the
        # highest number); bytes 9-10 of a header survive as they are.
        # No column property goes with the creation; the whole blob is
        # written in the last pass.
        by_name = {spec.name.lower(): spec for spec in specs}
        columns = [
            replace(
                by_name[column.name.lower()],
                ordinal=column.header_ordinal,
                required=False,
                allow_zero_length=None,
                default=None,
                validation_rule=None,
                validation_text=None,
            )
            for column in definition.columns
        ]
        owned_only = [
            definition.column_by_number(number).name
            for number, (_owned, free) in definition.column_usage_maps.items()
            if free == 0
        ]
        stamp = self._creation_stamp()
        table = self.dest._create_table_shell(  # pyright: ignore[reportPrivateUsage]
            entry.name,
            columns,
            flags=0,
            created=stamp,
            updated=stamp,
            owned_only=owned_only,
        )
        self.ids[entry.id] = table.definition.page
        self._patch_headers(source_table, table)
        self._copy_rows(source_table, table)
        self._set_counters(source_table, table)
        # A definition that runs past one page is written out once more
        # after the rows, and once more after each index is added, each
        # time onto a fresh continuation page (measured: the pages left
        # behind a wide table hold the version before the index twice and
        # the version with it twice).
        self._rewrite_if_long(table)
        creation = stamp
        for spec in indexes:
            # Each index stamps the row with the compaction time again; when
            # that is the stamp the row already carries, the engine writes the
            # double one step past it instead (measured under a frozen clock:
            # one index leaves the row a step past its creation stamp, two
            # leave it at the creation stamp).
            stamp = math.nextafter(creation, math.inf) if stamp == creation else creation
            self.dest.create_index(entry.name, spec, updated=stamp)
            self._rewrite_if_long(self.dest.table(entry.name))
        self._patch_index_tails(source_table, self.dest.table(entry.name))

    def _rewrite_if_long(self, table: Table) -> None:
        from pyopenvba.access._schema import serialize_definition

        d = table.definition
        if len(d.pages) > 1:
            self.dest._write_definition(serialize_definition(d), d.page, d.pages[1:], keep_tail=True)  # pyright: ignore[reportPrivateUsage]
            self.dest._definitions.pop(d.page, None)  # pyright: ignore[reportPrivateUsage]

    def _patch_headers(self, source_table: Table, table: Table) -> None:
        """Header bits a spec cannot ask for -- the misc flags a complex
        column's bookkeeping carries -- are copied from the source.  The
        engine's build carries them from the start, so they are patched
        where they lie rather than by a rewrite, which on a definition
        past one page would take a page the engine does not take."""
        from pyopenvba.access._complex import patch_column_header
        from pyopenvba.access._schema import serialize_definition

        d = table.definition
        by_name = {c.name.lower(): c for c in source_table.definition.columns}
        patched: list[int] = []
        for position, column in enumerate(d.columns):
            original = by_name[column.name.lower()]
            if original.misc_flags != column.misc_flags:
                column.raw = patch_column_header(column.raw, misc_flags=original.misc_flags)
                column.misc_flags = original.misc_flags
                patched.append(position)
        if not patched:
            return
        headers_end = OFFSET_INDEX_HEADERS + len(d.real_indexes) * SIZE_REAL_INDEX_HEADER + len(d.columns) * SIZE_COLUMN_HEADER
        if headers_end <= PAGE_SIZE:
            for position in patched:
                self.dest.patch_column_header(d, position, d.columns[position].raw)
        else:
            self.dest._write_definition(serialize_definition(d), d.page, d.pages[1:], keep_tail=True)  # pyright: ignore[reportPrivateUsage]
            self.dest._definitions.pop(d.page, None)  # pyright: ignore[reportPrivateUsage]

    def _copy_rows(self, source_table: Table, table: Table) -> None:
        definition = source_table.definition
        primary = next((i for i in source_table.indexes if i.is_primary_key), None)
        if primary is not None:
            rows: Iterable[bytes | None] = (source_table.fetch_row(page, row) for _key, page, row in primary.entries())
        else:
            rows = (data for _page, _slot, data in source_table.raw_rows())
        target = table.definition
        by_name = {c.name.lower(): c for c in target.columns}
        for data in rows:
            if data is None:
                continue
            parts = split_row(definition, data)
            encoded: dict[int, bytes | None] = {}
            booleans: set[int] = set()
            for column in definition.columns:
                destination = by_name.get(column.name.lower())
                if destination is None:
                    continue
                if column.type_code == TYPE_BOOLEAN:
                    if parts.present.get(column.number):
                        booleans.add(destination.number)
                    continue
                raw = parts.values.get(column.number)
                if raw is None:
                    encoded[destination.number] = None
                elif column.is_long_value:
                    encoded[destination.number] = self._copy_long_value(raw, table, destination)
                else:
                    encoded[destination.number] = raw
            table._append_encoded(encoded, booleans)  # pyright: ignore[reportPrivateUsage]

    def _copy_long_value(self, raw: bytes, table: Table, column: ColumnDef) -> bytes:
        ref = decode_long_value_ref(raw)
        if ref.kind == LongValueRef.KIND_INLINE:
            return raw
        data = read_long_value(self.source.store, ref)
        return write_long_value(self.dest.store, table._long_value_maps(column), data, self.dest.lval_stamp)  # pyright: ignore[reportPrivateUsage]

    def _patch_index_tails(self, source_table: Table, table: Table) -> None:
        """The bytes past an index entry's flags, which a spec cannot ask
        for (the complex column's own unique index carries a 2 there),
        are copied from the source's entry of the same name."""
        from pyopenvba.access._schema import serialize_definition

        d = table.definition
        by_name = {li.name.lower(): source_table.definition.real_indexes[li.real_index] for li in source_table.definition.logical_indexes}
        changed = False
        for li in d.logical_indexes:
            original = by_name.get(li.name.lower())
            if original is None:
                continue
            real = d.real_indexes[li.real_index]
            if real.raw[47:] != original.raw[47:]:
                real.raw = real.raw[:47] + original.raw[47:]
                changed = True
        if changed:
            self.dest._write_definition(serialize_definition(d), d.page, d.pages[1:], keep_tail=True)  # pyright: ignore[reportPrivateUsage]
            self.dest._definitions.pop(d.page, None)  # pyright: ignore[reportPrivateUsage]

    def _set_counters(self, source_table: Table, table: Table) -> None:
        """The AutoNumber counter comes out at the largest value present
        (zero when no row is) with the source's increment, the complex-id
        counter one past the source's; all measured."""
        from pyopenvba.access._tdef import OFFSET_COMPLEX_MARKER

        d = table.definition
        increment = source_table.definition.complex_marker
        if increment != d.complex_marker:
            d.complex_marker = increment
            self.dest.patch_definition(d, OFFSET_COMPLEX_MARKER, struct.pack("<i", increment))
        for column in d.columns:
            if column.auto_number and column.type_code != TYPE_COMPLEX:
                largest = 0
                for row in table.rows():
                    value = row.get(column.name)
                    if isinstance(value, int) and value > largest:
                        largest = value
                d.next_autonumber = largest & 0xFFFFFFFF
                self.dest.patch_definition(d, OFFSET_NEXT_AUTONUMBER, struct.pack("<I", d.next_autonumber))
        if any(c.type_code == TYPE_COMPLEX for c in d.columns):
            d.last_complex_id = source_table.definition.last_complex_id + 1
            self.dest.patch_definition(d, OFFSET_LAST_COMPLEX_ID, struct.pack("<I", d.last_complex_id))

    # -- phase 3: complex columns -----------------------------------------

    def copy_complex_columns(self) -> None:
        if not any(e.name == "MSysComplexColumns" for e in self.dest.catalog()):
            return  # a Jet 4 file has no complex columns
        complex_columns = self.dest.table("MSysComplexColumns")
        for row in self.source.table("MSysComplexColumns").rows():
            values = {k: v for k, v in row.items() if v is not None}
            for column in ("ConceptualTableID", "FlatTableID"):
                values[column] = self.ids[_as_int(row[column])]
            complex_columns.insert_row(values)

    # -- phase 4: owners and permissions ----------------------------------

    def write_owners_and_permissions(self) -> None:
        order: list[CatalogEntry] = list(self._containers())
        for container in self._containers():
            if container.name != "Relationships":
                order.extend(self._children(container))
        objects = self.dest.table("MSysObjects")
        aces = self.dest.table("MSysACEs")
        for entry in order:
            new_id = self.ids[entry.id]
            source_row = self.src.rows[entry.id]
            changes: dict[str, object] = {
                "Flags": _signed(entry.flags),
                "DateCreate": entry.date_create_serial,
                "DateUpdate": entry.date_update_serial,
            }
            owner = source_row.get("Owner")
            if owner is not None:
                changes["Owner"] = owner
            blob = source_row.get("LvProp")
            if blob is not None:
                changes["LvProp"] = blob
            objects.update_row(self._dest_row(new_id), changes)
            for ace in self.src.aces.get(entry.id, []):
                aces.insert_row({"SID": ace["SID"], "ACM": ace["ACM"], "FInheritable": ace["FInheritable"], "ObjectId": new_id})
        self.dest.forget_catalog()

    # -- phase 5: relationships -------------------------------------------

    def copy_relationships(self) -> None:
        from pyopenvba.access.database import RELATION_DELETE_CASCADE, RELATION_UPDATE_CASCADE

        relationships = {r.name.lower(): r for r in self.source.relationships()}
        container = next(e for e in self.src.entries if e.type == _TYPE_CONTAINER and e.name == "Relationships")
        for entry in self._children(container):
            relation = relationships.get(entry.name.lower())
            if relation is None:
                raise AccessError(f"the Relationships container names {entry.name!r}, which has no MSysRelationships rows")
            # The referencing table is stamped first; the referenced table's
            # stamp is read after it and steps one double past it when the
            # clock has not moved (measured under a frozen clock).
            table_stamp = self.clock()
            referenced_stamp = self.clock()
            current = next(e.date_update_serial for e in self.dest.catalog() if e.type == _TYPE_TABLE and e.name.lower() == relation.referenced_table.lower())
            if relation.table.lower() == relation.referenced_table.lower() or current == table_stamp:
                # One table, one stamp; and a referenced table already carrying
                # this compaction's stamp keeps it (measured: a table that is a
                # foreign key's referencing side and another's referenced side).
                referenced_stamp = table_stamp
            elif referenced_stamp <= table_stamp:
                referenced_stamp = math.nextafter(table_stamp, math.inf)
            self.dest.create_relationship(
                relation.name,
                relation.table,
                relation.columns,
                relation.referenced_table,
                relation.referenced_columns,
                cascade_updates=bool(relation.attributes & RELATION_UPDATE_CASCADE),
                cascade_deletes=bool(relation.attributes & RELATION_DELETE_CASCADE),
                created=entry.date_create_serial,
                updated=entry.date_update_serial,
                table_updated=table_stamp,
                referenced_updated=referenced_stamp,
                permissions=self.src.aces.get(entry.id),
                owner=_owner_bytes(self.src.rows[entry.id].get("Owner")),
            )
            self.ids[entry.id] = next(e.id for e in self.dest.catalog() if e.type == _TYPE_RELATIONSHIP and e.name == entry.name)


def _serial(when: object) -> float:
    """A datetime as the stored serial (days since 1899-12-30)."""
    import datetime as _dt

    if isinstance(when, (int, float)):
        return float(when)
    if isinstance(when, _dt.datetime):
        delta = when - _dt.datetime(1899, 12, 30)
        return delta.days + delta.seconds / 86400 + delta.microseconds / 86400e6
    raise AccessError(f"{when!r} is not a datetime")


def _owner_bytes(value: object) -> bytes | None:
    return value if isinstance(value, bytes) else None


def _signed(value: int) -> int:
    """A catalog Flags value as the signed Long the row stores."""
    return value - (1 << 32) if value >= 1 << 31 else value


def _as_int(value: object) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise AccessError(f"expected an integer, got {value!r}")
    return value

