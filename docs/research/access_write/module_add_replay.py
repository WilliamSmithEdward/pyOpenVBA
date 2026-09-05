"""Replay an Access module add from a before/after pair, deriving every row
and payload from the pair, in the order the engine's pages showed:

  1. existing rows re-stamped (a new long value waits, an inline one goes now)
  2. the replaced rows deleted in id order -- their long values stay in
     place for now; a temporary row takes the skipped id and goes away;
     the new rows in id order with their inline values, long values held
  3. the new streams' values: module and __SRP streams by id, then dir,
     then _VBA_PROJECT (PROJECT waits)
  4. the rewritten streams (__SRP_0, __SRP_1) and PROJECT
  5. the replaced rows' long values given back, in id order
  6. a PropDataCopy row under the new storage comes and goes
  7. the catalog: the module's row, MSysDb's blob nine times, permissions,
     the navigation pane

    python docs/research/access_write/module_add_replay.py before.accdb after.accdb

``before.accdb`` is a database, ``after.accdb`` a copy of it after Access
added one module (``module_ops.ps1 -Op add``); the script rebuilds the
after file from the before file with the engine writers and reports the
pages that differ, which on the measured pairs is none.
"""

import struct
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "src"))
from pyopenvba.access import AccessDatabase  # noqa: E402
from pyopenvba.access._datapage import DataPage  # noqa: E402
from pyopenvba.access._lval import free_long_value  # noqa: E402
from pyopenvba.access._rows import decode_long_value_ref, split_row  # noqa: E402
from pyopenvba.access._tdef import OFFSET_ROW_COUNT  # noqa: E402

PAGE = 4096
INLINE_MAX = 64


def rows_by_id(db: AccessDatabase, table: str) -> dict[int, tuple[object, dict[str, object]]]:
    out: dict[int, tuple[object, dict[str, object]]] = {}
    for rid, row in db.table(table).rows_with_ids():
        out[int(str(row["Id"]))] = (rid, dict(row))
    return out


def serials(db: AccessDatabase, table: str, ident: int) -> dict[str, float]:
    t = db.table(table)
    d = t.definition
    for _page, _slot, raw in t.raw_rows():
        parts = split_row(d, raw)
        if int(str(t.decode(parts)["Id"])) == ident:
            return {c: struct.unpack("<d", parts.values[d.column(c).number])[0] for c in ("DateCreate", "DateUpdate")}
    raise KeyError(ident)


def replay(before_path: Path, after_path: Path) -> tuple[bytes, bytes]:
    after = AccessDatabase(after_path)
    db = AccessDatabase(before_path)
    st = db.table("MSysAccessStorage")
    b_rows, a_rows = rows_by_id(db, "MSysAccessStorage"), rows_by_id(after, "MSysAccessStorage")
    stamp = lambda ident: serials(after, "MSysAccessStorage", ident)  # noqa: E731
    new_ids = sorted(set(a_rows) - set(b_rows))
    gone_ids = sorted(set(b_rows) - set(a_rows))
    long_new = {i for i in new_ids if isinstance(a_rows[i][1].get("Lv"), bytes) and len(a_rows[i][1]["Lv"]) > INLINE_MAX}
    changed = [i for i in sorted(set(b_rows) & set(a_rows)) if b_rows[i][1] != a_rows[i][1]]

    # 1. existing rows: stamps now; a long value waits
    long_rewritten: list[int] = []
    for ident in changed:
        new = a_rows[ident][1]
        values: dict[str, object] = {"DateUpdate": stamp(ident)["DateUpdate"]}
        if new.get("Lv") != b_rows[ident][1].get("Lv"):
            if isinstance(new.get("Lv"), bytes) and len(new["Lv"]) > INLINE_MAX:
                long_rewritten.append(ident)
            else:
                values["Lv"] = new.get("Lv")
        st.update_row(b_rows[ident][0], values)

    # 2. the replaced rows go, their long values kept back; the skipped ids;
    #    the new rows
    deferred: list[tuple[tuple[int, int], object]] = []
    for ident in gone_ids:
        rid = b_rows[ident][0]
        d = st.definition
        data = st.fetch_row(rid.page, rid.slot)
        parts = split_row(d, data)
        exact = st._exact_values(parts)  # noqa: SLF001
        for i, real, columns in st._real_indexes():  # noqa: SLF001
            key = st._key(real, columns, exact)  # noqa: SLF001
            if key is not None:
                st._btree(i, real).delete(key, rid.page, rid.slot)  # noqa: SLF001
                st._row_left_index(i, real)  # noqa: SLF001
        for column in d.columns:
            raw = parts.values.get(column.number)
            if column.is_long_value and raw:
                deferred.append((st._long_value_maps(column), decode_long_value_ref(raw)))  # noqa: SLF001
        page = DataPage(db.store.read(rid.page))
        page.remove_row(rid.slot)
        st._row_removed(rid.page, page, settle=True, retire=True)  # noqa: SLF001
        d.row_count -= 1
        db.patch_definition(d, OFFSET_ROW_COUNT, struct.pack("<I", d.row_count))
    vba_folder = next((i for i, (_r, row) in a_rows.items() if row["Name"] == "VBA" and row["Type"] == 1 and int(str(row["ParentId"])) != 1), 18)
    for skipped in range(max(b_rows) + 1, new_ids[0]):
        temp = st.insert_row({"Name": "temp", "ParentId": vba_folder, "Type": 2, **stamp(new_ids[0])})
        st.delete_row(temp)
    inserted: dict[int, object] = {}
    for ident in new_ids:
        row = a_rows[ident][1]
        values = {"Name": row["Name"], "ParentId": row["ParentId"], "Type": row["Type"], **stamp(ident)}
        if ident not in long_new and row.get("Lv") is not None:
            values["Lv"] = row["Lv"]
        inserted[ident] = st.insert_row(values)

    # 3. the new streams' values
    names = {i: str(a_rows[i][1]["Name"]) for i in new_ids}
    first = [i for i in sorted(long_new) if names[i] not in ("dir", "_VBA_PROJECT", "PROJECT")]
    first += [i for i in sorted(long_new) if names[i] == "dir"] + [i for i in sorted(long_new) if names[i] == "_VBA_PROJECT"]
    for ident in first:
        st.update_row(inserted[ident], {"Lv": a_rows[ident][1]["Lv"]})

    # 4. the rewritten streams, then PROJECT
    for ident in long_rewritten:
        st.update_row(b_rows[ident][0], {"Lv": a_rows[ident][1]["Lv"]})
    for ident in [i for i in sorted(long_new) if names[i] == "PROJECT"]:
        st.update_row(inserted[ident], {"Lv": a_rows[ident][1]["Lv"]})

    # 5. the replaced rows' long values come back
    for maps, ref in deferred:
        free_long_value(db.store, maps, ref)

    # 6. PropDataCopy under the new storage folder
    modules_folder = next(i for i, (_r, row) in a_rows.items() if row["Name"] == "Modules" and row["Type"] == 1)
    new_folder = next((i for i in new_ids if a_rows[i][1]["Type"] == 1 and int(str(a_rows[i][1]["ParentId"])) == modules_folder), None)
    if new_folder is not None:
        prop = next(i for i in new_ids if int(str(a_rows[i][1]["ParentId"])) == new_folder)
        copy_stamp = stamp(prop)
        temp = st.insert_row({"Name": "PropDataCopy", "ParentId": new_folder, "Type": 2, **copy_stamp})
        st.delete_row(temp)

    # 7. the catalog
    objects = db.table("MSysObjects")
    b_obj, a_obj = rows_by_id(db, "MSysObjects"), rows_by_id(after, "MSysObjects")
    for ident in sorted(set(a_obj) - set(b_obj)):
        row = a_obj[ident][1]
        s = serials(after, "MSysObjects", ident)
        rid = objects.insert_row({"Id": ident, "ParentId": row["ParentId"], "Name": row["Name"], "Type": row["Type"], "Flags": 0, **s})
        objects.update_row(rid, {"Owner": row["Owner"]})
    msysdb = next(i for i, (_r, row) in a_obj.items() if row["Name"] == "MSysDb")
    msysdb_stamp = serials(after, "MSysObjects", msysdb)["DateUpdate"]
    for _ in range(9):
        objects.update_row(b_obj[msysdb][0], {"LvProp": a_obj[msysdb][1]["LvProp"], "DateUpdate": msysdb_stamp})
    aces = db.table("MSysACEs")
    new_object_ids = set(a_obj) - set(b_obj)
    for r in after.table("MSysACEs").rows():
        if int(str(r["ObjectId"])) in new_object_ids:
            aces.insert_row(dict(r))
    groups = db.table("MSysNavPaneGroupCategories")
    for rid, row in list(groups.rows_with_ids()):
        groups.update_row(rid, {"Position": row["Position"]})
    nav = db.table("MSysNavPaneObjectIDs")
    nav.truncate()
    for r in after.table("MSysNavPaneObjectIDs").rows():
        nav.insert_row(dict(r))
    return db.to_bytes(), after_path.read_bytes()


def report(before_path: Path, after_path: Path) -> None:
    ours, engine = replay(before_path, after_path)
    names = {e.id: e.name for e in AccessDatabase(after_path).catalog()}
    differing = [n for n in range(1, min(len(ours), len(engine)) // PAGE) if ours[n * PAGE : (n + 1) * PAGE] != engine[n * PAGE : (n + 1) * PAGE]]
    print(f"{before_path.name} -> {after_path.name}: {len(ours) // PAGE} vs {len(engine) // PAGE} pages, {len(differing)} differ: {differing}")
    for n in differing[:12]:
        a, b = ours[n * PAGE : (n + 1) * PAGE], engine[n * PAGE : (n + 1) * PAGE]
        offs = [i for i in range(PAGE) if a[i] != b[i]]
        owner = struct.unpack_from("<I", b, 4)[0]
        label = "LVAL" if owner == 0x4C41564C else names.get(owner, owner)
        print(f"   page {n}: ours {a[0]:#04x} engine {b[0]:#04x} owner {label}: {len(offs)} bytes from {offs[0]:#x}: ours {a[offs[0]:offs[0]+12].hex()} engine {b[offs[0]:offs[0]+12].hex()}")


if __name__ == "__main__":
    report(Path(sys.argv[1]), Path(sys.argv[2]))
