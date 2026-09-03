"""Rename a module through the storage engine, one place at a time.

Every location is an ordinary row now, so each is an `update_row` rather
than a byte patch that cannot resize.  The steps are separable so a probe
can apply a subset and ask Access what it makes of the result.

Measured payload shapes (a two-module project, `Module1` and `Alpha`):

    \\x03DirData   <u32 0> 04 <len> <name UTF-16>  per module, then
                   04 00 00 00, where len is the name's bytes plus four
    PROJECTwm      <name MBCS> 00 <name UTF-16> 00 00  per module, then
                   00 00
    PROJECT        text: a `Module=<name>` line per module and a
                   `<name>=<window rect>` line under [Workspace]
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
for _where in (ROOT / "src", ROOT / "docs/research/access_write", ROOT / "docs/research/pcode"):
    if str(_where) not in sys.path:
        sys.path.insert(0, str(_where))

from pyopenvba.access import AccessDatabase  # noqa: E402

STORAGE = "MSysAccessStorage"
MODULE_TYPE = -32761


def _rename_catalog(db: AccessDatabase, old: str, new: str) -> None:
    objects = db.table("MSysObjects")
    for rid, row in objects.rows_with_ids():
        if row["Type"] == MODULE_TYPE and row["Name"] == old:
            objects.update_row(rid, {"Name": new})
            return
    raise LookupError(f"no module row named {old!r}")


def _rename_nav_pane(db: AccessDatabase, old: str, new: str) -> None:
    table = db.table("MSysNavPaneObjectIDs")
    for rid, row in table.rows_with_ids():
        if row["Name"] == old:
            table.update_row(rid, {"Name": new})
            return
    raise LookupError(f"no navigation-pane row named {old!r}")


def _dir_data_entry(name: str) -> bytes:
    text = name.encode("utf-16-le")
    return bytes((4, len(text) + 4)) + text


def _project_wm_entry(name: str) -> bytes:
    return name.encode("latin-1") + bytes(1) + name.encode("utf-16-le") + bytes(2)


def rename_dir_data(payload: bytes, old: str, new: str) -> bytes:
    want = _dir_data_entry(old)
    if want not in payload:
        raise LookupError("DirData holds no entry for that name")
    return payload.replace(want, _dir_data_entry(new))


def rename_project_wm(payload: bytes, old: str, new: str) -> bytes:
    want = _project_wm_entry(old)
    if want not in payload:
        raise LookupError("PROJECTwm holds no entry for that name")
    return payload.replace(want, _project_wm_entry(new))


def rename_project(text: str, old: str, new: str) -> str:
    """The `Module=` line and the `[Workspace]` line.  The stream's lines
    end CR LF, so the end-of-line anchor has to allow the CR."""
    tail = "(?=" + chr(92) + "r?$)"
    text = re.sub("(?m)^Module=" + re.escape(old) + tail, "Module=" + new, text)
    return re.sub("(?m)^" + re.escape(old) + "=", new + "=", text)


def _rename_storage(db: AccessDatabase, old: str, new: str) -> list[str]:
    table = db.table(STORAGE)
    touched: list[str] = []
    for rid, row in list(table.rows_with_ids()):
        name, payload = row["Name"], row.get("Lv")
        if not isinstance(payload, bytes):
            continue
        if name == "\x03DirData" and _dir_data_entry(old) in payload:
            table.update_row(rid, {"Lv": rename_dir_data(payload, old, new)})
            touched.append("DirData")
        elif name == "PROJECTwm" and _project_wm_entry(old) in payload:
            table.update_row(rid, {"Lv": rename_project_wm(payload, old, new)})
            touched.append("PROJECTwm")
        elif name == "PROJECT":
            text = payload.decode("latin-1")
            fixed = rename_project(text, old, new)
            if fixed != text:
                table.update_row(rid, {"Lv": fixed.encode("latin-1")})
                touched.append("PROJECT")
    return touched


def drop_srp(db: AccessDatabase) -> int:
    """Retire the compiled cache rows, which is what Access executes."""
    table = db.table(STORAGE)
    doomed = [rid for rid, row in table.rows_with_ids() if str(row["Name"]).startswith("__SRP_")]
    for rid in doomed:
        table.delete_row(rid, retire_empty=False)
    return len(doomed)


STEPS = ("catalog", "nav", "storage", "dir", "attribute", "project", "srp")


def rename(source: Path, target: Path, old: str, new: str, *, steps: set[str]) -> list[str]:
    db = AccessDatabase(source)
    done: list[str] = []
    if "catalog" in steps:
        _rename_catalog(db, old, new)
        done.append("MSysObjects")
    if "nav" in steps:
        _rename_nav_pane(db, old, new)
        done.append("MSysNavPaneObjectIDs")
    if "storage" in steps:
        done += _rename_storage(db, old, new)
    if "attribute" in steps:
        from module_stream import rename_module_stream

        done.append(rename_module_stream(db, source, old, new))
    if "dir" in steps:
        from project_streams import rename_in_project_streams

        done += rename_in_project_streams(db, old, new, {"dir"})
    if "project" in steps:
        from module_create import invalidate_cache
        from vba_project_table import rename_in_vba_project

        storage = db.table(STORAGE)
        for rid, row in list(storage.rows_with_ids()):
            if row["Name"] == "_VBA_PROJECT" and isinstance(row.get("Lv"), bytes):
                try:
                    blob, note = rename_in_vba_project(row["Lv"], old, new)
                except LookupError:
                    # A module this library created is not in the compiled
                    # cache at all, so there is nothing to rename there;
                    # mark the cache stale and VBA rebuilds it from source.
                    blob, note = invalidate_cache(row["Lv"]), "_VBA_PROJECT marked stale"
                storage.update_row(rid, {"Lv": blob})
                done.append(note)
    if "srp" in steps:
        done.append(f"dropped {drop_srp(db)} __SRP_ rows")
    db.save(target)
    return done


if __name__ == "__main__":
    src, dst, old, new = Path(sys.argv[1]), Path(sys.argv[2]), sys.argv[3], sys.argv[4]
    print(rename(src, dst, old, new, steps=set(sys.argv[5:]) or set(STEPS)))
