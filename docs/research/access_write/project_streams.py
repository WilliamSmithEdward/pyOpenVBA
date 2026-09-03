"""The two places inside the compiled VBA project that carry a module's
name: the `dir` stream's MODULENAME / MODULENAMEUNICODE records, and the
module's own `Attribute VB_Name` line.  Both rows are MS-OVBA compressed,
so each is decompressed, edited and recompressed, then written back
through the engine's row writer.
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
for _where in (ROOT / "src", ROOT / "docs/research/access_write", ROOT / "docs/research/pcode"):
    if str(_where) not in sys.path:
        sys.path.insert(0, str(_where))

from pyopenvba.access import AccessDatabase  # noqa: E402
from pyopenvba.vba import compress, decompress  # noqa: E402

MODULENAME = bytes.fromhex("1900")  # [MS-OVBA] 2.3.4.2.3.2.1
MODULENAMEUNICODE = bytes.fromhex("4700")  # [MS-OVBA] 2.3.4.2.3.2.2
STORAGE = "MSysAccessStorage"


def rename_in_dir(stream: bytes, old: str, new: str, code_page: str = "latin-1") -> bytes:
    """Rewrite one module's two name records in a decompressed dir stream."""
    out = bytearray(stream)
    for record, encoding in ((MODULENAME, code_page), (MODULENAMEUNICODE, "utf-16-le")):
        want, text = old.encode(encoding), new.encode(encoding)
        header = record + len(want).to_bytes(4, "little") + want
        at = out.find(header)
        if at < 0:
            raise LookupError(f"the dir stream has no {record.hex()} record for {old!r}")
        out[at : at + len(header)] = record + len(text).to_bytes(4, "little") + text
    return bytes(out)


def rename_attribute(source: bytes, old: str, new: str) -> bytes:
    """The module's own `Attribute VB_Name = "..."` line."""
    want = f'Attribute VB_Name = "{old}"'.encode("latin-1")
    if want not in source:
        raise LookupError(f"the module holds no VB_Name attribute for {old!r}")
    return source.replace(want, f'Attribute VB_Name = "{new}"'.encode("latin-1"))


def rename_in_project_streams(db: AccessDatabase, old: str, new: str, steps: set[str]) -> list[str]:
    table = db.table(STORAGE)
    touched: list[str] = []
    for rid, row in list(table.rows_with_ids()):
        name, payload = row["Name"], row.get("Lv")
        if not isinstance(payload, bytes) or not payload:
            continue
        if name == "dir" and "dir" in steps:
            stream = decompress(payload)
            table.update_row(rid, {"Lv": compress(rename_in_dir(stream, old, new))})
            touched.append("dir")
        elif name == old and "attribute" in steps:
            source = decompress(payload)
            table.update_row(rid, {"Lv": compress(rename_attribute(source, old, new)), "Name": new})
            touched.append(f"module row {old} -> {new}")
    return touched
