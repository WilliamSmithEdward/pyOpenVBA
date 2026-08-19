"""Rewrite the executable statements of an Access VBA module, in pure Python.

    python docs/research/access_write/rewrite_module.py \
        in.accdb out.accdb "acc = 7 * 6 + 100" "idx = acc - 2" ...

Each argument replaces one executable statement, in order. The statement
count must match the original, because the module's compiled p-code
carries a temporary-slot table sized by Access at compile time (see
``README.md``); adding or removing statements needs that table rebuilt and
is not supported yet. Statement *contents* are otherwise unconstrained:
different operators, constants, and control flow all work, and the p-code
is free to grow or shrink.

What gets rewritten, in order: the module's p-code and source text, then
the dir stream's MODULEOFFSET, then the MSysAccessStorage length of both
rows -- the last of which is what makes a resized row survive loading.

Dev-only research tool. Verify results by running the macro in real
Access; a database that merely *reads back* correctly proves nothing (see
the lessons document).
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from accdb_write import Perf, load_module, write_module
from vba_compile import CompileError, compile_line, name_table


def rewrite(src_db: Path, out_db: Path, statements: list[str]) -> None:
    names = name_table(src_db)
    info = load_module(src_db)
    perf = Perf(info["row"], info["modoff"])
    source = perf.source().decode("latin-1").split("\r\n")

    # Executable statements are the line-table entries our compiler can
    # rebuild; everything else (headers, declarations, blanks) is left as is.
    targets: list[int] = []
    for index in range(len(perf.recs)):
        text = source[index + 1] if index + 1 < len(source) else ""
        try:
            if compile_line(text, names) is not None:
                targets.append(index)
        except CompileError:
            continue
    if len(targets) != len(statements):
        raise SystemExit(
            f"{src_db.name} has {len(targets)} executable statements but "
            f"{len(statements)} replacements were given; counts must match")

    lines = list(perf.lines)
    recs = [bytearray(r) for r in perf.recs]
    new_source = list(source)
    for index, text in zip(targets, statements, strict=True):
        indent = len(source[index + 1]) - len(source[index + 1].lstrip())
        stmt = " " * indent + text.strip()
        lines[index] = compile_line(stmt, names)
        recs[index][3] = indent          # record byte 3 is the indent
        new_source[index + 1] = stmt

    blob = "\r\n".join(new_source).encode("latin-1")
    new_row, new_modoff = perf.build(lines=lines, recs=recs, new_source=blob)

    shutil.copy(src_db, out_db)
    data = bytearray(out_db.read_bytes())
    write_module(data, info, new_row, new_modoff)
    out_db.write_bytes(bytes(data))
    print(f"{out_db.name}: rewrote {len(statements)} statements; "
          f"module row {len(info['row'])} -> {len(new_row)} bytes")


if __name__ == "__main__":
    if len(sys.argv) < 4:
        raise SystemExit(__doc__)
    rewrite(Path(sys.argv[1]), Path(sys.argv[2]), sys.argv[3:])
