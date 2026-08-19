"""Replace the body of an Access VBA procedure, in pure Python.

    python docs/research/access_write/rewrite_module.py \
        in.accdb out.accdb "acc = 0" "Do While idx < 10" ...

    python docs/research/access_write/rewrite_module.py \
        in.accdb out.accdb --file program.vba

Every executable statement in the template's procedure is replaced by the
statements given. The count is free: statements may be added or removed,
and the compiled p-code, the line table, the source text and the header's
procedure line counters are all rebuilt to match.

Two constraints remain, and both fail loudly rather than corrupting the
database:

* **Names must already exist in the project.** Generated code can only
  use identifiers Access has already placed in the project table, since
  adding one means rebuilding the `_VBA_PROJECT` symbol buckets. An
  unknown name raises ``CompileError``.
* **The module row must still fit its 4 KB page.** Growing past the free
  space raises ``ValueError``; spilling onto a fresh page needs the LVAL
  chain allocator, which is not implemented.

Dev-only research tool. Verify by running the macro in real Access; a
database that merely reads back correctly proves nothing.
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from accdb_write import Perf, load_module, write_module
from vba_compile import CompileError, compile_line, name_table

# Record byte 3 is the source indent; bytes 0-2 mark an executable
# statement line. Byte 6 is a frame-size hint that Access recomputes, so
# an approximation is enough -- it does not affect execution.
_EXEC_RECORD_PREFIX = b"\x00\x81\x08\x00"


def _statement_record(text: str, code: bytes) -> bytearray:
    rec = bytearray(_EXEC_RECORD_PREFIX + b"\x00" * 8)
    rec[3] = len(text) - len(text.lstrip())
    rec[6:8] = (12 + 8 * max(1, len(code) // 4)).to_bytes(2, "little")
    return rec


def rewrite(src_db: Path, out_db: Path, statements: list[str]) -> None:
    names = name_table(src_db)
    info = load_module(src_db)
    perf = Perf(info["row"], info["modoff"])
    source = perf.source().decode("latin-1").split("\r\n")

    # Executable statements are the line-table entries our compiler can
    # rebuild; the surrounding header, declarations and procedure lines
    # are carried over untouched.
    targets = []
    for index in range(len(perf.recs)):
        text = source[index + 1] if index + 1 < len(source) else ""
        try:
            if compile_line(text, names) is not None:
                targets.append(index)
        except CompileError:
            continue
    if not targets:
        raise SystemExit(f"{src_db.name}: no executable statements found")
    first, last = targets[0], targets[-1]

    body, body_recs, body_src = [], [], []
    for text in statements:
        stmt = text.rstrip()
        code = compile_line(stmt, names)
        if code is None:
            raise SystemExit(f"not an executable statement: {stmt!r}")
        body.append(code)
        body_recs.append(_statement_record(stmt, code))
        body_src.append(stmt)

    lines = perf.lines[:first] + body + perf.lines[last + 1:]
    recs = ([bytearray(r) for r in perf.recs[:first]] + body_recs
            + [bytearray(r) for r in perf.recs[last + 1:]])
    new_source = source[:first + 1] + body_src + source[last + 2:]
    blob = "\r\n".join(new_source).encode("latin-1")

    new_row, new_modoff = perf.build(lines=lines, recs=recs, new_source=blob)
    shutil.copy(src_db, out_db)
    data = bytearray(out_db.read_bytes())
    write_module(data, info, new_row, new_modoff)
    out_db.write_bytes(bytes(data))
    print(f"{out_db.name}: {len(targets)} -> {len(statements)} statements, "
          f"{perf.num_lines} -> {len(recs)} lines, "
          f"module row {len(info['row'])} -> {len(new_row)} bytes")


if __name__ == "__main__":
    if len(sys.argv) < 4:
        raise SystemExit(__doc__)
    if sys.argv[3] == "--file":
        text = Path(sys.argv[4]).read_text().splitlines()
        stmts = [ln for ln in text if ln.strip()]
    else:
        stmts = sys.argv[3:]
    rewrite(Path(sys.argv[1]), Path(sys.argv[2]), stmts)
