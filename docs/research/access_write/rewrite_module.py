"""Replace the body of an Access VBA procedure, in pure Python.

    python docs/research/access_write/rewrite_module.py \
        in.accdb out.accdb "acc = 0" "Do While idx < 10" ...

    python docs/research/access_write/rewrite_module.py \
        in.accdb out.accdb --file program.vba

    python docs/research/access_write/rewrite_module.py \
        in.accdb out.accdb --module ModB --file program.vba

The body of the module's first procedure -- everything between its
``FuncDefn`` and ``EndFunc`` lines -- is replaced by the statements given.
The count is free: statements may be added or removed, and the compiled
p-code, the line table, the source text and the header's procedure line
counters are all rebuilt to match. A procedure with no executable
statements at all, empty or entirely comments, is handled too.

Names the program introduces are appended to the project identifier table
automatically, so generated code is not limited to names Access already
created. ``--module`` picks the module in a project holding several, and
``--proc`` the procedure within it.

What remains needs a page allocated, and fails loudly rather than
corrupting the database: a single-row module may grow into its page's free
space, and a chained one up to its chain's capacity, but no further.

Dev-only research tool. Verify by running the macro in real Access; a
database that merely reads back correctly proves nothing.
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from accdb_write import (
    Perf,
    append_identifiers,
    find_project_row,
    load_module,
    set_lval_payload,
    write_module,
)
from vba_compile import comment_record, compile_line, is_comment, name_table, referenced_names

# Record byte 3 is the source indent; bytes 0-2 mark an executable
# statement line. Byte 6 is a frame-size hint that Access recomputes, so
# an approximation is enough -- it does not affect execution.
_EXEC_RECORD_PREFIX = b"\x00\x81\x08\x00"

# A comment line points at text rather than p-code: kind 0x09, with no
# indent (its E3 record carries that) and no frame-size hint.
_COMMENT_RECORD_PREFIX = b"\x00\x80\x09\x00"


# A procedure opens with FuncDefn and closes with EndFunc, each alone on
# its line. Their opcodes are the low 10 bits of the line's first word.
_FUNCDEFN, _ENDFUNC = 150, 105


def _first_opcode(code: bytes | None) -> int | None:
    if not code or len(code) < 2:
        return None
    return int.from_bytes(code[:2], "little") & 0x03FF


def _procedures(perf: Perf, source: list[str]) -> list[tuple[str, int, int]]:
    """``(name, first_body_line, last_body_line)`` for each procedure.

    Anchoring on FuncDefn/EndFunc rather than on "the statements we can
    recompile" means a procedure with no executable statements -- an empty
    one, or a body that is all comments -- still has a findable body.
    """
    out: list[tuple[str, int, int]] = []
    start = None
    for index, code in enumerate(perf.lines):
        opcode = _first_opcode(code)
        if opcode == _FUNCDEFN:
            start = index
        elif opcode == _ENDFUNC and start is not None:
            header = source[start] if start < len(source) else ""
            out.append((_procedure_name(header), start + 1, index - 1))
            start = None
    return out


def _procedure_name(header: str) -> str:
    """The name from a `Sub Foo(...)` / `Function Foo(...)` header line."""
    words = header.replace("(", " ").split()
    for index, word in enumerate(words):
        if word.lower() in ("sub", "function", "property") and index + 1 < len(words):
            return words[index + 1]
    return ""


def _procedure_body(perf: Perf, source: list[str],
                    procedure: str | None) -> tuple[int, int]:
    """Body bounds of the requested procedure, or of the only one."""
    found = _procedures(perf, source)
    if not found:
        raise SystemExit("no procedure (FuncDefn .. EndFunc) found in module")
    if procedure is None:
        if len(found) > 1:
            names = ", ".join(name or "?" for name, _, _ in found)
            raise SystemExit(
                f"module holds {len(found)} procedures ({names}); "
                "pass --proc <name> to choose one")
        return found[0][1], found[0][2]
    for name, first, last in found:
        if name.lower() == procedure.lower():
            return first, last
    names = ", ".join(name or "?" for name, _, _ in found)
    raise SystemExit(f"procedure {procedure!r} not found; have {names}")


def _require_reproducible(perf: Perf, info: dict) -> None:
    """Refuse to write a module we cannot rebuild byte for byte.

    Rebuilding with no changes exercises the whole layout model. If the
    result differs from what is on disk, some part of this module is laid
    out in a way the model does not cover, and writing it would corrupt
    the database rather than fail. Large comment-heavy modules are one
    known case: their line records point into a plaintext text region
    instead of the p-code, so the p-code region is not where the model
    expects it to end.

    ``verify_identity.py`` runs the same check across a whole corpus.
    """
    rebuilt, modoff = perf.build()
    if rebuilt[:perf.cafe] == info["row"][:perf.cafe] and modoff == info["modoff"]:
        return
    raise SystemExit(
        f"module {info.get('name')!r} does not rebuild byte-for-byte, so its "
        "layout is not fully modelled; refusing to write it. Run "
        "verify_identity.py on this database for detail.")


def _statement_record(text: str, code: bytes) -> bytearray:
    if is_comment(text):
        return bytearray(_COMMENT_RECORD_PREFIX + b"\x00" * 8)
    rec = bytearray(_EXEC_RECORD_PREFIX + b"\x00" * 8)
    rec[3] = len(text) - len(text.lstrip())
    rec[6:8] = (12 + 8 * max(1, len(code) // 4)).to_bytes(2, "little")
    return rec


def _encode_line(text: str, names: dict) -> bytes | None:
    """Encode one body line; a comment becomes a text record."""
    if is_comment(text):
        return comment_record(text)
    return compile_line(text, names)


def _add_missing_identifiers(out_db: Path, statements: list[str]) -> dict:
    """Append project identifiers for any name the program introduces."""
    names = name_table(out_db)
    wanted: list[str] = []
    seen = {k.lower() for k in names}
    for text in statements:
        for token in referenced_names(text):
            if token.lower() not in seen:
                seen.add(token.lower())
                wanted.append(token)
    if not wanted:
        return names
    page, slot, row = find_project_row(out_db)
    data = bytearray(out_db.read_bytes())
    new_row = append_identifiers(row, wanted)
    set_lval_payload(data, page, slot, new_row, len(row))
    out_db.write_bytes(bytes(data))
    print(f"added {len(wanted)} identifier(s): {', '.join(wanted)}")
    return name_table(out_db)


def rewrite(src_db: Path, out_db: Path, statements: list[str],
            module: str | None = None, procedure: str | None = None) -> None:
    shutil.copy(src_db, out_db)
    info = load_module(out_db, module)
    perf = Perf(info["row"], info["modoff"])
    attributes, source = perf.attribute_lines(), perf.source_lines()
    # Check the layout is one we model before interpreting anything in
    # it, then resolve the target -- both before touching the file, so a
    # refusal leaves no appended identifiers behind.
    _require_reproducible(perf, info)
    first, last = _procedure_body(perf, source, procedure)
    names = _add_missing_identifiers(out_db, statements)

    body, body_recs, body_src = [], [], []
    for text in statements:
        stmt = text.rstrip()
        code = _encode_line(stmt, names)
        if code is None:
            raise SystemExit(f"not a statement or comment: {stmt!r}")
        body.append(code)
        body_recs.append(_statement_record(stmt, code))
        body_src.append(stmt)

    lines = perf.lines[:first] + body + perf.lines[last + 1:]
    recs = ([bytearray(r) for r in perf.recs[:first]] + body_recs
            + [bytearray(r) for r in perf.recs[last + 1:]])
    new_source = (attributes + source[:first] + body_src
                  + source[last + 1:])
    blob = "\r\n".join(new_source).encode("latin-1")

    new_row, new_modoff = perf.build(lines=lines, recs=recs, new_source=blob)
    data = bytearray(out_db.read_bytes())
    write_module(data, info, new_row, new_modoff)
    out_db.write_bytes(bytes(data))
    print(f"{out_db.name}: body {last - first + 1} -> {len(statements)} "
          f"lines, module {perf.num_lines} -> {len(recs)} lines, "
          f"row {len(info['row'])} -> {len(new_row)} bytes")


if __name__ == "__main__":
    if len(sys.argv) < 4:
        raise SystemExit(__doc__)
    argv = sys.argv[1:]
    module = None
    procedure = None
    for flag in ("--module", "--proc"):
        if flag in argv:
            i = argv.index(flag)
            value = argv[i + 1]
            del argv[i:i + 2]
            if flag == "--module":
                module = value
            else:
                procedure = value
    stmts = (
        [ln for ln in Path(argv[3]).read_text().splitlines() if ln.strip()]
        if argv[2] == "--file" else argv[2:]
    )
    rewrite(Path(argv[0]), Path(argv[1]), stmts, module, procedure)
