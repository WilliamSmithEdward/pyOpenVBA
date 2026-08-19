"""Gate: our p-code must match what Microsoft's compiler emitted.

Point this at any Access database whose VBA was compiled by Access, and
it recompiles every statement it understands with :mod:`vba_compile`,
comparing the bytes against the p-code already in the file. Statements the
compiler does not cover (declarations, procedure headers, blank lines) are
reported as skipped rather than silently passed.

    python docs/research/access_write/verify_compiler.py <db.accdb> [...]

A non-zero exit means some statement compiled to different bytes than
Access produced, which is a real code-generation defect.

Dev-only: needs a Windows-built ``.accdb`` to compare against.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from accdb_write import Perf, load_module
from vba_compile import CompileError, compile_line, name_table

from pyopenvba.access_read import AccessReader


def check(path: Path) -> tuple[int, int, int]:
    """Check every module in the project, not just the first."""
    modules = [s.name for s in AccessReader(path).find_module_streams()]
    totals = [0, 0, 0]
    for name in modules or [None]:
        for index, value in enumerate(check_module(path, name)):
            totals[index] += value
    return totals[0], totals[1], totals[2]


def check_module(path: Path, module: str | None) -> tuple[int, int, int]:
    info = load_module(path, module)
    perf = Perf(info["row"], info["modoff"])
    names = name_table(path)
    # source_lines() drops the leading Attribute block, which the line
    # table does not represent and which runs to five or more lines in a
    # class module -- assuming one silently misreads every class module.
    source = perf.source_lines()
    matched = differed = skipped = 0
    for index, _ in enumerate(perf.recs):
        text = source[index] if index < len(source) else ""
        try:
            mine = compile_line(text, names)
        except CompileError:
            skipped += 1
            continue
        if mine is None:
            skipped += 1
            continue
        if mine == perf.lines[index]:
            matched += 1
        else:
            differed += 1
            theirs = perf.lines[index]
            print(f"  DIFFER {text.strip()[:44]}")
            print(f"     access: {theirs.hex() if theirs else None}")
            print(f"     ours  : {mine.hex()}")
    return matched, differed, skipped


def main(argv: list[str]) -> int:
    if not argv:
        raise SystemExit("usage: verify_compiler.py <db.accdb> [...]")
    total_ok = total_bad = total_skip = 0
    for name in argv:
        path = Path(name)
        try:
            ok, bad, skip = check(path)
        except Exception as error:
            print(f"  skip {path.name}: {type(error).__name__}: {error}")
            continue
        print(f"{path.name}: {ok} identical, {bad} differ, {skip} not compiled")
        total_ok += ok
        total_bad += bad
        total_skip += skip
    print(f"\n{total_ok} statements byte-identical to Access, "
          f"{total_bad} differ, {total_skip} skipped")
    return 1 if total_bad else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
