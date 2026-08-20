"""Gate: rewriting a module with its own content must change nothing.

Every storage rule in :mod:`accdb_write` was found by this check failing.
Rebuilding an unmodified module exercises the whole write path -- the
perfcache layout, the per-procedure line counters, the LVAL row or chain,
the long-value descriptor -- and the only correct result is a file
identical to the one we started with, byte for byte across the whole row
-- not merely the header. A rule that is merely close enough
to load shows up here as a diff.

    python docs/research/access_write/verify_identity.py <db.accdb> [...]

Two levels are checked per module:

* **header** -- ``Perf.build()`` with no changes reproduces the module
  stream's pre-``0xCAFE`` bytes;
* **file** -- writing that stream back through ``set_lval_payload``
  reproduces the database byte for byte.

Non-zero exit on any difference. Dev-only research tool.

A warning about fixtures, learned the hard way: build the VBA source you
feed Access with ``open(path, "w", newline="")``. Python's text mode
translates ``
`` on Windows, so a source written with ``

`` lands on
disk as ``


``. Access accepts it and produces a module with a blank
line between every real one -- roughly twice the line records you meant --
which looks exactly like a decoder bug in whatever you test next.
"""
from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from accdb_write import Perf, load_module, set_lval_payload

from pyopenvba.access_read import AccessReader


def check(path: Path) -> tuple[int, int]:
    """Return ``(modules_checked, failures)`` for one database."""
    try:
        modules = [s.name for s in AccessReader(path).find_module_streams()]
    except Exception as error:
        print(f"  skip {path.name}: {type(error).__name__}: {error}")
        return 0, 0
    if not modules:
        return 0, 0

    checked = failures = 0
    original = path.read_bytes()
    with tempfile.TemporaryDirectory() as scratch:
        work = Path(scratch) / path.name
        shutil.copy(path, work)
        for name in modules:
            checked += 1
            info = load_module(work, name)
            perf = Perf(info["row"], info["modoff"])

            rebuilt, modoff = perf.build()
            original_row = bytes(info["row"])
            if rebuilt != original_row:
                differing = [
                    i for i in range(min(len(rebuilt), len(original_row)))
                    if rebuilt[i] != original_row[i]
                ]
                failures += 1
                where = ("header" if differing and differing[0] < perf.cafe
                         else "p-code" if differing
                         else "length")
                print(f"  ROW     {path.name}:{name} differs in {where} "
                      f"at {differing[:8]}"
                      f"{' ...' if len(differing) > 8 else ''}"
                      f" (len {len(rebuilt)} vs {len(original_row)})")
                continue
            if modoff != info["modoff"]:
                failures += 1
                print(f"  MODOFF  {path.name}:{name} "
                      f"{info['modoff']} -> {modoff}")
                continue

            data = bytearray(work.read_bytes())
            try:
                # Write back what `build` produced, not what was read.
                # Writing the original row made this check trivially true
                # and tested only `set_lval_payload`.
                set_lval_payload(data, info["page"], info["slot"],
                                 rebuilt, len(info["row"]))
            except Exception as error:
                failures += 1
                print(f"  WRITE   {path.name}:{name}: "
                      f"{type(error).__name__}: {error}")
                continue
            if bytes(data) != original:
                differing = sum(1 for a, b in zip(data, original, strict=False)
                                if a != b)
                failures += 1
                print(f"  FILE    {path.name}:{name} differs in "
                      f"{differing} bytes after a no-op rewrite")
    return checked, failures


def main(argv: list[str]) -> int:
    if not argv:
        raise SystemExit("usage: verify_identity.py <db.accdb> [...]")
    total = bad = 0
    for name in argv:
        checked, failures = check(Path(name))
        total += checked
        bad += failures
        if checked and not failures:
            print(f"  ok   {Path(name).name}: {checked} module(s) identical")
    print(f"\n{total - bad}/{total} modules rebuild identically, {bad} differ")
    return 1 if bad else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
