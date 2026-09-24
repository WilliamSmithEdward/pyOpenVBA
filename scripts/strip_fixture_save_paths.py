"""Take the folder Excel saved it in out of every committed fixture workbook.

    python scripts/strip_fixture_save_paths.py [--check]

rewrites each workbook under tests/fixtures that git tracks and whose
xl/workbook.xml records a save path, with fixture_workbook.without_save_path,
and checks every rewrite before writing it: the package's layout is one the
library writes back byte for byte, every other part is identical, header
and compressed body alike, in the same order, and xl/workbook.xml is the
same text less the one block. A workbook that fails a check is left as it
is and named. --check only names the workbooks that record a save path.
"""

from __future__ import annotations

import io
import subprocess
import sys
import zipfile
from dataclasses import astuple
from pathlib import Path

from fixture_workbook import SAVE_PATH, WORKBOOK_PART, without_save_path
from pyopenvba.powerquery._opc import Entry, OpcFile

ROOT = Path(__file__).resolve().parent.parent


def _tracked() -> list[Path]:
    listing = subprocess.run(["git", "ls-files", "tests/fixtures"], cwd=ROOT, capture_output=True, text=True,
                             check=True).stdout
    return [ROOT / name for name in listing.splitlines() if name.lower().endswith((".xlsx", ".xlsm"))]


def _records(path: Path) -> bool:
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        return WORKBOOK_PART in names and SAVE_PATH.search(package.read(WORKBOOK_PART).decode("utf-8")) is not None


def _header(entry: Entry) -> tuple[object, ...]:
    """Everything about an entry but its body and what follows from it."""
    fields = astuple(entry)
    return fields[:1] + fields[2:6] + fields[8:]


def _problem(before: bytes, after: bytes) -> str | None:
    """Why ``after`` is not ``before`` less its save path, or None."""
    rebuilt = OpcFile.parse(before)
    rebuilt.source = None
    if rebuilt.serialize() != before:
        return "its layout is not one the library writes back byte for byte"
    old, new = OpcFile.parse(before), OpcFile.parse(after)
    if old.names() != new.names():
        return "the parts changed"
    for was, now in zip(old.entries, new.entries, strict=True):
        if was.name != WORKBOOK_PART:
            if was != now:
                return f"{was.name} changed"
            continue
        if _header(was) != _header(now):
            return f"{WORKBOOK_PART}'s header changed"
        text = was.read().decode("utf-8")
        if len(SAVE_PATH.findall(text)) != 1 or SAVE_PATH.sub("", text) != now.read().decode("utf-8"):
            return f"{WORKBOOK_PART} is not the same text less one block"
    with zipfile.ZipFile(io.BytesIO(after)) as package:
        if package.testzip() is not None:
            return "a part fails its CRC"
    return None


def main() -> None:
    check = "--check" in sys.argv[1:]
    books = _tracked()
    carrying = [path for path in books if _records(path)]
    print(f"{len(books)} workbooks, {len(carrying)} record a save path")
    left: list[tuple[Path, str]] = []
    for path in carrying:
        if check:
            print("  ", path.relative_to(ROOT).as_posix())
            continue
        before = path.read_bytes()
        after = without_save_path(before)
        problem = _problem(before, after)
        if problem is not None:
            left.append((path, problem))
            continue
        path.write_bytes(after)
    if not check:
        print(f"rewrote {len(carrying) - len(left)}")
    for path, problem in left:
        print(f"left alone: {path.relative_to(ROOT).as_posix()} -- {problem}")


if __name__ == "__main__":
    main()
