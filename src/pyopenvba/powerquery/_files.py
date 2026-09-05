"""Queries on disk, so they can live in version control.

One ``.m`` file per query, holding the query's M and nothing else, beside
a small manifest that records what a file name cannot: the query's real
name, its description, and the group it sits in.  The manifest is what
makes the round trip exact, because a query may be called ``Sales/EU``
or ``Café`` and a file may not.
"""

from __future__ import annotations

import json
import re
from pathlib import Path

from pyopenvba.exceptions import PowerQueryError
from pyopenvba.powerquery.workbook import PowerQueryWorkbook

#: The manifest that sits beside the ``.m`` files.
MANIFEST = "queries.json"
_UNSAFE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_RESERVED = frozenset(
    "CON PRN AUX NUL COM1 COM2 COM3 COM4 COM5 COM6 COM7 COM8 COM9 "
    "LPT1 LPT2 LPT3 LPT4 LPT5 LPT6 LPT7 LPT8 LPT9".split()
)


def file_name(name: str, taken: set[str]) -> str:
    """A file name for a query name, unique among the ones already used."""
    safe = _UNSAFE.sub("_", name).rstrip(" .") or "query"
    if safe.upper() in _RESERVED:
        safe = f"_{safe}"
    stem, index = safe, 2
    while safe.lower() in taken:
        safe = f"{stem}_{index}"
        index += 1
    taken.add(safe.lower())
    return safe + ".m"


def pull_queries(
    workbook: str | Path,
    dest_dir: str | Path,
    *,
    encoding: str = "utf-8",
    overwrite: bool = True,
) -> list[Path]:
    """Write every query of `workbook` into `dest_dir` as a ``.m`` file."""
    book = PowerQueryWorkbook(workbook)
    out = Path(dest_dir)
    out.mkdir(parents=True, exist_ok=True)
    taken: set[str] = set()
    written: list[Path] = []
    manifest: list[dict[str, object]] = []
    for query in book.queries():
        target = out / file_name(query.name, taken)
        if target.exists() and not overwrite:
            raise PowerQueryError(f"{target} is already there; pass overwrite=True to replace it")
        target.write_text(query.formula, encoding=encoding, newline="")
        written.append(target)
        group = query.group
        manifest.append(
            {
                "name": query.name,
                "file": target.name,
                "description": query.description,
                "group": group.name if group else None,
                "loadTarget": query.load_target,
            }
        )
    index = out / MANIFEST
    index.write_text(json.dumps({"queries": manifest}, indent=2, ensure_ascii=False), encoding=encoding)
    written.append(index)
    return written


def push_queries(
    src_dir: str | Path,
    workbook: str | Path,
    *,
    out: str | Path | None = None,
    encoding: str = "utf-8",
    remove_missing: bool = False,
) -> list[str]:
    """Read ``.m`` files from `src_dir` back into `workbook` and save.

    A file whose query is already there updates it; one whose query is
    not adds it.  Queries the directory does not mention are left alone
    unless `remove_missing` says otherwise.

    Line endings are read as they sit on disk, so a query pulled and
    pushed again is the one that was there.
    """
    source = Path(src_dir)
    if not source.is_dir():
        raise PowerQueryError(f"{source} is not a directory")
    book = PowerQueryWorkbook(workbook)
    manifest: dict[str, dict[str, object]] = {}
    index = source / MANIFEST
    if index.is_file():
        try:
            loaded = json.loads(index.read_text(encoding=encoding))
        except json.JSONDecodeError as exc:
            raise PowerQueryError(f"{index} does not parse: {exc}") from exc
        for record in loaded.get("queries", []):
            manifest[str(record["file"])] = record
    touched: list[str] = []
    seen: set[str] = set()
    for path in sorted(source.glob("*.m")):
        record = manifest.get(path.name, {})
        name = str(record.get("name") or path.stem)
        with path.open("r", encoding=encoding, newline="") as handle:
            formula = handle.read()
        seen.add(name)
        description = record.get("description")
        if name in book.query_names():
            query = book.query(name)
            if query.formula != formula:
                query.formula = formula
            if description is not None and query.description != description:
                query.description = str(description)
        else:
            book.add_query(name, formula, description=None if description is None else str(description))
        touched.append(name)
    if remove_missing:
        for name in list(book.query_names()):
            if name not in seen:
                book.remove_query(name)
                touched.append(name)
    book.save(out)
    return touched
