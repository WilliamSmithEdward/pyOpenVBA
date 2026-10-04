"""How long save() takes as the workbook around the project grows.

A workbook is built with a small VBA project and worksheet parts of a
chosen size, one module is edited, and the save is timed. The project is
the same in every row, so what changes from row to row is only the size
of the parts the edit does not touch.

    python scripts/benchmark_save.py

Compares both public save modes in one process, alternating measurement
order and reporting medians of five samples. Package construction and
verification are outside the clock. These are synthetic package save
measurements, not Office startup timings. Outputs use a temporary directory.
"""

from __future__ import annotations

import io
import random
import sys
import tempfile
import time
import zipfile
from pathlib import Path
from statistics import median

import pyopenvba
from pyopenvba import ExcelFile

#: Rows in each of the three worksheet parts added to the workbook.
SIZES = (0, 2_000, 20_000, 100_000)
REPEATS = 5


def sheet_xml(rows: int, seed: int) -> bytes:
    """A worksheet part of ``rows`` rows of twenty numbers, as Excel spells one."""
    rng = random.Random(seed)
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>']
    for row in range(1, rows + 1):
        cells = "".join(f'<c r="{column}{row}" s="{rng.randint(0, 9)}"><v>{rng.random() * 1e6:.4f}</v></c>'
                        for column in "ABCDEFGHIJKLMNOPQRST")
        out.append(f'<row r="{row}" spans="1:20">{cells}</row>')
    out.append("</sheetData></worksheet>")
    return "".join(out).encode()


def build(path: Path, rows: int) -> None:
    """A new workbook with one module, and three worksheet parts of ``rows`` rows beside it."""
    with ExcelFile.create_new(path) as book:
        book.vba_project().add_module("Bench", "Public Sub Run()\r\n    Debug.Print 1\r\nEnd Sub\r\n")
        book.save()
    buffer = io.BytesIO()
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as out:
        for info in source.infolist():
            out.writestr(info, source.read(info.filename))
        for index in range(3 if rows else 0):
            out.writestr(f"xl/worksheets/data{index}.xml", sheet_xml(rows, index))
    path.write_bytes(buffer.getvalue())


def main() -> int:
    print(f"pyOpenVBA {pyopenvba.__version__}, Python {sys.version.split()[0]}, {sys.platform}")
    print(f"{'rows/sheet':>10}  {'file MB':>8}  {'parts MB':>8}  {'copy ms':>9}  {'rebuild ms':>10}  {'speedup':>8}")
    with tempfile.TemporaryDirectory() as folder:
        for rows in SIZES:
            source = Path(folder) / f"book{rows}.xlsm"
            build(source, rows)
            with zipfile.ZipFile(source) as package:
                unpacked = sum(info.file_size for info in package.infolist())
            samples: dict[bool, list[float]] = {False: [], True: []}
            for repeat in range(REPEATS):
                for full_rebuild in ((False, True) if repeat % 2 == 0 else (True, False)):
                    target = Path(folder) / f"out{full_rebuild}.xlsm"
                    with ExcelFile(source) as book:
                        edited = book.get_module("Bench") + "' edited\r\n"
                        book.set_module("Bench", edited)
                        start = time.perf_counter()
                        book.save(target, full_rebuild=full_rebuild)
                        samples[full_rebuild].append(time.perf_counter() - start)
                    with ExcelFile(target) as saved:
                        assert saved.get_module("Bench") == edited
            copied, rebuilt = median(samples[False]), median(samples[True])
            print(f"{rows:>10}  {source.stat().st_size / 1e6:>8.1f}  {unpacked / 1e6:>8.1f}"
                  f"  {copied * 1e3:>9.1f}  {rebuilt * 1e3:>10.1f}  {rebuilt / copied:>7.1f}x", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
