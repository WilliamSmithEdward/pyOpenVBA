"""
pyOpenVBA
=========
Read and write VBA sections of xlsm / xlsb / xls Excel files
using pure Python — no external dependencies.

Public API
----------
    from pyopenvba import ExcelFile, pull, push

    # In-process module edit
    with ExcelFile("workbook.xlsm") as wb:
        modules = wb.vba_modules()          # dict[name -> source]
        wb.set_module("Module1", new_src)
        wb.save("workbook_modified.xlsm")

    # Disk-based workflow (.bas / .cls files)
    pull("workbook.xlsm", "./vba_src")      # extract modules
    push("./vba_src", "workbook.xlsm")      # write edits back in place
"""

from pathlib import Path
from typing import Union

from pyopenvba.excel import ExcelFile
from pyopenvba.exceptions import (
    PyOpenVBAError,
    CFBError,
    VBAProjectError,
    UnsupportedFormatError,
)
from pyopenvba.vba import VBAModuleKind


def pull(
    workbook: Union[str, Path],
    dest_dir: Union[str, Path],
    *,
    encoding: str = "utf-8",
    overwrite: bool = True,
) -> list[Path]:
    """
    Export every VBA module from ``workbook`` into ``dest_dir`` as
    ``.bas`` / ``.cls`` files. Returns the paths written.
    """
    with ExcelFile(workbook) as wb:
        return wb.pull_modules(dest_dir, encoding=encoding, overwrite=overwrite)


def push(
    src_dir: Union[str, Path],
    workbook: Union[str, Path],
    *,
    out: Union[str, Path, None] = None,
    encoding: str = "utf-8",
    strict: bool = False,
) -> list[str]:
    """
    Update VBA modules in ``workbook`` from ``.bas`` / ``.cls`` files
    in ``src_dir`` and save. Saves in place unless ``out`` is given.
    Returns the list of updated module names.
    """
    with ExcelFile(workbook) as wb:
        updated = wb.push_modules(src_dir, encoding=encoding, strict=strict)
        wb.save(out)
    return updated


__all__ = [
    "ExcelFile",
    "VBAModuleKind",
    "PyOpenVBAError",
    "CFBError",
    "VBAProjectError",
    "UnsupportedFormatError",
    "pull",
    "push",
]

__version__ = "1.0.0"
