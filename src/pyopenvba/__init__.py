"""
pyOpenVBA
=========
Read and write VBA sections of xlsm / xlsb / xls Excel files
using pure Python — no external dependencies.

Public API
----------
    from pyopenvba import ExcelFile

    with ExcelFile("workbook.xlsm") as wb:
        modules = wb.vba_modules()          # dict[name -> source]
        wb.set_module("Module1", new_src)
        wb.save("workbook_modified.xlsm")
"""

from pyopenvba.excel import ExcelFile
from pyopenvba.exceptions import (
    PyOpenVBAError,
    CFBError,
    VBAProjectError,
    UnsupportedFormatError,
)
from pyopenvba.vba import VBAModuleKind

__all__ = [
    "ExcelFile",
    "VBAModuleKind",
    "PyOpenVBAError",
    "CFBError",
    "VBAProjectError",
    "UnsupportedFormatError",
]

__version__ = "0.1.0"
