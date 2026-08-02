"""
Excel file handler.

Supports:
  - .xlsm  (OOXML macro-enabled workbook — ZIP containing xl/vbaProject.bin)
  - .xlsb  (Binary workbook — ZIP containing xl/vbaProject.bin)
  - .xlam  (Macro-enabled add-in — same ZIP container as .xlsm)
  - .xls   (Legacy BIFF8 — the entire file is a CFB)

Usage
-----
    with ExcelFile("book.xlsm") as wb:
        project = wb.vba_project()        # -> VBAProject
        modules = wb.vba_modules()        # -> dict[str, str]
        wb.set_module("Module1", src)
        wb.save("book_out.xlsm")

All shared behavior (read, edit, pull/push, safety-gated save) lives in
:class:`pyopenvba._host.VBAHostFile`.
"""

from __future__ import annotations

from pathlib import Path

from pyopenvba._host import VBAHostFile

_ZIP_FORMATS = frozenset({".xlsm", ".xlsb", ".xlam"})
_CFB_FORMATS = frozenset({".xls"})
_VBA_ENTRY = "xl/vbaProject.bin"


class ExcelFile(VBAHostFile):
    """
    Open an Excel file and provide access to its VBA project.

    Can be used as a context manager::

        with ExcelFile("book.xlsm") as wb:
            ...
    """

    _zip_formats = _ZIP_FORMATS
    _cfb_formats = _CFB_FORMATS
    _vba_entry = _VBA_ENTRY
    _host_noun = "workbook"
    _no_vba_hint = (
        "Make sure the workbook has a VBA project (save as .xlsm in Excel)."
    )

    @classmethod
    def create_new(cls, path: str | Path) -> ExcelFile:
        """
        Create a new macro-enabled workbook at ``path`` containing an empty
        VBA project (``ThisWorkbook``, ``Sheet1``, and a bare ``Module1``)
        and return an open :class:`ExcelFile` for it.

        Supported extensions: ``.xlsm`` (default), ``.xlsb``, and
        ``.xlam`` (Excel add-in).

        The bytes are decoded from a baked-in template captured from a
        freshly Excel-authored file, so the resulting file opens
        cleanly in Excel without any "found a problem" repair prompt.

        ``path`` is overwritten if it already exists.
        """
        target = Path(path)
        suffix = target.suffix.lower()
        if suffix == ".xlsb":
            from pyopenvba._templates import EMPTY_XLSB_BYTES
            template = EMPTY_XLSB_BYTES
        elif suffix == ".xlam":
            from pyopenvba._templates import EMPTY_XLAM_BYTES
            template = EMPTY_XLAM_BYTES
        else:
            from pyopenvba._templates import EMPTY_XLSM_BYTES
            template = EMPTY_XLSM_BYTES
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(template)
        return cls(target)
