"""
Excel file handler.

Supports:
  - .xlsm  (OOXML macro-enabled workbook — ZIP containing xl/vbaProject.bin)
  - .xlsb  (Binary workbook — ZIP containing xl/vbaProject.bin)
  - .xls   (Legacy BIFF8 — the entire file is a CFB)

Usage
-----
    with ExcelFile("book.xlsm") as wb:
        project = wb.vba_project()        # -> VBAProject
        modules = wb.vba_modules()        # -> dict[str, str]
        wb.set_module("Module1", src)
        wb.save("book_out.xlsm")
"""

from __future__ import annotations

import io
import zipfile
from pathlib import Path
from typing import Union

from pyopenvba.cfb import CFB
from pyopenvba.exceptions import UnsupportedFormatError, VBAProjectError
from pyopenvba.vba import VBAProject, parse_vba_project

_ZIP_FORMATS = frozenset({".xlsm", ".xlsb", ".xlam"})
_CFB_FORMATS = frozenset({".xls"})
_VBA_ENTRY = "xl/vbaProject.bin"


class ExcelFile:
    """
    Open an Excel file and provide access to its VBA project.

    Can be used as a context manager::

        with ExcelFile("book.xlsm") as wb:
            ...
    """

    def __init__(self, path: Union[str, Path]) -> None:
        self._path = Path(path)
        self._suffix = self._path.suffix.lower()
        self._zip: zipfile.ZipFile | None = None
        self._cfb: CFB | None = None
        self._zip_bytes: bytearray | None = None   # in-memory copy for write-back
        self._project: VBAProject | None = None
        self._open()

    # ------------------------------------------------------------------
    # Context manager
    # ------------------------------------------------------------------

    def __enter__(self) -> "ExcelFile":
        return self

    def __exit__(self, *_: object) -> None:
        self.close()

    def close(self) -> None:
        if self._zip is not None:
            self._zip.close()
            self._zip = None

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def vba_project(self) -> VBAProject:
        """Parse and return the VBAProject (cached after first call)."""
        if self._project is None:
            cfb = self._get_cfb()
            self._project = parse_vba_project(cfb)
        return self._project

    def vba_modules(self) -> dict[str, str]:
        """Return a mapping of module name -> source code."""
        return {m.name: m.source for m in self.vba_project().modules}

    def module_names(self) -> list[str]:
        """Return the list of VBA module names."""
        return self.vba_project().module_names()

    def get_module(self, name: str) -> str:
        """Return the source code of a named VBA module."""
        return self.vba_project().get_module(name).source

    def set_module(self, name: str, source: str) -> None:
        """
        Replace the source code of an existing VBA module in memory.

        Changes are not written to disk until :meth:`save` is called.

        .. note::
            This updates only the decompressed source text tracked by
            pyOpenVBA.  Writing back requires re-compressing and patching
            the CFB, which is not yet implemented.  See the roadmap in
            README.md.
        """
        project = self.vba_project()
        for m in project.modules:
            if m.name.casefold() == name.casefold():
                m.source = source
                return
        raise KeyError(f"Module not found: {name!r}")

    def save(self, dest: Union[str, Path, None] = None) -> None:
        """
        Save the workbook.

        ``dest`` defaults to the original file path (in-place overwrite).

        .. warning::
            Write-back (re-compressing and patching the CFB / ZIP) is not
            yet implemented.  This method raises :exc:`NotImplementedError`
            until that feature lands.
        """
        raise NotImplementedError(
            "Write-back is not yet implemented. "
            "Tracked in: https://github.com/WilliamGreenlee/pyOpenVBA/issues/1"
        )

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _open(self) -> None:
        if self._suffix in _ZIP_FORMATS:
            self._open_zip()
        elif self._suffix in _CFB_FORMATS:
            self._open_cfb_direct()
        else:
            raise UnsupportedFormatError(
                f"Unsupported file extension: {self._suffix!r}. "
                f"Supported: {sorted(_ZIP_FORMATS | _CFB_FORMATS)}"
            )

    def _open_zip(self) -> None:
        raw = self._path.read_bytes()
        self._zip_bytes = bytearray(raw)
        self._zip = zipfile.ZipFile(io.BytesIO(raw), mode="r")
        if _VBA_ENTRY not in self._zip.namelist():
            raise VBAProjectError(
                f"{self._path.name!r} contains no {_VBA_ENTRY!r}. "
                "Make sure the workbook has a VBA project (save as .xlsm in Excel)."
            )

    def _open_cfb_direct(self) -> None:
        raw = self._path.read_bytes()
        self._cfb = CFB.from_bytes(raw)

    def _get_cfb(self) -> CFB:
        if self._cfb is not None:
            return self._cfb
        if self._zip is None:
            raise RuntimeError("ExcelFile is not open.")
        vba_bin = self._zip.read(_VBA_ENTRY)
        self._cfb = CFB.from_bytes(vba_bin)
        return self._cfb
