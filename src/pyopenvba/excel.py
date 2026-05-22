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
from pyopenvba.vba import VBAProject, parse_vba_project, write_back_modules
from pyopenvba.vba import VBAModuleKind

_ZIP_FORMATS = frozenset({".xlsm", ".xlsb", ".xlam"})
_CFB_FORMATS = frozenset({".xls"})
_VBA_ENTRY = "xl/vbaProject.bin"

# File extensions used by the VBE export/import workflow.
# - Standard procedural modules -> .bas
# - Everything else (class, document/sheet/workbook, designer/form) -> .cls
# We do not write .frm/.frx layout bytes; UserForm layout is preserved
# inside the CFB and never round-trips through disk.
_BAS_EXT = ".bas"
_CLS_EXT = ".cls"
_SOURCE_EXTS = frozenset({_BAS_EXT, _CLS_EXT})


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
        self._zip_bytes: bytes | None = None
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

    def vba_project_bytes(self) -> bytes:
        """Return the raw bytes of ``xl/vbaProject.bin`` (or the whole CFB
        file for legacy ``.xls``)."""
        if self._suffix in _CFB_FORMATS:
            return self._path.read_bytes()
        assert self._zip is not None
        return self._zip.read(_VBA_ENTRY)

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
        """
        project = self.vba_project()
        for m in project.modules:
            if m.name.casefold() == name.casefold():
                m.source = source
                m.dirty = True
                return
        raise KeyError(f"Module not found: {name!r}")

    def validate(self) -> list[str]:
        """Return cross-structure inconsistency messages; empty list means OK."""
        return self.vba_project().validate(self._get_cfb())

    # ------------------------------------------------------------------
    # Push / pull (disk-based module sync)
    # ------------------------------------------------------------------

    def pull_modules(
        self,
        dest_dir: Union[str, Path],
        *,
        encoding: str = "utf-8",
        overwrite: bool = True,
    ) -> list[Path]:
        """
        Export every VBA module's source to a file in ``dest_dir``.

        Standard procedural modules are written as ``<name>.bas``; class,
        document, and designer modules are written as ``<name>.cls``.
        Source bytes use CRLF line endings to match VBE's own export
        format. UserForm layout (``.frx``) is **not** exported — it is
        preserved verbatim inside the workbook on save.

        Returns the list of file paths written.
        """
        out_dir = Path(dest_dir)
        out_dir.mkdir(parents=True, exist_ok=True)

        written: list[Path] = []
        for m in self.vba_project().modules:
            ext = _BAS_EXT if m.kind == VBAModuleKind.standard else _CLS_EXT
            target = out_dir / f"{m.name}{ext}"
            if target.exists() and not overwrite:
                raise FileExistsError(
                    f"Refusing to overwrite {target} (overwrite=False)."
                )
            text = m.source.replace("\r\n", "\n").replace("\r", "\n")
            data = text.replace("\n", "\r\n").encode(encoding, errors="replace")
            target.write_bytes(data)
            written.append(target)
        return written

    def push_modules(
        self,
        src_dir: Union[str, Path],
        *,
        encoding: str = "utf-8",
        strict: bool = False,
    ) -> list[str]:
        """
        Update module source from files in ``src_dir``.

        Each ``<name>.bas`` or ``<name>.cls`` file is matched (case-
        insensitively) to a module of the same logical name and its
        source replaces the module's current source. Line endings are
        normalised to CRLF.

        Files whose stem does not match any module are reported via
        ``strict``:

        - ``strict=False`` (default): unmatched files are ignored.
        - ``strict=True``: raises :class:`KeyError` on the first
          unmatched file.

        Does **not** write to disk — call :meth:`save` afterwards to
        persist.

        Returns the list of module names that were updated.
        """
        src = Path(src_dir)
        if not src.is_dir():
            raise NotADirectoryError(f"Not a directory: {src}")

        project = self.vba_project()
        by_name = {m.name.casefold(): m for m in project.modules}
        updated: list[str] = []

        for child in sorted(src.iterdir()):
            if not child.is_file():
                continue
            if child.suffix.lower() not in _SOURCE_EXTS:
                continue
            key = child.stem.casefold()
            module = by_name.get(key)
            if module is None:
                if strict:
                    raise KeyError(
                        f"No module matches file {child.name!r} "
                        f"(known: {sorted(m.name for m in project.modules)})."
                    )
                continue
            raw = child.read_bytes().decode(encoding, errors="replace")
            text = raw.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
            if module.source != text:
                module.source = text
                module.dirty = True
            updated.append(module.name)
        return updated

    def save(self, dest: Union[str, Path, None] = None) -> None:
        """
        Save the workbook, applying any pending module edits.

        ``dest`` defaults to the original file path (in-place overwrite).

        Only ``xl/vbaProject.bin`` is rewritten; every other ZIP entry is
        preserved byte-for-byte along with its compression method and
        metadata so the workbook's non-VBA structure remains intact.

        Legacy ``.xls`` (raw CFB) writes the CFB bytes directly.
        """
        cfb = self._get_cfb()
        if self._project is not None:
            write_back_modules(cfb, self._project)
        # [MS-OVBA] writers MUST NOT emit performance-cache (__SRP_*) streams.
        try:
            cfb.drop_streams_in_storage("VBA", lambda n: n.startswith("__SRP_"))
        except KeyError:
            pass
        new_cfb_bytes = cfb.to_bytes()
        out_path = Path(dest) if dest is not None else self._path

        if self._suffix in _CFB_FORMATS:
            out_path.write_bytes(new_cfb_bytes)
            return

        if self._zip is None:
            raise RuntimeError("ExcelFile is not open.")

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as out_zip:
            for info in self._zip.infolist():
                if info.filename == _VBA_ENTRY:
                    out_info = zipfile.ZipInfo(
                        filename=info.filename,
                        date_time=info.date_time,
                    )
                    out_info.compress_type = info.compress_type
                    out_info.external_attr = info.external_attr
                    out_info.create_system = info.create_system
                    out_zip.writestr(out_info, new_cfb_bytes)
                else:
                    data = self._zip.read(info.filename)
                    out_info = zipfile.ZipInfo(
                        filename=info.filename,
                        date_time=info.date_time,
                    )
                    out_info.compress_type = info.compress_type
                    out_info.external_attr = info.external_attr
                    out_info.create_system = info.create_system
                    out_zip.writestr(out_info, data)
        out_path.write_bytes(buf.getvalue())

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
        self._zip_bytes = raw
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
