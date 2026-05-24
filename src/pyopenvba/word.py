"""
Word file handler.

Supports:
  - .docm  (OOXML macro-enabled document -- ZIP containing word/vbaProject.bin)
  - .dotm  (OOXML macro-enabled template -- ZIP containing word/vbaProject.bin)
  - .doc   (Legacy Word -- the entire file is a CFB)

Usage
-----
    with WordFile("document.docm") as doc:
        project = doc.vba_project()       # -> VBAProject
        modules = doc.vba_modules()       # -> dict[str, str]
        doc.set_module("Module1", src)
        doc.save("document_out.docm")
"""

from __future__ import annotations

import io
import warnings
import zipfile
from pathlib import Path
from typing import Union

from pyopenvba.cfb import CFB
from pyopenvba.exceptions import UnsupportedFormatError, VBAProjectError
from pyopenvba.vba import VBAProject, parse_vba_project, write_back_modules
from pyopenvba.vba import VBAModuleKind
from pyopenvba.vba import (
    compress,
    detect_signature,
    invalidate_vba_project_cache,
    rebuild_module_stream,
    serialize_dir_stream,
    serialize_project_stream,
    serialize_projectwm,
)

_ZIP_FORMATS = frozenset({".docm", ".dotm"})
_CFB_FORMATS = frozenset({".doc"})
_VBA_ENTRY = "word/vbaProject.bin"

_BAS_EXT = ".bas"
_CLS_EXT = ".cls"
_SOURCE_EXTS = frozenset({_BAS_EXT, _CLS_EXT})


class WordFile:
    """
    Open a Word file and provide access to its VBA project.

    Can be used as a context manager::

        with WordFile("document.docm") as doc:
            ...
    """

    def __init__(self, path: Union[str, Path]) -> None:
        self._path = Path(path)
        self._suffix = self._path.suffix.lower()
        self._zip: zipfile.ZipFile | None = None
        self._cfb: CFB | None = None
        self._project: VBAProject | None = None
        self._open()

    # ------------------------------------------------------------------
    # Constructors
    # ------------------------------------------------------------------

    @classmethod
    def create_new(cls, path: Union[str, Path]) -> "WordFile":
        """
        Create a new macro-enabled document (``.docm``) at ``path``
        containing an empty VBA project (``ThisDocument`` and a bare
        ``Module1``) and return an open :class:`WordFile` for it.

        The bytes are decoded from a baked-in template captured from a
        freshly Word-authored document, so the resulting file opens
        cleanly in Word without any repair prompt.

        ``path`` is overwritten if it already exists.
        """
        from pyopenvba._templates import EMPTY_DOCM_BYTES

        target = Path(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(EMPTY_DOCM_BYTES)
        return cls(target)

    # ------------------------------------------------------------------
    # Context manager
    # ------------------------------------------------------------------

    def __enter__(self) -> "WordFile":
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
        """Return the raw bytes of ``word/vbaProject.bin`` (or the whole CFB
        file for legacy ``.doc``)."""
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

        ``source`` may be either a full source replacement (starting with
        ``Attribute VB_*`` or ``VERSION ... CLASS``) or a bare body.  When
        a bare body is supplied, the module's existing attribute header is
        automatically re-prepended so document modules keep their host-binding
        ``Attribute VB_*`` lines.

        Changes are not written to disk until :meth:`save` is called.
        """
        from pyopenvba.vba import split_attribute_header
        project = self.vba_project()
        for m in project.modules:
            if m.name.casefold() == name.casefold():
                supplied_header, _ = split_attribute_header(source)
                if supplied_header:
                    m.source = source
                    m.attribute_header = supplied_header
                else:
                    header = m.attribute_header or split_attribute_header(m.source)[0]
                    m.attribute_header = header
                    m.source = header + source
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
        Source bytes use CRLF line endings to match VBE's own export format.

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
        insensitively) to a module of the same logical name and its source
        replaces the module's current source.  Line endings are normalised
        to CRLF.

        Files whose stem does not match any module are silently skipped
        (``strict=False``) or raise :class:`KeyError` (``strict=True``).

        Does **not** write to disk -- call :meth:`save` afterwards.

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

    def save(
        self,
        dest: Union[str, Path, None] = None,
        *,
        allow_protected: bool = False,
        allow_invalidate_signature: bool = False,
    ) -> None:
        """
        Save the document, applying any pending module edits.

        ``dest`` defaults to the original file path (in-place overwrite).

        Only ``word/vbaProject.bin`` is rewritten; every other ZIP entry is
        preserved byte-for-byte.

        Legacy ``.doc`` (raw CFB) writes the CFB bytes directly.

        Safety gates mirror those of :class:`~pyopenvba.excel.ExcelFile`:

        - Password-protected projects raise ``VBAProjectError`` on mutation
          unless ``allow_protected=True`` is passed.
        - Digital-signature streams are dropped on mutation with a
          ``UserWarning`` unless ``allow_invalidate_signature=True``.
        """
        cfb = self._get_cfb()
        if self._project is not None:
            project = self._project
            rename_map = dict(project.pending_renames)
            add_names = set(project.pending_adds)
            delete_names = set(project.pending_deletes)
            has_source_edits = any(m.dirty for m in project.modules)
            mutating = bool(
                rename_map or add_names or delete_names or has_source_edits
            )

            if (
                mutating
                and project.protection is not None
                and project.protection.has_password
                and not allow_protected
            ):
                raise VBAProjectError(
                    "Refusing to save: the VBA project is password-protected. "
                    "Pass allow_protected=True to override (the password "
                    "material will be preserved verbatim, which may leave "
                    "the document inconsistent)."
                )

            if mutating:
                sig_info = detect_signature(cfb)
                if sig_info.present:
                    for sig_stream in (
                        "_VBA_PROJECT_SIGNATURE",
                        "_VBA_PROJECT_SIGNATURE_AGILE",
                        "_VBA_PROJECT_SIGNATURE_V3",
                    ):
                        try:
                            cfb.remove_stream_in_storage("VBA", sig_stream)
                        except KeyError:
                            pass
                        try:
                            cfb.remove_stream(sig_stream)
                        except KeyError:
                            pass
                    if not allow_invalidate_signature:
                        warnings.warn(
                            "Dropped stale VBA digital signature streams "
                            f"({', '.join(sig_info.kinds)}) because the "
                            "project was modified.  Re-sign externally to "
                            "restore trust.  Pass "
                            "allow_invalidate_signature=True to silence.",
                            UserWarning,
                            stacklevel=2,
                        )

            for old, new in rename_map.items():
                try:
                    cfb.rename_stream_in_storage("VBA", old, new)
                except KeyError:
                    pass

            add_modules_for_project: list[tuple[str, str]] = []
            for name in add_names:
                module = next(
                    (m for m in project.modules if m.stream_name == name), None
                )
                if module is None:
                    continue
                seed = rebuild_module_stream(module, project.code_page)
                try:
                    cfb.add_stream_to_storage("VBA", name, seed)
                except ValueError:
                    cfb.write_stream_in_storage("VBA", name, seed)
                module.dirty = False
                decl_key = (
                    "Module" if module.kind == VBAModuleKind.standard else "Class"
                )
                add_modules_for_project.append((module.name, decl_key))

            for name in delete_names:
                try:
                    cfb.remove_stream_in_storage("VBA", name)
                except KeyError:
                    pass

            project.pending_renames.clear()
            project.pending_adds.clear()
            project.pending_deletes.clear()

            write_back_modules(cfb, project)

            if project.dir_structure_dirty:
                new_dir_raw = serialize_dir_stream(project)
                cfb.write_stream_in_storage("VBA", "dir", compress(new_dir_raw))
                try:
                    project_raw = cfb.get_stream("PROJECT")
                except KeyError:
                    project_raw = None
                if project_raw is not None:
                    new_project = serialize_project_stream(
                        project_raw,
                        rename_map,
                        add_modules=add_modules_for_project,
                        delete_names=delete_names,
                    )
                    cfb.write_stream("PROJECT", new_project)
                try:
                    cfb.get_stream_in_storage("VBA", "PROJECTwm")
                except KeyError:
                    pass
                else:
                    wm_pairs = [
                        (m.name, m.name_unicode or m.name)
                        for m in project.modules
                    ]
                    cfb.write_stream_in_storage(
                        "VBA",
                        "PROJECTwm",
                        serialize_projectwm(wm_pairs, code_page=project.code_page),
                    )
                project.dir_structure_dirty = False

            if mutating:
                invalidate_vba_project_cache(cfb)

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
            raise RuntimeError("WordFile is not open.")

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
        self._zip = zipfile.ZipFile(io.BytesIO(raw), mode="r")
        if _VBA_ENTRY not in self._zip.namelist():
            raise VBAProjectError(
                f"{self._path.name!r} contains no {_VBA_ENTRY!r}. "
                "Make sure the document has a VBA project (save as .docm in Word)."
            )

    def _open_cfb_direct(self) -> None:
        raw = self._path.read_bytes()
        self._cfb = CFB.from_bytes(raw)

    def _get_cfb(self) -> CFB:
        if self._cfb is not None:
            return self._cfb
        if self._zip is None:
            raise RuntimeError("WordFile is not open.")
        vba_bin = self._zip.read(_VBA_ENTRY)
        self._cfb = CFB.from_bytes(vba_bin)
        return self._cfb
