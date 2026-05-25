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
import warnings
import zipfile
from pathlib import Path

from pyopenvba.cfb import CFB
from pyopenvba.exceptions import UnsupportedFormatError, VBAProjectError
from pyopenvba.vba import (
    VBAModuleKind,
    VBAProject,
    compress,
    detect_signature,
    invalidate_vba_project_cache,
    parse_vba_project,
    rebuild_module_stream,
    serialize_dir_stream,
    serialize_project_stream,
    serialize_projectwm,
    write_back_modules,
)

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

    def __init__(self, path: str | Path) -> None:
        self._path = Path(path)
        self._suffix = self._path.suffix.lower()
        self._zip: zipfile.ZipFile | None = None
        self._cfb: CFB | None = None
        self._zip_bytes: bytes | None = None
        self._project: VBAProject | None = None
        self._open()

    # ------------------------------------------------------------------
    # Constructors
    # ------------------------------------------------------------------

    @classmethod
    def create_new(cls, path: str | Path) -> ExcelFile:
        """
        Create a new macro-enabled workbook at ``path`` containing an empty
        VBA project (``ThisWorkbook``, ``Sheet1``, and a bare ``Module1``)
        and return an open :class:`ExcelFile` for it.

        Supported extensions: ``.xlsm`` (default) and ``.xlsb``.

        The bytes are decoded from a baked-in template captured from a
        freshly Excel-authored workbook, so the resulting file opens
        cleanly in Excel without any "found a problem" repair prompt.

        ``path`` is overwritten if it already exists.
        """
        target = Path(path)
        suffix = target.suffix.lower()
        if suffix == ".xlsb":
            from pyopenvba._templates import EMPTY_XLSB_BYTES
            template = EMPTY_XLSB_BYTES
        else:
            from pyopenvba._templates import EMPTY_XLSM_BYTES
            template = EMPTY_XLSM_BYTES
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(template)
        return cls(target)

    # ------------------------------------------------------------------
    # Context manager
    # ------------------------------------------------------------------

    def __enter__(self) -> ExcelFile:
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

        ``source`` may be either a full source replacement (starting with
        ``Attribute VB_*`` or ``VERSION ... CLASS``) or a bare body.  When
        a bare body is supplied, the module's existing attribute header is
        automatically re-prepended so document modules (``ThisWorkbook``,
        ``Sheet1``, ...) keep their host-binding ``Attribute VB_*`` lines.
        This mirrors the VBE UX where the user only types the body.

        Changes are not written to disk until :meth:`save` is called.
        """
        from pyopenvba.vba import split_attribute_header
        project = self.vba_project()
        for m in project.modules:
            if m.name.casefold() == name.casefold():
                supplied_header, _ = split_attribute_header(source)
                if supplied_header:
                    # Full-source replacement — also refresh the cached header.
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
        dest_dir: str | Path,
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
        src_dir: str | Path,
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

    def save(
        self,
        dest: str | Path | None = None,
        *,
        allow_protected: bool = False,
        allow_invalidate_signature: bool = False,
    ) -> None:
        """
        Save the workbook, applying any pending module edits.

        ``dest`` defaults to the original file path (in-place overwrite).

        Only ``xl/vbaProject.bin`` is rewritten; every other ZIP entry is
        preserved byte-for-byte along with its compression method and
        metadata so the workbook's non-VBA structure remains intact.

        Legacy ``.xls`` (raw CFB) writes the CFB bytes directly.

        Safety gates:

        - If the project is password-protected (``has_password``) and the
          save would emit any change, raise ``VBAProjectError`` unless
          ``allow_protected=True`` is passed.  Saving a protected project
          without re-encrypting the password material would leave the
          workbook in an inconsistent state.
        - If the project carries any digital-signature stream and the
          save would emit any change, the existing signature streams are
          dropped (they are guaranteed to be stale) and a
          ``UserWarning`` is emitted.  Set
          ``allow_invalidate_signature=True`` to silence the warning.
        - If the save would emit any change, the ``_VBA_PROJECT``
          performance cache body is zeroed (header preserved) so Office
          regenerates the cache on next open ([MS-OVBA] 2.3.4.1).
        """
        cfb = self._get_cfb()
        if self._project is not None:
            project = self._project
            # Snapshot pending mutations.  Logical name == stream name in
            # Excel-saved files, so these dicts drive both the CFB-level
            # operations and the PROJECT-stream rewrite.
            rename_map = dict(project.pending_renames)
            add_names = set(project.pending_adds)
            delete_names = set(project.pending_deletes)
            has_source_edits = any(m.dirty for m in project.modules)
            mutating = bool(
                rename_map or add_names or delete_names or has_source_edits
            )

            # Safety gate 1: refuse to mutate a password-protected project
            # unless the caller explicitly opts in.
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
                    "the workbook inconsistent)."
                )

            # Safety gate 2: any change invalidates a present digital
            # signature.  Drop the stale signature streams and warn.
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

            # 1. Apply renames first so that pre-existing streams are at
            #    their new names before any other lookup runs.
            for old, new in rename_map.items():
                try:
                    cfb.rename_stream_in_storage("VBA", old, new)
                except KeyError:
                    pass

            # 2. Create brand-new streams for pending adds.
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
                    # Stream already exists (e.g. add-then-save called twice).
                    cfb.write_stream_in_storage("VBA", name, seed)
                module.dirty = False
                decl_key = (
                    "Module" if module.kind == VBAModuleKind.standard else "Class"
                )
                add_modules_for_project.append((module.name, decl_key))

            # 3. Delete streams the user removed in-memory.
            for name in delete_names:
                try:
                    cfb.remove_stream_in_storage("VBA", name)
                except KeyError:
                    pass

            project.pending_renames.clear()
            project.pending_adds.clear()
            project.pending_deletes.clear()

            # 4. Replace contents of any remaining dirty (pre-existing) modules.
            write_back_modules(cfb, project)

            # 5. Rewrite the dir + PROJECT streams when the module set's
            #    identity has changed (add / rename / delete).  PROJECT is
            #    always rewritten on a structural save so that any duplicate
            #    declarations or stale ``[Workspace]`` entries left behind by
            #    earlier buggy writes are scrubbed via the dedup pass in
            #    ``serialize_project_stream``.
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
                # Rewrite PROJECTwm to enumerate the current module set in
                # both MBCS and Unicode forms.  Required whenever the module
                # identity set changes ([MS-OVBA] 2.3.4.4).
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

            # 6. Invalidate the _VBA_PROJECT performance cache so Office
            #    regenerates it on next open ([MS-OVBA] 2.3.4.1 -- the
            #    cache MUST be ignored on read; the verbatim cache may
            #    reference offsets that no longer match the updated
            #    module set or source).
            if mutating:
                invalidate_vba_project_cache(cfb)
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
