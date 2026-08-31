"""
Shared host-file implementation for Excel / Word / PowerPoint facades.

Each host format is a thin subclass of :class:`VBAHostFile` that supplies
five class attributes (container extensions, the vbaProject.bin entry
path, and two message fragments) plus its own ``create_new()``.  All
reading, editing, pull/push, and the safety-gated ``save()`` pipeline
live here exactly once.

The class is private: the public API surface remains
``pyopenvba.ExcelFile`` / ``WordFile`` / ``PowerPointFile``.
"""

from __future__ import annotations

import io
import warnings
import zipfile
from pathlib import Path
from typing import ClassVar, TypeVar

from pyopenvba.cfb import CFB
from pyopenvba.exceptions import UnsupportedFormatError, VBAProjectError
from pyopenvba.forms import VBAForm, read_forms
from pyopenvba.vba import (
    VBAModuleKind,
    VBAProject,
    compress,
    detect_signature,
    invalidate_vba_project_cache,
    normalize_class_source,
    parse_vba_project,
    rebuild_module_stream,
    serialize_dir_stream,
    serialize_project_stream,
    serialize_projectwm,
    split_attribute_header,
    write_back_modules,
)

# File extensions used by the VBE export/import workflow.
# - Standard procedural modules -> .bas
# - Everything else (class, document/sheet/workbook, designer/form) -> .cls
# We do not write .frm/.frx layout bytes; UserForm layout is preserved
# inside the CFB and never round-trips through disk.
_BAS_EXT = ".bas"
_CLS_EXT = ".cls"
_SOURCE_EXTS = frozenset({_BAS_EXT, _CLS_EXT})

_HostT = TypeVar("_HostT", bound="VBAHostFile")


class VBAHostFile:
    """Open an Office file and provide access to its VBA project.

    Subclasses define the container parameters:

    - ``_zip_formats``: extensions stored as OOXML ZIP containers.
    - ``_cfb_formats``: legacy extensions stored as a CFB container.
    - ``_vba_entry``: ZIP entry path of ``vbaProject.bin``.
    - ``_host_noun``: "workbook" / "document" / "presentation", used in
      user-facing messages.
    - ``_no_vba_hint``: sentence appended when the ZIP has no VBA entry.

    For ``.doc`` and ``.xls`` the legacy container *is* the VBA project's
    CFB, so the two extraction hooks below are identities.  ``.ppt``
    embeds the project deeper and overrides them.
    """

    _zip_formats: ClassVar[frozenset[str]]
    _cfb_formats: ClassVar[frozenset[str]]
    _vba_entry: ClassVar[str]
    _host_noun: ClassVar[str]
    _no_vba_hint: ClassVar[str]

    def __init__(self, path: str | Path) -> None:
        self._path = Path(path)
        self._suffix = self._path.suffix.lower()
        self._zip: zipfile.ZipFile | None = None
        self._cfb: CFB | None = None
        self._project: VBAProject | None = None
        self._container_raw: bytes = b""
        self._forms: list[VBAForm] | None = None
        self._open()

    # ------------------------------------------------------------------
    # Context manager
    # ------------------------------------------------------------------

    def __enter__(self: _HostT) -> _HostT:
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
        """Return the raw bytes of the VBA project CFB: the
        ``vbaProject.bin`` ZIP entry, or the project extracted from a
        legacy container."""
        if self._suffix in self._cfb_formats:
            return self._vba_cfb_bytes(self._path.read_bytes())
        assert self._zip is not None
        return self._zip.read(self._vba_entry)

    def vba_modules(self) -> dict[str, str]:
        """Return a mapping of module name -> source code."""
        return {m.name: m.source for m in self.vba_project().modules}

    def forms(self) -> list[VBAForm]:
        """Return the UserForm designer surfaces: controls and their nesting.

        The code-behind of a form is a module like any other; this is the
        *design* beside it, which no module source carries.  Raises
        :class:`~pyopenvba.exceptions.FormParseError` if a form's designer
        streams do not reconcile.

        The result is cached, so property edits made on it are the ones
        :meth:`save` writes back.
        """
        if self._forms is None:
            self._forms = read_forms(
                self._get_cfb(), code_page=self.vba_project().code_page
            )
        return self._forms

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
        automatically re-prepended so host-bound document modules keep
        their ``Attribute VB_*`` lines.  This mirrors the VBE UX where the
        user only types the body.

        When the target is a class-kind module (``VBAModuleKind.other``),
        a supplied full source is normalized from file-export form to
        stream form first (``VERSION ... CLASS`` preamble stripped,
        ``Attribute VB_Base`` preserved or restored), so ``.cls`` files
        exported from the VBE are accepted as-is.

        Changes are not written to disk until :meth:`save` is called.
        """
        project = self.vba_project()
        for m in project.modules:
            if m.name.casefold() == name.casefold():
                supplied_header, _ = split_attribute_header(source)
                if supplied_header and m.kind == VBAModuleKind.other:
                    # Convert file-export form to stream form.  prior_header
                    # keeps the module's existing VB_Base line (including
                    # host CLSIDs on document modules, which share the
                    # 0x0022 module kind with plain classes).
                    prior = m.attribute_header or split_attribute_header(m.source)[0]
                    source = normalize_class_source(source, prior_header=prior)
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
        format.  UserForm layout (``.frx``) is **not** exported — it is
        preserved verbatim inside the file on save.

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
        source replaces the module's current source.  Line endings are
        normalised to CRLF.  Class-kind files in VBE export form are
        normalized to stream form (see
        :func:`pyopenvba.vba.normalize_class_source`).

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
            new_header = ""
            if module.kind == VBAModuleKind.other and split_attribute_header(text)[0]:
                # Class-kind files may be VBE exports (file-export form);
                # convert to stream form, preserving the module's existing
                # VB_Base line via prior_header.
                prior = module.attribute_header or split_attribute_header(module.source)[0]
                text = normalize_class_source(text, prior_header=prior)
                new_header = split_attribute_header(text)[0]
            if module.source != text:
                module.source = text
                if new_header:
                    module.attribute_header = new_header
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
        Save the file, applying any pending module edits.

        ``dest`` defaults to the original file path (in-place overwrite).

        Only the ``vbaProject.bin`` entry is rewritten; every other ZIP
        entry is preserved byte-for-byte along with its compression
        method and metadata so the file's non-VBA structure remains
        intact.  Legacy raw-CFB formats write the CFB bytes directly.

        Safety gates:

        - If the project is password-protected (``has_password``) and the
          save would emit any change, raise ``VBAProjectError`` unless
          ``allow_protected=True`` is passed.  Saving a protected project
          without re-encrypting the password material would leave the
          file in an inconsistent state.
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
            # Office-saved files, so these dicts drive both the CFB-level
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
                    f"the {self._host_noun} inconsistent)."
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

            # 2. Create brand-new streams for pending adds.  Sorted order
            #    keeps the emitted PROJECT declarations and CFB layout
            #    byte-deterministic across processes.
            add_modules_for_project: list[tuple[str, str]] = []
            for name in sorted(add_names):
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
            for name in sorted(delete_names):
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
                        code_page=project.code_page,
                    )
                    cfb.write_stream("PROJECT", new_project)
                # Rewrite PROJECTwm to enumerate the current module set in
                # both MBCS and Unicode forms.  Required whenever the module
                # identity set changes ([MS-OVBA] 2.3.4.4).
                # PROJECTwm lives at the project root as a sibling of the
                # VBA storage ([MS-OVBA] 2.2.1), so it is addressed without
                # a storage qualifier.
                try:
                    cfb.get_stream("PROJECTwm")
                except KeyError:
                    pass
                else:
                    wm_pairs = [
                        (m.name, m.name_unicode or m.name)
                        for m in project.modules
                    ]
                    cfb.write_stream(
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
        # Designer edits live beside the VBA storage, so they are written
        # whether or not any module changed.  write_back() touches only the
        # streams whose bytes actually differ, so an unedited form is a
        # no-op here and the CFB stays byte-identical.
        for form in self._forms or ():
            form.write_back(cfb)

        # [MS-OVBA] writers MUST NOT emit performance-cache (__SRP_*) streams.
        try:
            cfb.drop_streams_in_storage("VBA", lambda n: n.startswith("__SRP_"))
        except KeyError:
            pass
        new_cfb_bytes = cfb.to_bytes()
        out_path = Path(dest) if dest is not None else self._path

        if self._suffix in self._cfb_formats:
            out_path.write_bytes(self._container_bytes(new_cfb_bytes))
            return

        if self._zip is None:
            raise RuntimeError(f"{type(self).__name__} is not open.")

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as out_zip:
            for info in self._zip.infolist():
                if info.filename == self._vba_entry:
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
        if self._suffix in self._zip_formats:
            self._open_zip()
        elif self._suffix in self._cfb_formats:
            self._open_cfb_direct()
        else:
            raise UnsupportedFormatError(
                f"Unsupported file extension: {self._suffix!r}. "
                f"Supported: {sorted(self._zip_formats | self._cfb_formats)}"
            )

    def _open_zip(self) -> None:
        raw = self._path.read_bytes()
        self._zip = zipfile.ZipFile(io.BytesIO(raw), mode="r")
        if self._vba_entry not in self._zip.namelist():
            raise VBAProjectError(
                f"{self._path.name!r} contains no {self._vba_entry!r}. "
                + self._no_vba_hint
            )

    def _vba_cfb_bytes(self, container: bytes) -> bytes:
        """Extract the VBA project CFB from a legacy container's bytes.

        The default is the identity: for ``.doc`` and ``.xls`` the file
        itself is the project's CFB.
        """
        return container

    def _container_bytes(self, vba_cfb: bytes) -> bytes:
        """Rebuild the legacy container around a modified project CFB.

        The inverse of :meth:`_vba_cfb_bytes`, and the identity for the
        formats whose container is the project.
        """
        return vba_cfb

    def _open_cfb_direct(self) -> None:
        self._container_raw = self._path.read_bytes()
        self._cfb = CFB.from_bytes(self._vba_cfb_bytes(self._container_raw))

    def _get_cfb(self) -> CFB:
        if self._cfb is not None:
            return self._cfb
        if self._zip is None:
            raise RuntimeError(f"{type(self).__name__} is not open.")
        vba_bin = self._zip.read(self._vba_entry)
        self._cfb = CFB.from_bytes(vba_bin)
        return self._cfb
