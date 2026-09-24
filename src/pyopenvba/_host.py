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
from pyopenvba._new_project import with_project
from pyopenvba._package_signature import signature_parts, without_signature
from pyopenvba._references import ReferenceManager, module_offset, reference_spans
from pyopenvba.exceptions import NoVBAProjectError, UnsupportedFormatError, VBAProjectError
from pyopenvba.forms import VBAForm, create_form, form_names, read_forms
from pyopenvba.vba import (
    SignatureInfo,
    VBAModuleKind,
    VBAProject,
    VBAReference,
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
# We do not write .frm/.frx layout bytes: a form's design stays in the CFB
# and never round-trips through disk.  It is not opaque, though -- see
# forms.py, which reads and writes it in place.
_BAS_EXT = ".bas"
_CLS_EXT = ".cls"
_SOURCE_EXTS = frozenset({_BAS_EXT, _CLS_EXT})

_HostT = TypeVar("_HostT", bound="VBAHostFile")


class VBAHostFile(ReferenceManager):
    """Open an Office file and provide access to its VBA project.

    Subclasses define the container parameters:

    - ``_zip_formats``: extensions stored as OOXML ZIP containers.
    - ``_cfb_formats``: legacy extensions stored as a CFB container.
    - ``_vba_entry``: ZIP entry path of ``vbaProject.bin``.
    - ``_host_noun``: "workbook" / "document" / "presentation", used in
      user-facing messages.
    - ``_application``: "Excel" / "Word" / "PowerPoint", the application
      that creates a project when its first macro is written.
    - ``_project_storage``: the storage at the root of the legacy
      container that holds the project, or None where the project is
      found another way.

    For ``.doc`` and ``.xls`` the legacy container *is* the VBA project's
    CFB, so the two extraction hooks below are identities.  ``.ppt``
    embeds the project deeper and overrides them.

    A file saved before its first macro has no project at all: no
    ``vbaProject.bin`` in a zip, no project storage in a legacy file
    (tests/fixtures/no_vba/).  It opens all the same.  Listing reads
    answer empty, and a read or write that needs the project raises
    :class:`~pyopenvba.exceptions.NoVBAProjectError`.
    """

    _zip_formats: ClassVar[frozenset[str]]
    _cfb_formats: ClassVar[frozenset[str]]
    _vba_entry: ClassVar[str]
    _host_noun: ClassVar[str]
    _application: ClassVar[str]
    _project_storage: ClassVar[str | None]
    #: The formats add_vba_project gives a project to, each measured against its application.
    _project_formats: ClassVar[frozenset[str]] = frozenset()
    #: The package part a zip-based file's project is related from.
    _main_part: ClassVar[str] = ""

    def __init__(self, path: str | Path) -> None:
        self._path = Path(path)
        self._suffix = self._path.suffix.lower()
        self._zip: zipfile.ZipFile | None = None
        self._cfb: CFB | None = None
        self._project: VBAProject | None = None
        self._container_raw: bytes = b""
        self._forms: list[VBAForm] | None = None
        self._has_project = True
        # Set by add_vba_project: the project is not in the package yet, and these are its document modules.
        self._project_added = False
        self._document_names: set[str] = set()
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

    def has_vba_project(self) -> bool:
        """Whether the file has a VBA project, empty or not.

        A macro-enabled file saved before its first macro, and a binary
        file that never held one, have none.  That is their normal shape
        rather than damage: listing reads answer empty, and a read or
        write of the project, a module, a form or a reference raises
        :class:`~pyopenvba.exceptions.NoVBAProjectError`.  A project that
        holds no modules is still a project.  :meth:`add_vba_project`
        gives a ``.xlsm``, ``.docm`` or ``.pptm`` one.
        """
        return self._has_project

    def add_vba_project(self) -> VBAProject:
        """Give a file with no VBA project one, as its application does for a first macro.

        The project holds what the application puts in a new one: in
        Excel a document module for the workbook and one for each sheet,
        named as Excel names them, in Word ``ThisDocument``, and in
        PowerPoint nothing.  Add modules to it with
        ``vba_project().add_module(...)``, then :meth:`save`, which writes
        the project into the package as the application does -- except
        a project the application would not write: Excel writes only the
        code names for one with no code in it, and PowerPoint nothing for
        one with no module (see :mod:`pyopenvba._new_project`).

        Supported for ``.xlsm``, ``.docm`` and ``.pptm``; a file that has
        a project already raises
        :class:`~pyopenvba.exceptions.VBAProjectError`.
        """
        if self._has_project:
            raise VBAProjectError(f"{self._path.name!r} already has a VBA project.")
        if self._zip is None or self._suffix not in self._project_formats:
            supported = ", ".join(sorted(self._project_formats))
            raise UnsupportedFormatError(
                f"Adding a VBA project to a {self._suffix} file is not supported; "
                f"it is for {supported}, whose applications were measured doing it."
            )
        wanted = self._new_documents()
        with zipfile.ZipFile(io.BytesIO(self._project_template())) as template:
            cfb = CFB.from_bytes(template.read(self._vba_entry))
        project = parse_vba_project(cfb)
        headers = {name.casefold(): header for name, header in wanted}
        for module in list(project.modules):
            key = module.name.casefold()
            unwanted = key not in headers
            reheaded = headers.get(key) is not None and module.attribute_header != headers[key]
            if unwanted or reheaded:
                # Taken out; one wanted under another header comes back under its own below.
                project.delete_module(module.name)
        present = {module.name.casefold() for module in project.modules}
        for name, header in wanted:
            if name.casefold() not in present and header is not None:
                project.add_module(name, header, kind=VBAModuleKind.other)
        # Written through now, so a module added next is declared after these, as the application declares it.
        self._document_names = set(headers)
        self._apply_project(cfb, project, mutating=True)
        project = parse_vba_project(cfb)
        self._cfb, self._project, self._forms = cfb, project, None
        self._has_project = self._project_added = True
        return project

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
        if not self._has_project:
            raise NoVBAProjectError(self._no_project_message())
        if self._project_added:
            # Made by add_vba_project: the project as it stands, edits waiting for save() aside.
            return self._get_cfb().to_bytes()
        if self._suffix in self._cfb_formats:
            return self._vba_cfb_bytes(self._path.read_bytes())
        assert self._zip is not None
        return self._zip.read(self._vba_entry)

    def _reference_data(self) -> tuple[bytes, int]:
        project = self.vba_project()
        return project.dir_raw, project.code_page

    def _write_reference_data(self, raw: bytes) -> None:
        offset = module_offset(raw)
        references = [span.reference for span in reference_spans(raw)]
        project = self.vba_project()
        project.dir_raw, project.dir_modules_offset = raw, offset
        project.references = references
        project.dir_references_dirty = True

    def _reference_host(self) -> str:
        return {"workbook": "excel", "document": "word", "presentation": "powerpoint"}[self._host_noun]

    def _reference_forms(self) -> list[str]:
        return form_names(self._get_cfb())

    def references(self) -> list[VBAReference]:
        """Declared libraries in priority order; none in a file with no project."""
        return super().references() if self._has_project else []

    def vba_modules(self) -> dict[str, str]:
        """Return a mapping of module name -> source code."""
        if not self._has_project:
            return {}
        return {m.name: m.source for m in self.vba_project().modules}

    def vba_signature(self) -> SignatureInfo:
        """The VBA project's digital signature, wherever the file keeps it.

        That is the signature streams inside the project, and in a
        zip-based file the parts beside ``vbaProject.bin``, which is where
        Excel, Word and PowerPoint keep one. A save that changes the
        project drops it.
        """
        if not self._has_project:
            return SignatureInfo()
        info = detect_signature(self._get_cfb())
        if self._zip is not None:
            parts = signature_parts(self._zip.namelist(), self._zip.read, self._vba_entry)
            info.kinds += [kind for kind in dict.fromkeys(parts.values()) if kind not in info.kinds]
            info.parts = list(parts)
            info.present = info.present or bool(parts)
        return info

    def forms(self) -> list[VBAForm]:
        """Return the UserForm designer surfaces: controls and their nesting.

        The code-behind of a form is a module like any other; this is the
        *design* beside it, which no module source carries.  Raises
        :class:`~pyopenvba.exceptions.FormParseError` if a form's designer
        streams do not reconcile.

        The result is cached, so property edits made on it are the ones
        :meth:`save` writes back.
        """
        if not self._has_project:
            return []
        if self._forms is None:
            self._forms = read_forms(
                self._get_cfb(), code_page=self.vba_project().code_page
            )
        return self._forms

    def add_form(
        self,
        name: str,
        *,
        caption: str | None = None,
        width: float | None = None,
        height: float | None = None,
    ) -> VBAForm:
        """Add an empty UserForm to the project.

        Creates the designer storage and the code-behind module together:
        a storage without a module is not a component the host will show,
        and a module without a storage is a class rather than a form.
        Geometry is in points; the defaults are the size Excel gives a new
        form.

        The form is returned ready to edit -- ``add_control`` and friends
        work on it straight away -- and lands on disk at :meth:`save`.
        """
        project = self.vba_project()
        cfb = self._get_cfb()
        header = create_form(
            cfb,
            name,
            caption=caption,
            width=width,
            height=height,
            code_page=project.code_page,
        )
        project.add_module(name, header, kind=VBAModuleKind.other)
        # The cache was read before this form existed.
        self._forms = None
        return next(f for f in self.forms() if f.name == name)

    def module_names(self) -> list[str]:
        """Return the list of VBA module names."""
        if not self._has_project:
            return []
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
        if not self._has_project:
            return []
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

        Returns the list of file paths written, none for a file with no
        project.
        """
        out_dir = Path(dest_dir)
        out_dir.mkdir(parents=True, exist_ok=True)

        written: list[Path] = []
        for m in self.vba_project().modules if self._has_project else ():
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
        persist.  A file with no VBA project refuses with
        :class:`~pyopenvba.exceptions.NoVBAProjectError`: its first macro
        has to be written in the host application.

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
        - If the project carries a digital signature and the save would
          emit any change, the signature is dropped, since it is
          guaranteed to be stale, and a ``UserWarning`` is emitted.  In a
          zip-based file that takes out the signature parts beside
          ``vbaProject.bin``, their relationships and their Overrides, as
          Excel, Word and PowerPoint do; any signature streams inside the
          project go too.  Set ``allow_invalidate_signature=True`` to
          silence the warning.
        - If the save would emit any change, the ``_VBA_PROJECT``
          performance cache body is zeroed (header preserved) so Office
          regenerates the cache on next open ([MS-OVBA] 2.3.4.1).  An edit
          to a UserForm's design counts: adding or removing a control
          changes the form class's members.

        A file with no VBA project can hold no edit, since every write
        refuses, so it is written out as it was read.
        """
        if not self._has_project:
            if dest is not None:
                Path(dest).write_bytes(self._container_raw)
            return
        cfb = self._get_cfb()

        # Designer edits land first, and count as a mutation: adding or
        # removing a control changes the form class's members, so leaving
        # the _VBA_PROJECT performance cache in place makes Office load a
        # member list that no longer matches the form.  Measured -- with
        # the cache left alone, Excel refuses the edited form outright.
        # write_back only sets in-memory overrides, so a gate below can
        # still refuse the save without anything having been written.
        forms_dirty = False
        for form in self._forms or ():
            # Not any(...): that short-circuits, and every form has to be
            # written, not just the first dirty one.
            forms_dirty |= form.write_back(cfb)
        # What the package's other parts need: None drops a part.
        package_edits: dict[str, bytes | None] = {}

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
                rename_map
                or add_names
                or delete_names
                or has_source_edits
                or forms_dirty
                or project.dir_references_dirty
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
            # signature.  Drop it and warn.
            if mutating:
                package_edits = self._drop_signature(cfb, allow_invalidate_signature)

            self._apply_project(cfb, project, mutating=mutating)
        if forms_dirty and self._project is None:
            # A designer-only edit on a host whose modules were never
            # parsed still has to invalidate the cache, and still leaves
            # a signature stale.
            invalidate_vba_project_cache(cfb)
            package_edits = self._drop_signature(cfb, allow_invalidate_signature)

        # [MS-OVBA] writers MUST NOT emit performance-cache (__SRP_*) streams.
        try:
            cfb.drop_streams_in_storage("VBA", lambda n: n.startswith("__SRP_"))
        except KeyError:
            pass
        new_cfb_bytes = cfb.to_bytes()
        self._write_container(dest, new_cfb_bytes, package_edits)

    def _apply_project(self, cfb: CFB, project: VBAProject, *, mutating: bool) -> None:
        """Write ``project``'s pending renames, adds, deletes and edits into ``cfb``."""
        rename_map = dict(project.pending_renames)
        add_names = set(project.pending_adds)
        delete_names = set(project.pending_deletes)
        # 1. Apply renames first so that pre-existing streams are at
        #    their new names before any other lookup runs.
        for old, new in rename_map.items():
            try:
                cfb.rename_stream_in_storage("VBA", old, new)
            except KeyError:
                pass

        # 2. Create brand-new streams for pending adds, in the order
        #    the modules were added, which is the order Office declares
        #    them in; being a list, it is also deterministic across
        #    processes, as a set's order is not.
        add_modules_for_project: list[tuple[str, str]] = []
        for module in project.modules:
            name = module.stream_name
            if name not in add_names:
                continue
            seed = rebuild_module_stream(module, project.code_page)
            try:
                cfb.add_stream_to_storage("VBA", name, seed)
            except ValueError:
                # Stream already exists (e.g. add-then-save called twice).
                cfb.write_stream_in_storage("VBA", name, seed)
            module.dirty = False
            if module.kind == VBAModuleKind.standard:
                decl_key = "Module"
            elif module.name in form_names(cfb):
                # A designer is declared BaseClass, not Class.  The
                # test is structural -- its name is also a storage
                # beside VBA/ -- which is the same one forms.py uses
                # to find forms at all.
                decl_key = "BaseClass"
            elif module.name.casefold() in self._document_names:
                decl_key = "Document"
            else:
                decl_key = "Class"
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
            project.dir_raw = new_dir_raw

        # Reference-only writes preserve every module metadata byte, but
        # still invalidate compiled state and honor the save safety gates.
        if project.dir_references_dirty:
            cfb.write_stream_in_storage("VBA", "dir", compress(project.dir_raw))
            project.dir_references_dirty = False

        # 6. Invalidate the _VBA_PROJECT performance cache so Office
        #    regenerates it on next open ([MS-OVBA] 2.3.4.1 -- the
        #    cache MUST be ignored on read; the verbatim cache may
        #    reference offsets that no longer match the updated
        #    module set or source).
        if mutating:
            invalidate_vba_project_cache(cfb)

    def _write_container(self, dest: str | Path | None, new_cfb_bytes: bytes,
                         package_edits: dict[str, bytes | None]) -> None:
        """Write the file with ``new_cfb_bytes`` as its project and ``package_edits`` applied."""
        out_path = Path(dest) if dest is not None else self._path

        if self._suffix in self._cfb_formats:
            out_path.write_bytes(self._container_bytes(new_cfb_bytes))
            return

        if self._zip is None:
            raise RuntimeError(f"{type(self).__name__} is not open.")

        # A project add_vba_project made goes into the package as the host
        # application writes one: with its relationship and content type,
        # and in Excel with the code names, which Excel writes even for a
        # project it does not write (see _new_project.py).
        added: dict[str, bytes] = {}
        if self._project_added and self._project is not None:
            package_edits.update(self._project_edits())
            if self._writes_project(self._project):
                package_edits.update(with_project(self._zip.namelist(), self._zip.read, self._main_part))
                added[self._vba_entry] = new_cfb_bytes

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
                    if info.filename in package_edits:
                        replacement = package_edits[info.filename]
                        if replacement is None:
                            continue
                        data = replacement
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
            for name, data in added.items():
                out_info = zipfile.ZipInfo(filename=name, date_time=(1980, 1, 1, 0, 0, 0))
                out_info.compress_type = zipfile.ZIP_DEFLATED
                out_zip.writestr(out_info, data)
        out_path.write_bytes(buf.getvalue())

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _drop_signature(self, cfb: CFB, allow_invalidate_signature: bool) -> dict[str, bytes | None]:
        """Take out a signature the save leaves stale, and warn unless told not to.

        Signature streams inside the project go from ``cfb``; the parts a
        zip-based file keeps beside it come back as the package's edits,
        None for each part that goes (see :mod:`pyopenvba._package_signature`).
        """
        found = detect_signature(cfb)
        if found.present:
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
        kinds = list(found.kinds)
        edits: dict[str, bytes | None] = {}
        if self._zip is not None:
            names = self._zip.namelist()
            kinds += [kind for kind in signature_parts(names, self._zip.read, self._vba_entry).values()
                      if kind not in kinds]
            edits = without_signature(names, self._zip.read, self._vba_entry)
        if kinds and not allow_invalidate_signature:
            warnings.warn(
                f"Dropped the stale VBA digital signature ({', '.join(kinds)}) "
                "because the project was modified.  Re-sign externally to "
                "restore trust.  Pass allow_invalidate_signature=True to "
                "silence.",
                UserWarning,
                stacklevel=3,
            )
        return edits

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
        self._container_raw = raw
        self._zip = zipfile.ZipFile(io.BytesIO(raw), mode="r")
        # Saved before its first macro, a macro-enabled file has no project part at all.
        self._has_project = self._vba_entry in self._zip.namelist()

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
        try:
            project = self._vba_cfb_bytes(self._container_raw)
        except NoVBAProjectError:
            self._has_project = False
            return
        cfb = CFB.from_bytes(project)
        # A binary file that never held a macro has no project storage at its root.
        storage = self._project_storage
        if storage is not None and storage.casefold() not in {name.casefold() for name in cfb.list_storages_at()}:
            self._has_project = False
            return
        self._cfb = cfb

    def _no_project_message(self) -> str:
        message = (
            f"{self._path.name!r} has no VBA project: it is a {self._host_noun} "
            f"with no macros.  The first macro has to be written in "
            f"{self._application}, which creates the project"
        )
        if self._suffix in self._project_formats:
            return message + ", or add_vba_project() gives it one."
        return message + "."

    def _project_template(self) -> bytes:
        """The package whose project a new one starts from: the host's own template."""
        raise NotImplementedError

    def _new_documents(self) -> list[tuple[str, str | None]]:
        """The document modules the application puts in a new project, with the
        header of each; None keeps the template's module as it is."""
        return []

    def _writes_project(self, project: VBAProject) -> bool:
        """Whether the application writes a project like this one into the file."""
        return True

    def _project_edits(self) -> dict[str, bytes]:
        """The parts a new project changes whether or not it is written."""
        return {}

    def _get_cfb(self) -> CFB:
        if not self._has_project:
            raise NoVBAProjectError(self._no_project_message())
        if self._cfb is not None:
            return self._cfb
        if self._zip is None:
            raise RuntimeError(f"{type(self).__name__} is not open.")
        vba_bin = self._zip.read(self._vba_entry)
        self._cfb = CFB.from_bytes(vba_bin)
        return self._cfb
