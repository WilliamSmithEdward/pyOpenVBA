"""The Office-host shape of the Access API.

``ExcelFile``, ``WordFile`` and ``PowerPointFile`` expose a VBA project
through ``vba_project()`` and a form's design through ``forms()``; these
classes give ``AccessDatabase`` the same surface, so code written against
one host reads the same against another.  Each call goes straight to the
database's own writers, which change the in-memory pages; nothing reaches
disk until ``save()``.
"""

from __future__ import annotations

import struct
from pathlib import Path
from typing import TYPE_CHECKING

from pyopenvba.access._designs import (
    GUID_LENGTH,
    TEXT_VALUE_TYPES,
    AccessDesign,
    DesignObject,
    DesignRecord,
)
from pyopenvba.access._vba import VBAModule
from pyopenvba.access_read import AccessError
from pyopenvba.vba import VBAModuleKind

if TYPE_CHECKING:
    from pyopenvba.access.database import AccessDatabase, Reference

#: How the host-style module kinds map onto the database's own.
MODULE_KINDS: dict[object, str] = {
    VBAModuleKind.standard: "module",
    VBAModuleKind.other: "class",
    "standard": "module",
    "module": "module",
    "other": "class",
    "class": "class",
}

_BAS_EXT = ".bas"
_CLS_EXT = ".cls"
_SOURCE_EXTS = frozenset({_BAS_EXT, _CLS_EXT})
_CRLF = "\r\n"


def module_kind(kind: VBAModuleKind | str) -> str:
    """``"module"`` or ``"class"`` for a host-style or database-style kind."""
    try:
        return MODULE_KINDS[kind]
    except KeyError:
        raise AccessError(f"kind must be VBAModuleKind.standard or .other (or 'module' / 'class'), not {kind!r}") from None


def export_body(text: str) -> str:
    """The body of a module file as the database stores it: a VBE export's
    ``VERSION ... END`` preamble and its ``Attribute`` lines removed, line
    endings CRLF."""
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    if lines and lines[0].upper().startswith("VERSION "):
        end = next((i for i, line in enumerate(lines) if line.strip().upper() == "END"), None)
        lines = lines[end + 1 :] if end is not None else lines[1:]
    while lines and lines[0].startswith("Attribute VB_"):
        lines = lines[1:]
    return _CRLF.join(lines)


def crlf(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", _CRLF)


class AccessVBAProject:
    """The database's VBA project, shaped like a host's ``vba_project()``."""

    def __init__(self, database: AccessDatabase) -> None:
        self._db = database

    @property
    def modules(self) -> list[VBAModule]:
        return self._db.modules()

    def module_names(self) -> list[str]:
        return [module.name for module in self._db.modules()]

    @property
    def references(self) -> list[Reference]:
        return self._db.references()

    def add_module(
        self,
        name: str,
        source: str = "Option Compare Database",
        *,
        kind: VBAModuleKind | str = VBAModuleKind.standard,
    ) -> VBAModule:
        """Add a module; ``kind`` as the other hosts take it
        (``VBAModuleKind.standard`` or ``VBAModuleKind.other`` for a class)."""
        return self._db.create_module(name, source, kind=module_kind(kind))

    def rename_module(self, old_name: str, new_name: str) -> VBAModule:
        self._db.rename_module(old_name, new_name)
        return self._db.module(new_name)

    def delete_module(self, name: str) -> None:
        self._db.delete_module(name)


def decode_record(record: DesignRecord) -> object:
    """A record's value as the property means it: text, a GUID or raw
    bytes, a float, or a number."""
    if record.value_type in TEXT_VALUE_TYPES:
        text = record.text()
        return text if text is not None else record.value
    if record.value_type in (9, 11) or (record.value_type == 9 and len(record.value) == GUID_LENGTH):
        return bytes(record.value)
    if record.value_type == 6 and len(record.value) == 4:
        return struct.unpack("<f", record.value)[0]
    if record.value_type == 8 and len(record.value) == 8:
        return struct.unpack("<d", record.value)[0]
    if 1 <= len(record.value) <= 8:
        return int.from_bytes(record.value, "little")
    return bytes(record.value)


class AccessControl:
    """A section or control of a form or report, with the design object
    it was read from."""

    def __init__(self, form: AccessForm, obj: DesignObject) -> None:
        self._form = form
        self._object = obj

    # -- what the design object says --------------------------------------

    @property
    def object(self) -> DesignObject:
        return self._object

    @property
    def name(self) -> str:
        return self._object.name or ""

    @property
    def type_name(self) -> str | None:
        return self._object.type_name

    @property
    def kind(self) -> str:
        """The control's type name, as the other hosts call it."""
        return self._object.type_name or ""

    @property
    def marker(self) -> int | None:
        return self._object.marker

    @property
    def type(self) -> int | None:
        return self._object.type

    @property
    def code(self) -> int | None:
        return self._object.code

    @property
    def records(self) -> tuple[DesignRecord, ...]:
        return self._object.records

    @property
    def is_section(self) -> bool:
        return self._object.is_section

    @property
    def tab_index(self) -> int | None:
        value = self.properties().get("TabIndex")
        return value if isinstance(value, int) else None

    def property_value(self, code: int) -> bytes | None:
        return self._object.property_value(code)

    def properties(self) -> dict[str, object]:
        """Every property the control carries, by name where the code is
        named and as ``Unidentified<code>`` where it is not."""
        out: dict[str, object] = {}
        for record in self._object.records:
            out[record.name or f"Unidentified{record.code}"] = decode_record(record)
        return out

    def get(self, name: str) -> object:
        return self.properties().get(name)

    def set_property(self, name: str, value: object) -> None:
        """Change one property, at the slot the control type's schema
        gives it."""
        self._form._set_control_property(self.name, name, value)  # pyright: ignore[reportPrivateUsage]
        self._object = self._form.control(self.name).object

    def __repr__(self) -> str:
        return f"AccessControl({self.name!r}, {self.kind!r})"


class AccessForm:
    """A form or report, shaped like a host's ``VBAForm``: its controls
    and sections read and edited in place, code put behind it."""

    def __init__(self, database: AccessDatabase, design: AccessDesign) -> None:
        self._db = database
        self._design = design

    def _refresh(self) -> None:
        self._design = self._db._design(self.kind, self.name)  # pyright: ignore[reportPrivateUsage]

    # -- the design as read -------------------------------------------------

    @property
    def design(self) -> AccessDesign:
        return self._design

    @property
    def name(self) -> str:
        return self._design.name

    @property
    def kind(self) -> str:
        """``"form"`` or ``"report"``."""
        return self._design.kind

    @property
    def objects(self) -> tuple[DesignObject, ...]:
        return self._design.objects

    @property
    def root(self) -> DesignObject:
        return self._design.root

    @property
    def sections(self) -> tuple[AccessControl, ...]:
        return tuple(AccessControl(self, obj) for obj in self._design.sections)

    @property
    def controls(self) -> tuple[AccessControl, ...]:
        """The named controls, in design order."""
        return tuple(AccessControl(self, obj) for obj in self._design.controls)

    def walk(self) -> list[AccessControl]:
        """Every control, in the order the design lists them; a page's
        controls follow the page and a tab control's pages follow it."""
        return list(self.controls)

    def control(self, name: str) -> AccessControl:
        """One control or section by name."""
        for obj in self._design.objects[1:]:
            if obj.name is not None and obj.name.lower() == name.lower():
                return AccessControl(self, obj)
        raise AccessError(f"the {self.kind} {self.name!r} has no control named {name!r}")

    # -- the design's own properties --------------------------------------

    def properties(self) -> dict[str, object]:
        return AccessControl(self, self._design.root).properties()

    def get(self, name: str) -> object:
        return self.properties().get(name)

    def set_property(self, name: str, value: object) -> None:
        """Change one property of the form or report itself."""
        self._db.set_design_property(self.name, name, value, kind=self.kind)
        self._refresh()

    def _set_control_property(self, control: str, name: str, value: object) -> None:
        self._db.set_control_property(self.name, control, name, value, kind=self.kind)
        self._refresh()

    # -- editing ------------------------------------------------------------

    def add_control(
        self,
        control_type: str,
        name: str,
        *,
        section: str = "Detail",
        container: str | None = None,
        parent: str | None = None,
        left: int = 0,
        top: int = 0,
        width: int = 1440,
        height: int = 240,
        caption: str | None = None,
    ) -> AccessControl:
        """Put a control on the design; sizes in twips, as Access keeps
        them.  ``container`` (the other hosts' word) and ``parent`` both
        name the control that will hold it, which only a tab control does."""
        holder = parent if parent is not None else container
        self._db.add_control(
            self.name,
            control_type,
            name,
            kind=self.kind,
            section=section,
            parent=holder,
            left=left,
            top=top,
            width=width,
            height=height,
            caption=caption,
        )
        self._refresh()
        return self.control(name)

    def remove_control(self, name: str) -> None:
        """Take a control off the design, with anything it holds."""
        self._db.remove_control(self.name, name, kind=self.kind)
        self._refresh()

    def set_code(self, code: str) -> VBAModule:
        """Put VBA behind the form or report (``Form_<name>`` or
        ``Report_<name>``), creating the module when there is none."""
        return self._db.set_design_code(self.name, code, kind=self.kind)

    def __repr__(self) -> str:
        return f"AccessForm({self.name!r}, {self.kind!r}, {len(self._design.controls)} controls)"


def pull_modules(database: AccessDatabase, dest_dir: str | Path, *, encoding: str, overwrite: bool) -> list[Path]:
    """Every module's body as a ``.bas`` (standard) or ``.cls`` (class)
    file, CRLF line endings, as the other hosts export theirs."""
    out_dir = Path(dest_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    modules = database.modules()
    targets = [(module, out_dir / f"{module.name}{_BAS_EXT if module.kind == 'module' else _CLS_EXT}") for module in modules]
    if not overwrite:
        for _module, target in targets:
            if target.exists():
                raise FileExistsError(f"Refusing to overwrite {target} (overwrite=False).")
    written: list[Path] = []
    for module, target in targets:
        target.write_bytes(crlf(module.source).encode(encoding, errors="replace"))
        written.append(target)
    return written


def push_modules(database: AccessDatabase, src_dir: str | Path, *, encoding: str, strict: bool) -> list[str]:
    """Replace each module's source from the ``.bas`` / ``.cls`` file of
    its name in ``src_dir``; a file that matches no module is skipped, or
    refused with ``strict``.  Returns the modules updated."""
    src = Path(src_dir)
    if not src.is_dir():
        raise NotADirectoryError(f"Not a directory: {src}")
    by_name = {module.name.casefold(): module for module in database.modules()}
    updated: list[str] = []
    for child in sorted(src.iterdir()):
        if not child.is_file() or child.suffix.lower() not in _SOURCE_EXTS:
            continue
        module = by_name.get(child.stem.casefold())
        if module is None:
            if strict:
                raise KeyError(f"No module matches file {child.name!r} (known: {sorted(m.name for m in by_name.values())}).")
            continue
        text = child.read_bytes().decode(encoding, errors="replace")
        database.set_module_source(module.name, export_body(text))
        updated.append(module.name)
    return updated
