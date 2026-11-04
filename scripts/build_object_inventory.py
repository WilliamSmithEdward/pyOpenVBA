"""Generate the object-model inventory and constant tables.

Reads the type-library dumps in the pyVBAReference repository and writes
two committed Python modules, and reads the registered libraries themselves
for what the dumps do not hold:

    src/pyopenvba/interpreter/_inventory_data.py   every class's members
    src/pyopenvba/apps/excel/_constants_data.py    every xl / vb constant
    src/pyopenvba/interpreter/_type_library.json   how each member is called

The inventory is what lets an unknown member be answered honestly: a
member the real Worksheet has is a gap in pyOpenVBA, and one it has not
is VBA's error 438.  The same goes for assigning to a property the model
only reads, where the inventory knows which ones the real class lets a
macro set.  The constants are what let a macro say xlUp.

Usage (the default path is the sibling checkout):

    python scripts/build_object_inventory.py [path\\to\\pyVBAReference]
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any, cast

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_REFERENCE = ROOT.parent / "pyVBAReference"

#: The libraries whose members the interpreter needs to know about.
LIBRARIES = ("excel", "word", "powerpoint", "access", "office", "vba", "msforms", "stdole", "scripting", "vbide")

#: The libraries whose enumerations become named constants.
CONSTANT_LIBRARIES = ("excel", "vba", "office")

INVENTORY_OUT = ROOT / "src" / "pyopenvba" / "interpreter" / "_inventory_data.py"
CONSTANTS_OUT = ROOT / "src" / "pyopenvba" / "interpreter" / "_constants_data.py"
#: VBA's constants read from its own type library (scripts/dump_vba_typelib_constants.py): the dumps leave out
#: ColorConstants, KeyCodeConstants, SystemColorConstants, VbQueryClose and FormShowConstants.
VBA_TYPELIB = ROOT / "scripts" / "vba_typelib_constants.json"
TYPE_LIBRARY_OUT = ROOT / "src" / "pyopenvba" / "interpreter" / "_type_library.json"
#: The libraries whose calls the compiler checks, in the order a project looks in them: registered id and version.
TYPE_LIBRARIES = {
    "excel": ("{00020813-0000-0000-C000-000000000046}", 1, 9),
    "office": ("{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", 2, 8),
}
#: VARTYPE numbers as VBA names them.
_VARIANT_TYPES = {2: "Integer", 3: "Long", 4: "Single", 5: "Double", 6: "Currency", 7: "Date", 8: "String", 9: "Object", 11: "Boolean", 13: "Object", 17: "Byte",
                  19: "Long", 20: "LongLong", 22: "Long", 24: "", 25: ""}


def load_library(reference: Path, library: str) -> list[dict[str, object]]:
    folder = reference / "reference" / library / "json"
    if not folder.is_dir():
        raise SystemExit(f"no such library dump: {folder}")
    out: list[dict[str, object]] = []
    for path in sorted(folder.glob("*.json")):
        out.append(json.loads(path.read_text(encoding="utf-8")))
    return out


def member_names(type_doc: dict[str, object]) -> set[str]:
    names: set[str] = set()
    for group in ("properties", "methods", "events"):
        for entry in type_doc.get(group, []) or []:
            if isinstance(entry, dict) and isinstance(entry.get("name"), str):
                names.add(entry["name"].lower())
    return names


def settable_names(type_doc: dict[str, object]) -> set[str]:
    """The properties a macro can assign to: every one the type library does not mark read-only."""
    return {entry["name"].lower() for entry in type_doc.get("properties", []) or []
            if isinstance(entry, dict) and isinstance(entry.get("name"), str) and entry.get("access") != "read-only"}


def build_inventory(reference: Path) -> tuple[dict[str, str], dict[str, str]]:
    """Each class's members, and the properties among them a macro can set."""
    table: dict[str, str] = {}
    settable: dict[str, str] = {}
    for library in LIBRARIES:
        for type_doc in load_library(reference, library):
            kind = str(type_doc.get("kind", ""))
            if kind not in ("Class", "Interface", "Object", "Module"):
                continue
            name = str(type_doc.get("name", ""))
            names = member_names(type_doc)
            if name and names:
                table[f"{library}:{name.lower()}"] = ",".join(sorted(names))
                put = settable_names(type_doc)
                if put:
                    settable[f"{library}:{name.lower()}"] = ",".join(sorted(put))
    return table, settable


def build_constants(reference: Path) -> tuple[dict[str, dict[str, int | float | str]], set[str]]:
    """One table per library, so a Word host does not answer to xlUp, and the names declared Integer.

    An enumeration's members are Longs to VBA, however small; a module's
    constants have the type they are declared with, which the key codes
    declare Integer (tests/fixtures/vba_semantics).
    """
    tables: dict[str, dict[str, int | float | str]] = {}
    integers: set[str] = set()
    for library in CONSTANT_LIBRARIES:
        table: dict[str, int | float | str] = {}
        for type_doc in load_library(reference, library):
            for entry in type_doc.get("constants", []) or []:
                if not isinstance(entry, dict):
                    continue
                name = entry.get("name")
                value = entry.get("value")
                if not isinstance(name, str) or value is None or isinstance(value, bool):
                    continue
                if name in table and table[name] != value:
                    continue
                if isinstance(value, (int, float, str)):
                    table[name] = value
                    if entry.get("type") == "Integer":
                        integers.add(name.lower())
        tables[library] = table
    for group in json.loads(VBA_TYPELIB.read_text(encoding="utf-8"))["groups"].values():
        for name, entry in group.items():
            if name not in tables["vba"]:
                tables["vba"][name] = entry["value"]
                if entry["type"] == "Integer":
                    integers.add(name.lower())
    return tables, integers


def write_inventory(table: dict[str, str], settable: dict[str, str]) -> None:
    lines = [
        '"""Every member each Office class really has, generated from the type libraries.',
        "",
        "Written by scripts/build_object_inventory.py.  Do not edit by hand.",
        "",
        "The names are lowercased and joined with commas, one string per",
        "class, because a few hundred frozensets cost more to import than",
        "they save and the lookup splits on demand.",
        '"""',
        "",
        "from __future__ import annotations",
        "",
        "from functools import lru_cache",
        "",
        'VERSION = "type libraries via pyVBAReference"',
        "",
        "_NAMES: dict[str, str] = {",
    ]
    for key in sorted(table):
        lines.append(f"    {key!r}: {table[key]!r},")
    lines.extend(
        [
            "}",
            "",
            "",
            "@lru_cache(maxsize=None)",
            "def _split(key: str) -> frozenset[str]:",
            "    return frozenset(_NAMES[key].split(','))",
            "",
            "",
            "class _Members:",
            '    """The table, splitting each class\'s names the first time it is asked."""',
            "",
            "    def __contains__(self, key: object) -> bool:",
            "        return key in _NAMES",
            "",
            "    def __iter__(self):",
            "        return iter(_NAMES)",
            "",
            "    def __len__(self) -> int:",
            "        return len(_NAMES)",
            "",
            "    def keys(self):",
            "        return _NAMES.keys()",
            "",
            "    def items(self):",
            "        for key in _NAMES:",
            "            yield key, _split(key)",
            "",
            "    def get(self, key: str, default: frozenset[str] | None = None) -> frozenset[str] | None:",
            "        return _split(key) if key in _NAMES else default",
            "",
            "",
            "MEMBERS = _Members()",
            "",
            "#: The properties of each class a macro can assign to, joined the same way.",
            "_SETTABLE: dict[str, str] = {",
            *[f"    {key!r}: {settable[key]!r}," for key in sorted(settable)],
            "}",
            "",
            "",
            "@lru_cache(maxsize=None)",
            "def settable(key: str) -> frozenset[str]:",
            '    """The properties of the class under ``key`` a macro can assign to."""',
            "    return frozenset(_SETTABLE[key].split(',')) if key in _SETTABLE else frozenset()",
            "",
        ]
    )
    INVENTORY_OUT.write_text("\n".join(lines), encoding="utf-8")


def write_constants(tables: dict[str, dict[str, int | float | str]], integers: set[str]) -> None:
    lines = [
        '"""Every vb, xl and mso constant, generated from the type libraries.',
        "",
        "Written by scripts/build_object_inventory.py.  Do not edit by hand.",
        "",
        "One table per library: a project sees its host's constants and the",
        "VBA ones, never another host's.",
        '"""',
        "",
        "from __future__ import annotations",
        "",
        "from typing import Final",
        "",
        "Constants = dict[str, int | float | str]",
        "",
    ]
    names = {"vba": "VBA_CONSTANTS", "excel": "EXCEL_CONSTANTS", "office": "OFFICE_CONSTANTS"}
    for library, table in tables.items():
        lines.append(f"{names.get(library, library.upper() + '_CONSTANTS')}: Final[Constants] = {{")
        for key in sorted(table, key=str.lower):
            lines.append(f"    {key!r}: {table[key]!r},")
        lines.extend(["}", ""])
    lines.extend(["#: The whole-number constants declared Integer, lowercased; every other one is a Long.",
                  "INTEGER_CONSTANTS: Final[frozenset[str]] = frozenset({"])
    lines.extend(f"    {name!r}," for name in sorted(integers))
    lines.extend(["})", ""])
    CONSTANTS_OUT.write_text("\n".join(lines), encoding="utf-8")


def _type_text(info: Any, described: Any) -> str:
    """A return or parameter type as VBA names it: an enumeration is a Long, nothing returned is ''."""
    if isinstance(described, tuple) and isinstance(described[0], tuple):
        return _type_text(info, described[0])
    kind = described[0] if isinstance(described, tuple) else described
    if kind in (26, 27):  # A pointer to, or a safe array of, what it holds.
        return _type_text(info, described[1])
    if kind == 29:
        held = info.GetRefTypeInfo(described[1])
        return "Long" if held.GetTypeAttr().typekind == 0 else held.GetDocumentation(-1)[0]
    return _VARIANT_TYPES.get(kind, "Variant")


def _library_table(library: Any) -> dict[str, object]:
    classes: dict[str, object] = {}
    aliases: dict[str, str] = {}
    for index in range(library.GetTypeInfoCount()):
        info = library.GetTypeInfo(index)
        attr = info.GetTypeAttr()
        title = info.GetDocumentation(-1)[0]
        if attr.typekind == 5:  # A coclass answers to its default interface, not its events.
            for place in range(attr.cImplTypes):
                if info.GetImplTypeFlags(place) == 1:
                    aliases[title.lower()] = info.GetRefTypeInfo(info.GetRefTypeOfImplType(place)).GetDocumentation(-1)[0].lower()
            continue
        if attr.typekind != 4:
            continue
        members: dict[str, list[object]] = {}
        default = ""
        for place in range(attr.cFuncs):
            described = info.GetFuncDesc(place)
            names = info.GetNames(described.memid)
            # A parameter is optional by its flags (optional, has a default) or by being one of the last cParamsOpt.
            count = len(described.args)
            plain = count - max(described.cParamsOpt, 0)
            required = sum(1 for at, argument in enumerate(described.args) if at < plain and not argument[1] & 0x30)
            parameters = [f"{names[at + 1] if at + 1 < len(names) else ''}:{_type_text(info, argument[0])}" for at, argument in enumerate(described.args)]
            members.setdefault(names[0].lower(), []).append(
                [described.invkind, required, described.cParamsOpt == -1, _type_text(info, described.rettype), parameters])
            if described.memid == 0:
                default = names[0].lower()
        classes[title.lower()] = {"name": title, "dual": bool(attr.wTypeFlags & 0x40), "nonextensible": bool(attr.wTypeFlags & 0x80), "default": default, "members": members}
    return {"classes": classes, "aliases": aliases}


def write_type_library() -> None:
    """How each class's members are called, read from the registered libraries: what the VBA compiler reads.

    The dumps above hold names and help text. This needs pywin32 and Office on the machine, so a machine without
    them keeps the committed table. One class is one line, so that a rebuild's difference can be read.
    """
    try:
        import pythoncom
        import win32api
    except ImportError:
        print(f"no pywin32: {TYPE_LIBRARY_OUT.name} left as it is")
        return
    versions: dict[str, str] = {}
    lines: list[str] = []
    build = ""
    for name, (guid, major, minor) in TYPE_LIBRARIES.items():
        library = pythoncom.LoadRegTypeLib(guid, major, minor, 0)
        attributes = library.GetLibAttr()
        versions[name] = f"{attributes[3]}.{attributes[4]}"
        table = _library_table(library)
        if name == "excel":
            path = pythoncom.QueryPathOfRegTypeLib(guid, major, minor, 0).rstrip("\0")
            found = win32api.GetFileVersionInfo(path, "\\")
            build = f"{found['FileVersionMS'] >> 16}.{found['FileVersionMS'] & 0xFFFF}.{found['FileVersionLS'] >> 16}.{found['FileVersionLS'] & 0xFFFF}"
        classes = cast("dict[str, object]", table["classes"])
        lines.append(f'"{name}": {{"aliases": {json.dumps(table["aliases"], sort_keys=True)}, "classes": {{')
        lines.append(",\n".join(f"{json.dumps(key)}: {json.dumps(classes[key], separators=(',', ':'))}" for key in sorted(classes)))
        lines.append("}}")
    head = f'{{"excel_build": {json.dumps(build)}, "versions": {json.dumps(versions, sort_keys=True)}, "libraries": {{\n'
    body = "\n".join(line + ("," if line == "}}" else "") for line in lines).rstrip(",")
    TYPE_LIBRARY_OUT.write_text(head + body + "\n}}\n", encoding="utf-8", newline="\n")
    print(f"{len(TYPE_LIBRARIES)} libraries, Excel {build} -> {TYPE_LIBRARY_OUT.relative_to(ROOT)} ({TYPE_LIBRARY_OUT.stat().st_size} bytes)")


def main() -> int:
    write_type_library()
    reference = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_REFERENCE
    if not reference.is_dir():
        raise SystemExit(f"pyVBAReference is not at {reference}")
    inventory, settable = build_inventory(reference)
    constants, integers = build_constants(reference)
    write_inventory(inventory, settable)
    write_constants(constants, integers)
    print(f"{len(inventory)} classes -> {INVENTORY_OUT.relative_to(ROOT)} ({INVENTORY_OUT.stat().st_size} bytes)")
    print(f"{len(constants)} constants -> {CONSTANTS_OUT.relative_to(ROOT)} ({CONSTANTS_OUT.stat().st_size} bytes)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
