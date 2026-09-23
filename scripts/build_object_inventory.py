"""Generate the object-model inventory and constant tables.

Reads the type-library dumps in the pyVBAReference repository and writes
two committed Python modules:

    src/pyopenvba/interpreter/_inventory_data.py   every class's members
    src/pyopenvba/apps/excel/_constants_data.py    every xl / vb constant

The inventory is what lets an unknown member be answered honestly: a
member the real Worksheet has is a gap in pyOpenVBA, and one it has not
is VBA's error 438.  The constants are what let a macro say xlUp.

Usage (the default path is the sibling checkout):

    python scripts/build_object_inventory.py [path\\to\\pyVBAReference]
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_REFERENCE = ROOT.parent / "pyVBAReference"

#: The libraries whose members the interpreter needs to know about.
LIBRARIES = ("excel", "word", "powerpoint", "access", "office", "vba", "msforms", "stdole", "scripting")

#: The libraries whose enumerations become named constants.
CONSTANT_LIBRARIES = ("excel", "vba", "office")

INVENTORY_OUT = ROOT / "src" / "pyopenvba" / "interpreter" / "_inventory_data.py"
CONSTANTS_OUT = ROOT / "src" / "pyopenvba" / "interpreter" / "_constants_data.py"
#: VBA's constants read from its own type library (scripts/dump_vba_typelib_constants.py): the dumps leave out
#: ColorConstants, KeyCodeConstants, SystemColorConstants, VbQueryClose and FormShowConstants.
VBA_TYPELIB = ROOT / "scripts" / "vba_typelib_constants.json"


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


def build_inventory(reference: Path) -> dict[str, str]:
    table: dict[str, str] = {}
    for library in LIBRARIES:
        for type_doc in load_library(reference, library):
            kind = str(type_doc.get("kind", ""))
            if kind not in ("Class", "Interface", "Object", "Module"):
                continue
            name = str(type_doc.get("name", ""))
            names = member_names(type_doc)
            if name and names:
                table[f"{library}:{name.lower()}"] = ",".join(sorted(names))
    return table


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


def write_inventory(table: dict[str, str]) -> None:
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


def main() -> int:
    reference = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_REFERENCE
    if not reference.is_dir():
        raise SystemExit(f"pyVBAReference is not at {reference}")
    inventory = build_inventory(reference)
    constants, integers = build_constants(reference)
    write_inventory(inventory)
    write_constants(constants, integers)
    print(f"{len(inventory)} classes -> {INVENTORY_OUT.relative_to(ROOT)} ({INVENTORY_OUT.stat().st_size} bytes)")
    print(f"{len(constants)} constants -> {CONSTANTS_OUT.relative_to(ROOT)} ({CONSTANTS_OUT.stat().st_size} bytes)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
