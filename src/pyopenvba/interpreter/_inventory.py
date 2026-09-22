"""What the real applications' classes actually have on them.

This is what lets an unknown member be answered honestly.  ``Worksheet``
really does have a ``PivotTables`` property, so reaching for it is a gap
in pyOpenVBA and says so; ``Worksheet`` has no ``Pivottable`` property,
so reaching for that is VBA's run-time error 438, the same error Excel
would raise.

The tables are generated from the type libraries by
``scripts/build_object_inventory.py`` and committed, so nothing is read
from a machine at run time and the package still has no dependencies.
"""

from __future__ import annotations

from pyopenvba.interpreter._inventory_data import MEMBERS


def _keys(type_name: str, library: str) -> list[str]:
    """Where a type's members sit: under the class, or under its I- or _-named interface.

    Excel's Border, for one, carries its members on IBorder, and the
    coclass itself has none.
    """
    name = type_name.lower()
    if not library:
        return [name]
    prefix = library.lower()
    return [f"{prefix}:{name}", f"{prefix}:i{name}", f"{prefix}:_{name}"]


def type_known(type_name: str, library: str = "") -> bool:
    """Whether the inventory covers this type at all."""
    if any(key in MEMBERS for key in _keys(type_name, library)):
        return True
    return any(key.endswith(f":{type_name.lower()}") for key in MEMBERS)


def members_of(type_name: str, library: str = "") -> frozenset[str]:
    """Every member name the real type has, lowercased."""
    for key in _keys(type_name, library):
        direct = MEMBERS.get(key)
        if direct is not None:
            return direct
    suffix = f":{type_name.lower()}"
    for key in MEMBERS:
        if key.endswith(suffix):
            found = MEMBERS.get(key)
            if found is not None:
                return found
    return frozenset()


def member_exists(type_name: str, member: str, library: str = "") -> bool:
    """Whether the real type has a member of this name."""
    return member.lower() in members_of(type_name, library)
