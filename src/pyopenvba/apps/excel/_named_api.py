"""Python name CRUD shared by workbook and worksheet views."""
from __future__ import annotations
from dataclasses import dataclass
from pyopenvba._a1 import split_sheet
from pyopenvba.apps.excel._model import DefinedName, Names


@dataclass(frozen=True)
class NamedRange:
    """Detached snapshot of a defined name (which may refer to a formula)."""
    name: str
    refers_to: str
    scope: str | None
    visible: bool
    comment: str


def _snapshot(name: DefinedName) -> NamedRange:
    entry = name.entry
    scope, _ = split_sheet(entry.name)
    return NamedRange(entry.name, entry.refers_to, scope or None, entry.visible, entry.comment)


class NamedRangeAPI:
    @property
    def _named_collection(self) -> Names:
        raise NotImplementedError

    def named_ranges(self) -> list[NamedRange]:
        """List name snapshots; worksheet views list only their local names."""
        return [_snapshot(name) for name in self._named_collection.vba_items() if isinstance(name, DefinedName)]

    def named_range(self, name: str) -> NamedRange:
        return _snapshot(self._lookup_name(name))

    def _lookup_name(self, name: str) -> DefinedName:
        found = self._find_name(name)
        if found is None:
            raise KeyError(name)
        return found

    def _find_name(self, name: str) -> DefinedName | None:
        collection = self._named_collection
        if collection.sheet is None and "!" not in name:
            entry = next((entry for entry in collection.entries if entry.name.casefold() == name.casefold()), None)
            return DefinedName(entry) if entry is not None else None
        return collection.find(name)

    def add_named_range(self, name: str, refers_to: str, *, visible: bool = True, comment: str = "") -> NamedRange:
        """Create or replace a definition in this view's scope."""
        found = self._named_collection.Add(Name=name, RefersTo=refers_to, Visible=visible)
        assert isinstance(found, DefinedName)
        found.vba_set("Comment", comment)
        return _snapshot(found)

    def update_named_range(self, name: str, *, new_name: str | None = None,
                           refers_to: str | None = None, visible: bool | None = None,
                           comment: str | None = None) -> NamedRange:
        """Update a definition; omitted fields remain unchanged."""
        found = self._lookup_name(name)
        for member, value in (("Name", new_name), ("RefersTo", refers_to), ("Visible", visible), ("Comment", comment)):
            if value is not None:
                found.vba_set(member, value)
        return _snapshot(found)

    def remove_named_range(self, name: str) -> bool:
        found = self._find_name(name)
        if found is None:
            return False
        found.Delete()
        return True
