"""Plan range-copy name imports before changing the destination workbook."""
from __future__ import annotations
from typing import TYPE_CHECKING
from pyopenvba._a1 import quote_sheet
from pyopenvba.formula._parse import split_sheet, tokenize
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import NameEntry, Worksheet


class NameCopyPlan:
    def __init__(self, source: Worksheet, destination: Worksheet, policy: str):
        self.source = source
        self.destination = destination
        self.policy = policy
        self.entries: list[NameEntry] = []
        self.mapping: dict[int, str] = {}
        self.visiting: set[int] = set()

    def rewrite(self, text: str, *, definition: bool = False) -> str:
        for token in reversed(tokenize(text)):
            replacement = token.text
            scope, bare = split_sheet(token.text)
            if token.kind == "ref":
                if not definition and scope:
                    raise VBAUnsupportedError("Cross-workbook qualified references require external-link support")
                if definition:
                    if scope.casefold() != self.source.name.casefold() or "[" in scope:
                        raise VBAUnsupportedError("Copying names referring to other sheets requires external-link support")
                    # Relative definitions depend on the active cell, not the copy offset.
                    if any(not part.startswith("$") or part.count("$") < 2 for part in bare.split(":")):
                        raise VBAUnsupportedError("Copying relative defined names is not implemented")
                    replacement = quote_sheet(self.destination.name) + "!" + bare
            elif token.kind == "name":
                if text[token.at + len(token.text):].lstrip().startswith("(") or token.text.upper() in {"TRUE", "FALSE"}:
                    continue
                found = self.source.book.names_.find(token.text, scope=self.source)
                if found is None:
                    raise VBAUnsupportedError("Copying unresolved defined names is not implemented")
                replacement = self.import_name(found.entry)
                if not scope and not definition:
                    replacement = split_sheet(replacement)[1]
            if replacement != token.text:
                text = text[:token.at] + replacement + text[token.at + len(token.text):]
        return text

    def import_name(self, entry: NameEntry) -> str:
        from pyopenvba.apps.excel._model import NameEntry

        key = id(entry)
        if key in self.visiting:
            raise VBAUnsupportedError("Copying circular defined names is not implemented")
        if key in self.mapping:
            return self.mapping[key]
        scope, bare = split_sheet(entry.name)
        if scope and scope.casefold() != self.source.name.casefold():
            raise VBAUnsupportedError("Copying names scoped to another sheet is not implemented")
        target = self.destination
        existing = target.book.names_.find(bare, scope=target)
        occupied = {split_sheet(one.name)[1].casefold() for one in target.book.names_.entries + self.entries}
        if existing is not None:
            if self.policy == "error":
                raise error(1004, f"Defined name conflict: {bare}")
            if self.policy == "reuse":
                self.mapping[key] = existing.entry.name
                return existing.entry.name
        if (existing is not None or bare.casefold() in {split_sheet(one.name)[1].casefold() for one in self.entries}):
            number = 2
            base = bare
            while bare.casefold() in occupied:
                suffix = f"_{number}"
                bare = base[:255 - len(suffix)] + suffix
                number += 1
        wanted = quote_sheet(target.name) + "!" + bare if scope else bare
        # Reserve the spelling before following aliases so dependencies cannot take it.
        copied = NameEntry(wanted, "", target.book, visible=entry.visible, comment=entry.comment)
        self.entries.append(copied)
        self.mapping[key] = wanted if scope else bare
        self.visiting.add(key)
        copied.refers_to = self.rewrite(entry.refers_to, definition=True)
        self.visiting.remove(key)
        return self.mapping[key]
