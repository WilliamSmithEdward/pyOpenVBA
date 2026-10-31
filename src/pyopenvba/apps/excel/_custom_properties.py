"""Worksheet.CustomProperties: a sheet's named text values, as Excel keeps and saves them
(tests/fixtures/custom_properties.json, scripts/measure_custom_properties.py).

- Add(Name, Value) keeps the value as text, a number or date as CStr writes it and True as -1; a range gives its
  value. The properties are kept in the order of their names; a name there already, case aside, is error 9.
  An empty value is error 7, as is a range with nothing in it; Empty, Null or an array, 1004; leaving either
  argument out, 450.
- Items are found by number only: a name is a type mismatch, 0 error 1004, past the end error 9.
- A copy of the sheet has none.
- The sheet part lists them as customProperties' customPr name and r:id, each value in its own part,
  xl/customPropertyN.bin, UTF-16 text with no byte-order mark, numbered across the workbook in sheet order.
  Written only for a sheet with no other relationships, in a workbook without macros.
"""
from __future__ import annotations

import posixpath
import re
from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyopenvba._relationships import in_office_order
from pyopenvba._xml import attributes, escape
from pyopenvba.apps.excel._model import ExcelObject, Range
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import (
    EMPTY,
    MISSING,
    NULL,
    VBAArray,
    VBAInt,
    error,
    to_integer,
    to_text,
)

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Workbook, Worksheet
    from pyopenvba.powerquery._opc import OpcFile

TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperty"
CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.customProperty"
_ELEMENT = re.compile(r"<customProperties>.*?</customProperties>|<customProperties/>", re.DOTALL)
_RELATIONSHIP = re.compile(r"<Relationship\b[^>]*?(?:/>|>\s*</Relationship>)")
#: What follows customProperties in a worksheet part (ECMA-376 Part 1, 18.3.1.99).
_AFTER = re.compile(r"<(?:cellWatches|ignoredErrors|smartTags|drawing|legacyDrawing|legacyDrawingHF|drawingHF|"
                    r"picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b|</worksheet>")
_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'


@dataclass
class Property:
    name: str
    value: str


def properties(sheet: Worksheet) -> list[Property]:
    """The sheet's properties, read from its part the first time they are wanted."""
    if sheet.custom_properties is None:
        sheet.custom_properties = _read(sheet)
    return sheet.custom_properties


def _rels_part(part: str) -> str:
    folder, _, name = part.rpartition("/")
    return f"{folder}/_rels/{name}.rels"


def _read(sheet: Worksheet) -> list[Property]:
    package = sheet.book.package
    if package is None or not sheet.part_name or not package.has(sheet.part_name):
        return []
    xml = package.read(sheet.part_name).decode("utf-8", errors="replace")
    found = _ELEMENT.search(xml)
    if found is None:
        return []
    rels_part = _rels_part(sheet.part_name)
    rels = package.read(rels_part).decode("utf-8", errors="replace") if package.has(rels_part) else ""
    targets = {fields.get("Id", ""): fields.get("Target", "")
               for fields in (attributes(element) for element in _RELATIONSHIP.findall(rels))}
    folder = sheet.part_name.rpartition("/")[0]
    out: list[Property] = []
    for element in re.findall(r"<customPr\b[^>]*/?>", found.group(0)):
        fields = attributes(element)
        target = targets.get(fields.get("r:id", ""), "")
        part = posixpath.normpath(posixpath.join(folder, target)).lstrip("/") if target else ""
        value = package.read(part).decode("utf-16-le", errors="replace") if part and package.has(part) else ""
        out.append(Property(fields.get("name", ""), value))
    return out


def _order(name: str) -> bytes:
    from pyopenvba.apps.excel._names import collation

    return collation(name)


def _value_text(value: object) -> str:
    """What a property keeps of a value: text, a range's value, True as -1."""
    if isinstance(value, Range):
        value = value.vba_get("Value")
        if value is EMPTY:
            value = ""
    elif value is EMPTY or value is NULL or isinstance(value, VBAArray):
        raise error(1004)
    if isinstance(value, bool):
        text = "-1" if value else "0"
    else:
        text = to_text(value)
    if not text:
        raise error(7)
    return text


def _changed(sheet: Worksheet) -> None:
    sheet.book.custom_properties_changed = True


class CustomProperties(VBACollection, ExcelObject):
    """A sheet's custom properties."""

    vba_type_name = "CustomProperties"

    def __init__(self, sheet: Worksheet) -> None:
        self.sheet = sheet

    def vba_items(self) -> list[object]:
        return [CustomProperty(self.sheet, one) for one in properties(self.sheet)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        position = int(to_integer(index, "Long"))
        if position == 0:
            raise error(1004)
        return super().vba_lookup(position, items)

    @method
    def Add(self, Name: object = MISSING, Value: object = MISSING) -> object:
        if Name is MISSING or Value is MISSING:
            raise error(450)
        name = to_text(Name)
        if len(name) > 255:
            raise VBAUnsupportedError("A custom property name longer than 255 characters is not implemented")
        text = _value_text(Value)
        kept = properties(self.sheet)
        if any(one.name.lower() == name.lower() for one in kept):
            raise error(9)
        added = Property(name, text)
        kept.append(added)
        kept.sort(key=lambda one: _order(one.name))
        _changed(self.sheet)
        return CustomProperty(self.sheet, added)

    @member
    def Parent(self) -> object:
        return self.sheet

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @member
    def Creator(self) -> object:
        return VBAInt(1480803660, "Long")


class CustomProperty(ExcelObject):
    """One custom property of a sheet."""

    vba_type_name = "CustomProperty"

    def __init__(self, sheet: Worksheet, kept: Property) -> None:
        self.sheet = sheet
        self.kept = kept

    @member
    def Name(self) -> object:
        return self.kept.name

    @setter("Name")
    def _set_name(self, value: object) -> None:
        name = to_text(value)
        kept = properties(self.sheet)
        others = [one for one in kept if one is not self.kept]
        if any(one.name.lower() == name.lower() for one in others):
            raise VBAUnsupportedError("Renaming a custom property to a name another has is not implemented")
        before = [one.name for one in kept]
        self.kept.name = name
        if sorted(kept, key=lambda one: _order(one.name)) != kept:
            self.kept.name = before[kept.index(self.kept)]
            raise VBAUnsupportedError("Renaming a custom property out of its place among the others is not "
                                      "implemented")
        _changed(self.sheet)

    @member(default=True)
    def Value(self) -> object:
        return self.kept.value

    @setter("Value")
    def _set_value(self, value: object) -> None:
        self.kept.value = _value_text(value)
        _changed(self.sheet)

    @method
    def Delete(self) -> object:
        kept = properties(self.sheet)
        if self.kept in kept:
            kept.remove(self.kept)
            _changed(self.sheet)
        return EMPTY

    @member
    def Parent(self) -> object:
        return self.sheet

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @member
    def Creator(self) -> object:
        return VBAInt(1480803660, "Long")


def copied(source: Worksheet, copy: Worksheet) -> None:
    """A copy of a sheet has no custom properties; its part, taken from the sheet's, is written without them."""
    copy.custom_properties = []
    if properties(source):
        _changed(copy)


def write(book: Workbook, package: OpcFile) -> None:
    """Every sheet's custom properties put back into the package."""
    if not book.custom_properties_changed:
        return
    wanted = {id(sheet): properties(sheet) for sheet in book.sheets_}
    if any(wanted.values()) and book.format_code in (52, 53, 55):
        raise VBAUnsupportedError("Saving custom properties in a workbook with macros is not implemented")
    number = 0
    for sheet in book.sheets_:
        if not sheet.part_name or not package.has(sheet.part_name):
            continue
        rels_part = _rels_part(sheet.part_name)
        rels = package.read(rels_part).decode("utf-8", errors="replace") if package.has(rels_part) else ""
        folder = sheet.part_name.rpartition("/")[0]
        others: list[str] = []
        for element in _RELATIONSHIP.findall(rels):
            fields = attributes(element)
            if fields.get("Type") == TYPE:
                old = posixpath.normpath(posixpath.join(folder, fields.get("Target", ""))).lstrip("/")
                if package.has(old):
                    package.remove(old)
            else:
                others.append(element)
        kept = wanted[id(sheet)]
        if kept and others:
            raise VBAUnsupportedError("Saving custom properties on a sheet with other relationships is not "
                                      "implemented")
        xml = package.read(sheet.part_name).decode("utf-8", errors="replace")
        xml = _ELEMENT.sub("", xml)
        elements: list[str] = []
        listed: list[str] = []
        for at, one in enumerate(kept, start=1):
            number += 1
            package.write(f"xl/customProperty{number}.bin", one.value.encode("utf-16-le"))
            elements.append(f'<Relationship Id="rId{at}" Type="{TYPE}" Target="../customProperty{number}.bin"/>')
            listed.append(f'<customPr name="{escape(one.name)}" r:id="rId{at}"/>')
        if listed:
            following = _AFTER.search(xml, max(xml.find("</sheetData>"), 0))
            assert following is not None
            xml = xml[:following.start()] + f"<customProperties>{''.join(listed)}</customProperties>" \
                + xml[following.start():]
        package.write(sheet.part_name, xml.encode("utf-8"))
        if elements:
            package.write(rels_part, in_office_order(_HEAD + "".join(elements) + "</Relationships>").encode("utf-8"))
        elif package.has(rels_part) and not others:
            package.remove(rels_part)
        elif package.has(rels_part):
            package.write(rels_part, (_HEAD + "".join(others) + "</Relationships>").encode("utf-8"))
    _content_type(package, number > 0)
    book.custom_properties_changed = False


def _content_type(package: OpcFile, wanted: bool) -> None:
    """The bin Default for the property parts, there while any is."""
    raw = package.read("[Content_Types].xml").decode("utf-8")
    element = f'<Default Extension="bin" ContentType="{CONTENT_TYPE}"/>'
    present = re.search(r'<Default Extension="bin" ContentType="([^"]*)"/>', raw)
    if wanted and present is None:
        # Excel lists the Defaults by extension.
        defaults = list(re.finditer(r'<Default Extension="([^"]*)"[^>]*/>', raw))
        later = next((one for one in defaults if one.group(1).lower() > "bin"), None)
        at = later.start() if later is not None else defaults[-1].end() if defaults else \
            raw.find(">", raw.find("<Types")) + 1
        raw = raw[:at] + element + raw[at:]
    elif wanted and present is not None and present.group(1) != CONTENT_TYPE:
        raise VBAUnsupportedError("Saving custom properties beside other .bin parts is not implemented")
    elif not wanted and present is not None and present.group(1) == CONTENT_TYPE:
        raw = raw.replace(present.group(0), "")
    package.write("[Content_Types].xml", raw.encode("utf-8"))
