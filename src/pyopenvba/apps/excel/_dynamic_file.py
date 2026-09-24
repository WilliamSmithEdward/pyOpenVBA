"""Dynamic-array formulas in a file: the metadata that marks them, and the rich values behind their errors.

Measured in live Excel (scripts/measure_dynamic_arrays.py,
tests/fixtures/dynamic_arrays.json, its file parts). A dynamic-array
formula's cell is cm="1", its formula an array formula over the block it
spilled into, <f t="array" ref="E1:E3">, and xl/metadata.xml's first cell
metadata says XLDAPR fDynamic="1"; the cells it spilled into hold their
values and nothing else. An error only a dynamic array has is #VALUE! in
the cell, with vm="n": the rich value behind value metadata n in
xl/richData says what it is, #SPILL! (errorType 8) with why and how far
the answer reached, or #CALC! (13) for an empty array. A formula that is
#SPILL! is aca="1" ca="1" as well.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from pyopenvba.powerquery._opc import OpcFile
from pyopenvba.powerquery._sheets import add_content_type, add_relationship

_METADATA = "xl/metadata.xml"
_RICH_VALUES = "xl/richData/rdrichvalue.xml"
_STRUCTURES = "xl/richData/rdrichvaluestructure.xml"
_TYPES = "xl/richData/rdRichValueTypes.xml"
_RELATIONSHIPS = "xl/_rels/workbook.xml.rels"

#: The error types a rich value gives the errors only a dynamic array has.
_ERRORS = {8: "#SPILL!", 13: "#CALC!"}
_TYPE_OF = {name: number for number, name in _ERRORS.items()}
#: #CALC! for an empty array, FILTER's with nothing kept.
EMPTY_ARRAY = 3

_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_RICH = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
_DYNAMIC = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"
_FLAGS = ('minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" '
          'rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"')
_DAPR = (f'<metadataType name="XLDAPR" {_FLAGS} cellMeta="1"/>')
_RICH_TYPE = f'<metadataType name="XLRICHVALUE" {_FLAGS}/>'
_DAPR_BLOCK = ('<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
               '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>')
_RICH_BLOCK = '<bk><extLst><ext uri="{{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}}"><xlrd:rvb i="{0}"/></ext></extLst></bk>'
_VALUE_TYPES = (
    f'{_HEAD}<rvTypesInfo xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x" '
    f'xmlns:x="{_MAIN}"><global><keyFlags><key name="_Self"><flag name="ExcludeFromFile" value="1"/>'
    '<flag name="ExcludeFromCalcComparison" value="1"/></key>'
    + "".join(f'<key name="{key}"><flag name="ExcludeFromCalcComparison" value="1"/></key>'
              for key in ("_DisplayString", "_Flags", "_Format", "_SubLabel", "_Attribution", "_Icon", "_Display",
                          "_CanonicalPropertyNames", "_ClassificationId"))
    + "</keyFlags></global></rvTypesInfo>")
#: The keys of the two shapes of error a rich value takes: #SPILL!'s, with how far the answer reached, and #CALC!'s.
_SPILL_KEYS = ("colOffset", "errorType", "rwOffset", "subType")
_CALC_KEYS = ("errorType", "subType")


@dataclass(frozen=True)
class RichError:
    """An error only a dynamic array has, as a rich value keeps it: which, why, and for #SPILL! how many rows and
    columns past the formula's cell the answer reached."""

    error: str
    why: int
    rows: int = 0
    columns: int = 0


@dataclass
class Metadata:
    """What a file's metadata says of its cells: the cell metadata that marks a dynamic-array formula, and the error
    behind each value metadata, by the index a cell's cm and vm give."""

    dynamic: set[int] = field(default_factory=lambda: set())
    errors: dict[int, RichError] = field(default_factory=lambda: {})


def read(package: OpcFile) -> Metadata:
    """The dynamic-array metadata of a package, nothing where it has none."""
    found = Metadata()
    if not package.has(_METADATA):
        return found
    xml = package.read(_METADATA).decode("utf-8", errors="replace")
    types = re.findall(r'<metadataType\b[^>]*\bname="([^"]*)"', xml)
    future = {name: re.findall(r"<bk>(.*?)</bk>", body, re.DOTALL)
              for name, body in re.findall(r'<futureMetadata\b[^>]*\bname="([^"]*)"[^>]*>(.*?)</futureMetadata>', xml,
                                           re.DOTALL)}
    rich = _rich_values(package)
    for kind, target in (("cellMetadata", found.dynamic), ("valueMetadata", None)):
        section = re.search(rf"<{kind}\b[^>]*>(.*?)</{kind}>", xml, re.DOTALL)
        if section is None:
            continue
        for index, block in enumerate(re.findall(r"<bk>(.*?)</bk>", section.group(1), re.DOTALL), start=1):
            record = re.search(r'<rc\b[^>]*\bt="(\d+)"[^>]*\bv="(\d+)"', block)
            if record is None or not 1 <= int(record.group(1)) <= len(types):
                continue
            name, place = types[int(record.group(1)) - 1], int(record.group(2))
            blocks = future.get(name, [])
            if place >= len(blocks):
                continue
            if target is not None and name == "XLDAPR" and 'fDynamic="1"' in blocks[place]:
                target.add(index)
            elif target is None and name == "XLRICHVALUE":
                value = re.search(r'<xlrd:rvb\b[^>]*\bi="(\d+)"', blocks[place])
                error = rich.get(int(value.group(1))) if value is not None else None
                if error is not None:
                    found.errors[index] = error
    return found


def _rich_values(package: OpcFile) -> dict[int, RichError]:
    """Each rich value that is an error a dynamic array has, by its place."""
    if not package.has(_RICH_VALUES) or not package.has(_STRUCTURES):
        return {}
    structures = [re.findall(r'<k\b[^>]*\bn="([^"]*)"', body) for body in re.findall(
        r"<s\b[^>]*>(.*?)</s>", package.read(_STRUCTURES).decode("utf-8", errors="replace"), re.DOTALL)]
    out: dict[int, RichError] = {}
    values = package.read(_RICH_VALUES).decode("utf-8", errors="replace")
    for index, (structure, body) in enumerate(re.findall(r'<rv\b[^>]*\bs="(\d+)"[^>]*>(.*?)</rv>', values,
                                                         re.DOTALL)):
        keys = structures[int(structure)] if int(structure) < len(structures) else []
        fields = dict(zip(keys, re.findall(r"<v\b[^>]*>([^<]*)</v>", body), strict=False))
        try:
            number = {key: int(text) for key, text in fields.items()}
        except ValueError:
            continue
        error = _ERRORS.get(number.get("errorType", 0))
        if error is not None:
            out[index] = RichError(error, number.get("subType", 0), number.get("rwOffset", 0),
                                   number.get("colOffset", 0))
    return out


class Collector:
    """The metadata a save gives dynamic-array formulas, gathered sheet by sheet and written after them."""

    def __init__(self) -> None:
        self.dynamic = False
        self.errors: list[RichError] = []

    def cell(self) -> int:
        """The cm of a dynamic-array formula's cell."""
        self.dynamic = True
        return 1

    def value(self, error: RichError) -> int:
        """The vm of a cell showing an error only a dynamic array has."""
        self.errors.append(error)
        return len(self.errors)

    def write(self, package: OpcFile) -> None:
        """The metadata part, and the rich values behind the errors, with their content types and relationships."""
        if not self.dynamic:
            return
        errors = self.errors
        namespaces = f' xmlns:xlrd="{_RICH}"' if errors else ""
        types = _DAPR + (_RICH_TYPE if errors else "")
        blocks = _DAPR_BLOCK
        if errors:
            blocks += (f'<futureMetadata name="XLRICHVALUE" count="{len(errors)}">'
                       + "".join(_RICH_BLOCK.format(index) for index in range(len(errors))) + "</futureMetadata>")
        values = (f'<valueMetadata count="{len(errors)}">'
                  + "".join(f'<bk><rc t="2" v="{index}"/></bk>' for index in range(len(errors)))
                  + "</valueMetadata>") if errors else ""
        xml = (f'{_HEAD}<metadata xmlns="{_MAIN}"{namespaces} xmlns:xda="{_DYNAMIC}">'
               f'<metadataTypes count="{2 if errors else 1}">{types}</metadataTypes>{blocks}'
               f'<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>{values}</metadata>')
        _put(package, _METADATA, xml, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml",
             "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata", "metadata.xml")
        if errors:
            self._write_rich(package)

    def _write_rich(self, package: OpcFile) -> None:
        shapes: list[tuple[str, ...]] = []
        rows: list[str] = []
        for error in self.errors:
            keys = _SPILL_KEYS if error.error == "#SPILL!" else _CALC_KEYS
            if keys not in shapes:
                shapes.append(keys)
            fields = {"colOffset": error.columns, "errorType": _TYPE_OF[error.error], "rwOffset": error.rows,
                      "subType": error.why}
            rows.append(f'<rv s="{shapes.index(keys)}">' + "".join(f"<v>{fields[key]}</v>" for key in keys) + "</rv>")
        values = f'{_HEAD}<rvData xmlns="{_RICH}" count="{len(rows)}">{"".join(rows)}</rvData>'
        structures = (f'{_HEAD}<rvStructures xmlns="{_RICH}" count="{len(shapes)}">'
                      + "".join('<s t="_error">' + "".join(f'<k n="{key}" t="i"/>' for key in keys) + "</s>"
                                for keys in shapes) + "</rvStructures>")
        office = "http://schemas.microsoft.com/office/2017/06/relationships/"
        _put(package, _RICH_VALUES, values, "application/vnd.ms-excel.rdrichvalue+xml", office + "rdRichValue",
             "richData/rdrichvalue.xml")
        _put(package, _STRUCTURES, structures, "application/vnd.ms-excel.rdrichvaluestructure+xml",
             office + "rdRichValueStructure", "richData/rdrichvaluestructure.xml")
        _put(package, _TYPES, _VALUE_TYPES, "application/vnd.ms-excel.rdrichvaluetypes+xml",
             office + "rdRichValueTypes", "richData/rdRichValueTypes.xml")


def _put(package: OpcFile, part: str, xml: str, content_type: str, relationship: str, target: str) -> None:
    """Write a part, giving it a content type and a relationship from the workbook where it has none yet."""
    package.write(part, xml.encode("utf-8"))
    add_content_type(package, part, content_type)
    relationships = package.read(_RELATIONSHIPS).decode("utf-8", errors="replace") if package.has(_RELATIONSHIPS) \
        else ""
    if f'Target="{target}"' not in relationships:
        add_relationship(package, _RELATIONSHIPS, relationship, target)


__all__ = ["EMPTY_ARRAY", "Collector", "Metadata", "RichError", "read"]
