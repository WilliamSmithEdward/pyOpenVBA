"""xl/calcChain.xml, the order Excel last worked a workbook's formulas out in.

Measured in live Excel (scripts/measure_calc_chain.py, tests/fixtures/calc_chain.json).

A chain arrives with a file and stays byte for byte while Excel can still open it. Excel refuses the whole file, with
error 1004 or with its repair prompt, for one entry naming a cell with no formula, a sheet id no sheet has, or an array
flag on a formula that is no array, and for a chain with no entries. A formula left out of the chain costs nothing: it
still recalculates. An entry's ``i`` is the sheet's sheetId, and an entry without one has the sheet of the entry before.
A cell inside the ref of an array or a data table counts as a formula, though only the first holds an ``f``.

A chain that no longer holds is dropped with its Override and Relationship, and Excel builds its own as it opens the
file. This does not match Excel's bytes: after the same edit Excel writes a chain, reordered by the evaluation it ran.
The byte fidelity rule is broken knowingly here, and matching those bytes is a later measuring job.
"""
from __future__ import annotations

import re
from collections.abc import Callable
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, column_number, parse_area
from pyopenvba._relationships import in_office_order
from pyopenvba._xml import attributes as tag_attributes

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Workbook
    from pyopenvba.powerquery._opc import OpcFile

_RELATIONSHIPS = 'xl/_rels/workbook.xml.rels'
_CELL = re.compile(r'<c\b([^>]*?)(/?)>\s*(<f\b[^>]*>)?')
_REFERENCE = re.compile(r'([A-Z]{1,3})([1-9][0-9]*)')
_TRUE = ('1', 'true')


class _SheetFormulas:
    """Where one sheet's written xml holds formulas, by the cell's reference and by the areas of its arrays and data tables."""

    def __init__(self, xml: str) -> None:
        self.cells: set[str] = set()
        self.arrays: list[Area] = []
        self.tables: list[Area] = []
        start, end = xml.find('<sheetData'), xml.find('</sheetData>')
        if start < 0 or end < 0:
            return
        for found in _CELL.finditer(xml, start, end):
            if found.group(2) or not found.group(3):
                continue
            reference = tag_attributes(f'<c{found.group(1)}>').get('r', '')
            self.cells.add(reference)
            formula = tag_attributes(found.group(3))
            if formula.get('t') == 'array' and formula.get('ref'):
                self.arrays.append(parse_area(formula['ref']))
            elif formula.get('t') == 'dataTable' and formula.get('ref'):
                self.tables.append(parse_area(formula['ref']))

    def names(self, reference: str, *, array: bool) -> bool:
        """Whether a chain may name this cell, as an entry flagged as an array's or not."""
        found = _REFERENCE.fullmatch(reference)
        if found is None:
            return False
        column, row = column_number(found.group(1)), int(found.group(2))

        def within(areas: list[Area]) -> bool:
            return any(area.top <= row <= area.bottom and area.left <= column <= area.right for area in areas)

        if array:
            return within(self.arrays)
        return reference in self.cells or within(self.arrays) or within(self.tables)


def holds(chain: str, sheets: dict[int, str]) -> bool:
    """Whether Excel opens this chain given each sheetId's written xml."""
    entries = re.findall(r'<c\b[^>]*>', chain)
    if not entries:
        return False
    formulas: dict[int, _SheetFormulas] = {}
    sheet: int | None = None
    for entry in entries:
        attributes = tag_attributes(entry)
        if 'i' in attributes:
            sheet = int(attributes['i']) if attributes['i'].isdecimal() else None
        if sheet is None or sheet not in sheets:
            return False
        if sheet not in formulas:
            formulas[sheet] = _SheetFormulas(sheets[sheet])
        reference = attributes.get('r', '').replace('$', '').upper()
        if not formulas[sheet].names(reference, array=attributes.get('a', '').lower() in _TRUE):
            return False
    return True


def settle(book: Workbook, package: OpcFile, relationships: dict[str, tuple[str, str]],
           remove_part: Callable[[OpcFile, str], None]) -> None:
    """Drop the chain from a package about to be written if Excel would refuse it. Run after every sheet is written."""
    kept = {key: part for key, (part, kind) in relationships.items() if kind.endswith('/calcChain')}
    for relationship, part in kept.items():
        if not package.has(part):
            continue
        sheets = {sheet.sheet_id: package.read(sheet.part_name).decode('utf-8', errors='replace')
                  for sheet in book.sheets_ if sheet.part_name and package.has(sheet.part_name)}
        if holds(package.read(part).decode('utf-8', errors='replace'), sheets):
            continue
        text = package.read(_RELATIONSHIPS).decode('utf-8')
        text = re.sub(r'<Relationship\b[^>]*/>',
                      lambda found: '' if tag_attributes(found.group(0)).get('Id') == relationship else found.group(0), text)
        package.write(_RELATIONSHIPS, in_office_order(text).encode('utf-8'))
        remove_part(package, part)
