"""A workbook's dxfs at a save, in the order Excel writes them, and every part's names of one following them.

Excel keeps a workbook's dxfs as a stack (scripts/measure_dxf_table.py,
tests/fixtures/dxf_table/). Each one made goes first, and a save writes
them from the first down, leaving out those nothing names any more, every
part naming each where the save wrote it. A format a macro gives a rule is
a new dxf, first, even where an earlier rule's says the same; the one the
rule had goes. Opening a file, each name of a dxf takes a dxf of its own,
so two names of one come out as two, in the order Excel reads the parts:
the stylesheet's table styles; then each sheet, the active one first and
the rest in order, with its filter's and sort's colours, then its tables
and its pivot tables; then every sheet's conditional formats in that order,
by priority.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts_file import XmlElement
    from pyopenvba.apps.excel._model import Workbook, Worksheet
    from pyopenvba.apps.excel._styles import Stylesheet
    from pyopenvba.powerquery._opc import OpcFile

#: The attributes that name a dxf by its place.
NAMES = frozenset({"dxfId", "dataDxfId", "headerRowDxfId", "totalsRowDxfId", "headerRowBorderDxfId",
                   "tableBorderDxfId", "totalsRowBorderDxfId"})
_NAME = re.compile(r'\b(?:' + "|".join(sorted(NAMES)) + r')="(\d+)"')


@dataclass
class _Name:
    """A name of a dxf: the part it stands in ("" for the stylesheet), where its number stands there, the dxf it
    names, and for a rule its priority. A rule the worksheet's extension keeps names no dxf by its place, holding
    its own, ``content``, which Excel writes among the dxfs all the same; ``rule`` is its id."""

    part: str
    start: int
    end: int
    index: int
    priority: int = 0
    content: str = ""
    rule: str = ""


#: A rule a worksheet's Excel 2010 extension keeps, and the dxf it holds of its own.
_EXTENSION_RULE = re.compile(r"<x14:cfRule\b([^>]*)>(.*?)</x14:cfRule>", re.DOTALL)
_EXTENSION_DXF = re.compile(r"<x14:dxf>(.*?)</x14:dxf>|<x14:dxf/>", re.DOTALL)


def _extension_numbered(text: str, stylesheet: Stylesheet) -> str:
    """A sheet whose extension's rules hold their number formats under the ids and in the spelling Excel gives
    their codes, as it gives every rule's (tests/fixtures/number_format_resave/)."""
    def numbered(dxf: re.Match[str]) -> str:
        return re.sub(r"<numFmt\b[^>]*?/>", lambda element: stylesheet.rule_format_element(element.group()),
                      dxf.group())

    return _EXTENSION_DXF.sub(numbered, text) if "<x14:dxf" in text else text


def _extension_rules(part: str, text: str) -> list[_Name]:
    """The rules a sheet's extension keeps with a dxf of their own, as names of it."""
    out: list[_Name] = []
    for rule in _EXTENSION_RULE.finditer(text):
        dxf = _EXTENSION_DXF.search(rule.group(2))
        if dxf is None:
            continue
        priority = re.search(r'\bpriority="(\d+)"', rule.group(1))
        guid = re.search(r'\bid="([^"]*)"', rule.group(1))
        content = "<dxf/>" if dxf.group(1) is None else f"<dxf>{dxf.group(1)}</dxf>"
        out.append(_Name(part, 0, 0, -1, int(priority.group(1)) if priority else 0, content,
                         guid.group(1) if guid else ""))
    return out


def _names(part: str, text: str, *, rules: bool) -> list[_Name]:
    """The names of a dxf ``text`` holds, in its order: its conditional formats' rules', or every other one."""
    out: list[_Name] = []
    for found in _NAME.finditer(text):
        opened = text.rfind("<", 0, found.start())
        tag = re.match(r"<([\w:.-]+)", text[opened:])
        if tag is None or (tag.group(1).rsplit(":", 1)[-1] == "cfRule") != rules:
            continue
        priority = 0
        if rules:
            head = text[opened:text.find(">", found.end())]
            given = re.search(r'\bpriority="(\d+)"', head)
            priority = int(given.group(1)) if given else 0
        out.append(_Name(part, found.start(1), found.end(1), int(found.group(1)), priority))
    return out


def _order(book: Workbook) -> list[Worksheet]:
    """The sheets in the order Excel reads them: the one active as the workbook was read first, then the rest."""
    sheets = list(book.sheets_)
    active = book.view.saved_active
    if active in sheets:
        sheets.remove(active)
        sheets.insert(0, active)
    return sheets


def number_dxfs(book: Workbook, package: OpcFile) -> None:
    """Give the stylesheet the dxfs the save writes, in Excel's order and without those nothing names, and every
    part naming one the place the save gives it, in the part's text and as the model holds the part."""
    stylesheet = book.stylesheet
    walk = [_Name("", 0, 0, index) for index in stylesheet.table_style_dxfs()]
    read: dict[str, str] = {}
    texts: dict[str, str] = {}
    order = _order(book)
    for sheet in order:
        for part in [sheet.part_name, *(table.part for table in sheet.tables),
                     *(pivot.part for pivot in sheet.pivot_tables)]:
            if part and part not in texts and package.has(part):
                read[part] = package.read(part).decode("utf-8")
                texts[part] = _extension_numbered(read[part], stylesheet) if part == sheet.part_name else read[part]
                walk += _names(part, texts[part], rules=False)
    for sheet in order:
        if sheet.part_name in texts:
            text = texts[sheet.part_name]
            walk += sorted([*_names(sheet.part_name, text, rules=True), *_extension_rules(sheet.part_name, text)],
                           key=lambda name: name.priority)
    count, base, fresh = stylesheet.dxf_count, stylesheet.dxf_base_count, stylesheet.dxf_fresh

    def rank(at: int) -> tuple[int, ...]:
        # A dxf made since the file was read or saved goes first, the newest first; the file's stand as Excel read
        # them, each name its own, or once a save wrote them as they stood.
        name = walk[at]
        if name.content:
            place = None if fresh else stylesheet.extension_places.get(name.rule)
            return (1, -at) if place is None else (1, place, -at)
        if name.index >= base:
            return 0, -name.index, -at
        return (1, -at) if fresh else (1, name.index, -at)

    kept = sorted((at for at, name in enumerate(walk) if name.content or 0 <= name.index < count), key=rank)
    place = {at: position for position, at in enumerate(kept)}
    stylesheet.write_dxf_table([walk[at].content or walk[at].index for at in kept],
                               [place.get(at, name.index) for at, name in enumerate(walk) if not name.part])
    stylesheet.extension_places = {walk[at].rule: place[at] for at in kept if walk[at].content}
    renamed: dict[str, list[tuple[_Name, int]]] = {}
    for at, name in enumerate(walk):
        if name.part and not name.content:
            renamed.setdefault(name.part, []).append((name, place.get(at, name.index)))
    for part, text in texts.items():
        named = sorted(renamed.get(part, []), key=lambda item: item[0].start)
        for name, value in reversed(named):
            text = text[:name.start] + str(value) + text[name.end:]
        if text != read[part]:
            package.write(part, text.encode("utf-8"))
        texts[part] = text
    _follow(book, texts, renamed)


def _follow(book: Workbook, texts: dict[str, str], renamed: dict[str, list[tuple[_Name, int]]]) -> None:
    """The model's own hold of each part the save renamed dxfs in: its rules' dxfIds, by priority; a table's text;
    and a pivot table's tree, in its order."""
    for sheet in book.sheets_:
        rules = {name.priority: value for name, value in renamed.get(sheet.part_name, []) if name.priority}
        for rule in sheet.conditional_formats:
            priority = rule.root.get("priority", "")
            if rule.root.get("dxfId") is not None and not rule.deleted and priority.isdecimal() \
                    and int(priority) in rules:
                rule.root.set("dxfId", str(rules[int(priority)]))
        for table in sheet.tables:
            if table.part in texts and table.xml:
                table.xml = texts[table.part]
        for pivot in sheet.pivot_tables:
            if pivot.part in renamed:
                _renamed_tree(pivot.root, [value for _, value in renamed[pivot.part]])


def _renamed_tree(root: XmlElement, values: list[int]) -> None:
    """A tree whose names of a dxf, in its order, name ``values``."""
    given = iter(values)
    for element in root.iter():
        for key in list(element.attrib):
            if key in NAMES:
                element.set(key, str(next(given, int(element.get(key, "0")))))
