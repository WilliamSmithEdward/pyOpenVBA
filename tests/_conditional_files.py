"""Conditional formats as a saved workbook keeps them, for comparing the model's files with Excel's."""
from __future__ import annotations

import posixpath
import re
import zipfile
from pathlib import Path

_GUID = re.compile(r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}")


def formatting(path: Path) -> dict[str, str]:
    """Each sheet's conditionalFormatting elements and its own extLst as a file keeps them: each dxfId the dxf it
    names, since Excel numbers its dxfs in its own order at save."""
    with zipfile.ZipFile(path) as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        targets: dict[str, str] = {}
        for element in re.findall(r"<Relationship\b[^>]*>", package.read("xl/_rels/workbook.xml.rels").decode()):
            attributes = dict(re.findall(r'(\w+)="([^"]*)"', element))
            targets[attributes["Id"]] = attributes["Target"]
        listed = re.search(r"<dxfs\b.*?</dxfs>", package.read("xl/styles.xml").decode("utf-8"), re.DOTALL)
        dxfs = re.findall(r"<dxf>.*?</dxf>", listed.group()) if listed else []
        found: dict[str, str] = {}
        for name, rid in re.findall(r'<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', workbook):
            part = posixpath.normpath(posixpath.join("xl", targets[rid].lstrip("/").removeprefix("xl/")))
            text = package.read(part).decode("utf-8")
            groups = re.findall(r"<conditionalFormatting\b.*?</conditionalFormatting>", text, re.DOTALL)
            rest = re.sub(r"<conditionalFormatting\b.*?</conditionalFormatting>", "", text, flags=re.DOTALL)
            kept = "".join(groups) + "".join(re.findall(r"<extLst>.*?</extLst>", rest, re.DOTALL))
            found[name] = re.sub(r'dxfId="(\d+)"', lambda named: f'dxf="{dxfs[int(named.group(1))]}"', kept)
    return found


def number_formats(path: Path) -> tuple[dict[str, str], list[str]]:
    """Each sheet's first rule's numFmt element as a file keeps it, by the sheet's name, and the numFmt elements of
    the stylesheet's own numFmts, the cells' formats."""
    with zipfile.ZipFile(path) as package:
        styles = package.read("xl/styles.xml").decode("utf-8")
    listed = re.search(r"<numFmts\b.*?</numFmts>", styles, re.DOTALL)
    cells = re.findall(r"<numFmt [^>]*/>", listed.group()) if listed else []
    found: dict[str, str] = {}
    for name, kept in formatting(path).items():
        rule = re.search(r'<cfRule\b[^>]*?dxf="(<dxf>.*?</dxf>)"', kept, re.DOTALL)
        number = re.search(r"<numFmt [^>]*/>", rule.group(1)) if rule else None
        found[name] = number.group() if number else ""
    return found, cells


def guids(text: str, kept: frozenset[str] = frozenset()) -> str:
    """``text`` with each GUID not in ``kept`` named by the order it first appears; the ids Excel makes for rules it
    read without one stay as they are, being Excel's own and not random."""
    seen: dict[str, str] = {}

    def name(found: re.Match[str]) -> str:
        guid = found.group()
        if guid in kept or guid.startswith("{00000000-000E-0000-"):
            return guid
        return seen.setdefault(guid, f"{{GUID-{len(seen) + 1}}}")

    return _GUID.sub(name, text)


def all_guids(text: str) -> frozenset[str]:
    """Every GUID ``text`` holds."""
    return frozenset(_GUID.findall(text))
