"""The workbookPr attributes behind Workbook members: UpdateLinks, SaveLinkValues and CheckCompatibility.

A file keeps each in its workbookPr, as Excel saves it, and only where it
is not the default: updateLinks="never" or "always", saveExternalLinkValues
="0" and checkCompatibility="1" (tests/fixtures/workbook_links.json,
tests/fixtures/small_members.json). A save writes them in the schema's
order, and leaves the element as it was where none changed.
"""
from __future__ import annotations

import re
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts_file import XmlElement
    from pyopenvba.apps.excel._model import Workbook
    from pyopenvba.powerquery._opc import OpcFile

#: CT_WorkbookPr's attributes in the schema's order, which Excel writes them in.
_ORDER = ["date1904", "showObjects", "showBorderUnselectedTables", "filterPrivacy", "promptedSolutions",
          "showInkAnnotation", "backupFile", "saveExternalLinkValues", "updateLinks", "codeName", "hidePivotFieldList",
          "showPivotChartFilter", "allowRefreshQuery", "publishItems", "checkCompatibility", "autoCompressPictures",
          "refreshAllConnections", "defaultThemeVersion"]
_UPDATES = {"never": 2, "always": 3}


def read(book: Workbook, properties: XmlElement | None) -> None:
    """The settings a file's workbookPr, parsed, gives; a file without one has the defaults."""
    if properties is None:
        return
    book.update_links = _UPDATES.get(properties.get("updateLinks", ""), 1)
    book.save_link_values = properties.get("saveExternalLinkValues", "1") not in ("0", "false")
    book.check_compatibility = properties.get("checkCompatibility", "0") in ("1", "true")


def _meaning(key: str, value: str | None) -> object:
    if key == "updateLinks":
        return _UPDATES.get(value or "", 1)
    if key == "saveExternalLinkValues":
        return value not in ("0", "false")
    return value in ("1", "true")


def write(book: Workbook, package: OpcFile) -> None:
    """The settings in the workbook part's workbookPr, where they differ from what it says."""
    if not package.has("xl/workbook.xml"):
        return
    text = package.read("xl/workbook.xml").decode("utf-8")
    found = re.search(r"<workbookPr\b[^>]*?/?>", text)
    tag = found.group(0) if found is not None else "<workbookPr/>"
    wanted = {"saveExternalLinkValues": None if book.save_link_values else "0",
              "updateLinks": {2: "never", 3: "always"}.get(book.update_links),
              "checkCompatibility": "1" if book.check_compatibility else None}
    attributes = re.findall(r'\s([\w:]+)="([^"]*)"', tag)
    have = dict(attributes)
    if all(_meaning(key, have.get(key)) == _meaning(key, value) for key, value in wanted.items()):
        return
    kept = [(key, value) for key, value in attributes if key not in wanted]
    for key, value in wanted.items():
        if value is None:
            continue
        rank = _ORDER.index(key)
        at = next((index for index, (other, _) in enumerate(kept) if other in _ORDER and _ORDER.index(other) > rank),
                  len(kept))
        kept.insert(at, (key, value))
    rebuilt = "<workbookPr" + "".join(f' {key}="{value}"' for key, value in kept) + "/>"
    if found is not None:
        text = text[:found.start()] + rebuilt + text[found.end():]
    else:
        place = re.search(r"<fileVersion\b[^>]*/>|<fileSharing\b[^>]*/>", text)
        at = place.end() if place is not None else text.index(">", text.index("<workbook")) + 1
        text = text[:at] + rebuilt + text[at:]
    package.write("xl/workbook.xml", text.encode("utf-8"))
