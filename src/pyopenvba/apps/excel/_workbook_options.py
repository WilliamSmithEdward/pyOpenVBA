"""The workbook part's settings behind Workbook members, as Excel saves them.

A file keeps each only where it is not the default (tests/fixtures/workbook_links.json,
tests/fixtures/small_members.json, tests/fixtures/workbook_options.json):

- in its workbookPr, UpdateLinks as updateLinks="never" or "always", SaveLinkValues False as
  saveExternalLinkValues="0", CheckCompatibility as checkCompatibility="1", DisplayDrawingObjects as
  showObjects="placeholders" or "none", DisplayInkComments False as showInkAnnotation="0", and
  RemovePersonalInformation as filterPrivacy="1";
- ReadOnlyRecommended as <fileSharing readOnlyRecommended="1"/>, ahead of the workbookPr;
- EnableAutoRecover False as <fileRecoveryPr autoRecover="0"/>, after the calcPr.

A save writes the workbookPr's attributes in the schema's order, and leaves each element as it was where nothing in
it changed.
"""
from __future__ import annotations

import re
from typing import TYPE_CHECKING

from pyopenvba.exceptions import VBAUnsupportedError

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts_file import XmlElement
    from pyopenvba.apps.excel._model import Workbook
    from pyopenvba.powerquery._opc import OpcFile

_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
#: CT_WorkbookPr's attributes in the schema's order, which Excel writes them in.
_ORDER = ["date1904", "showObjects", "showBorderUnselectedTables", "filterPrivacy", "promptedSolutions",
          "showInkAnnotation", "backupFile", "saveExternalLinkValues", "updateLinks", "codeName", "hidePivotFieldList",
          "showPivotChartFilter", "allowRefreshQuery", "publishItems", "checkCompatibility", "autoCompressPictures",
          "refreshAllConnections", "defaultThemeVersion"]
_UPDATES = {"never": 2, "always": 3}
#: XlDisplayDrawingObjects: xlDisplayShapes, xlPlaceholders and xlHide, by the showObjects they are saved as.
_OBJECTS = {"all": -4104, "placeholders": 2, "none": 3}


def _true(value: str | None, default: bool) -> bool:
    return default if value is None else value in ("1", "true")


def read(book: Workbook, workbook: XmlElement) -> None:
    """The settings the parsed workbook part gives; what it leaves out has its default."""
    properties = workbook.find(_MAIN + "workbookPr")
    if properties is not None:
        book.update_links = _UPDATES.get(properties.get("updateLinks", ""), 1)
        book.save_link_values = _true(properties.get("saveExternalLinkValues"), True)
        book.check_compatibility = _true(properties.get("checkCompatibility"), False)
        book.drawing_objects = _OBJECTS.get(properties.get("showObjects", "all"), -4104)
        book.ink_comments = _true(properties.get("showInkAnnotation"), True)
        book.remove_personal = _true(properties.get("filterPrivacy"), False)
    sharing = workbook.find(_MAIN + "fileSharing")
    book.read_only_recommended = sharing is not None and _true(sharing.get("readOnlyRecommended"), False)
    recovery = workbook.find(_MAIN + "fileRecoveryPr")
    book.auto_recover = recovery is None or _true(recovery.get("autoRecover"), True)


def _meaning(key: str, value: str | None) -> object:
    if key == "updateLinks":
        return _UPDATES.get(value or "", 1)
    if key == "showObjects":
        return _OBJECTS.get(value or "all", value)
    if key in ("saveExternalLinkValues", "showInkAnnotation"):
        return _true(value, True)
    return _true(value, False)


def write(book: Workbook, package: OpcFile) -> None:
    """The settings in the workbook part, where they differ from what it says."""
    if not package.has("xl/workbook.xml"):
        return
    text = package.read("xl/workbook.xml").decode("utf-8")
    written = _with_properties(book, text)
    written = _with_element(written, "fileSharing", "readOnlyRecommended", "1" if book.read_only_recommended else None,
                            after=r"<fileVersion\b[^>]*/>")
    written = _with_element(written, "fileRecoveryPr", "autoRecover", None if book.auto_recover else "0",
                            after=r"<calcPr\b[^>]*?(?:/>|>.*?</calcPr>)")
    if written != text:
        package.write("xl/workbook.xml", written.encode("utf-8"))


def _with_properties(book: Workbook, text: str) -> str:
    if book.drawing_objects not in _OBJECTS.values():
        raise VBAUnsupportedError(f"saving DisplayDrawingObjects {book.drawing_objects}, which Excel takes but no "
                                  "showObjects names, is not implemented")
    found = re.search(r"<workbookPr\b[^>]*?/?>", text)
    tag = found.group(0) if found is not None else "<workbookPr/>"
    shown = next(name for name, value in _OBJECTS.items() if value == book.drawing_objects)
    wanted = {"saveExternalLinkValues": None if book.save_link_values else "0",
              "updateLinks": {2: "never", 3: "always"}.get(book.update_links),
              "checkCompatibility": "1" if book.check_compatibility else None,
              "showObjects": None if shown == "all" else shown,
              "showInkAnnotation": None if book.ink_comments else "0",
              "filterPrivacy": "1" if book.remove_personal else None}
    attributes = re.findall(r'\s([\w:]+)="([^"]*)"', tag)
    have = dict(attributes)
    if all(_meaning(key, have.get(key)) == _meaning(key, value) for key, value in wanted.items()):
        return text
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
        return text[:found.start()] + rebuilt + text[found.end():]
    place = re.search(r"<fileVersion\b[^>]*/>|<fileSharing\b[^>]*/>", text)
    at = place.end() if place is not None else text.index(">", text.index("<workbook")) + 1
    return text[:at] + rebuilt + text[at:]


def _with_element(text: str, name: str, attribute: str, value: str | None, *, after: str) -> str:
    """An element of one attribute the model owns: written with ``value``, or the attribute taken out, and the
    element with it where nothing else is left; placed after ``after``'s match where it is new."""
    found = re.search(rf"<{name}\b[^>]*?/>", text)
    if found is None:
        if value is None:
            return text
        place = re.search(after, text, re.DOTALL)
        at = place.end() if place is not None else text.index(">", text.index("<workbook")) + 1
        return text[:at] + f'<{name} {attribute}="{value}"/>' + text[at:]
    pairs = [(key, held) for key, held in re.findall(r'\s([\w:]+)="([^"]*)"', found.group(0)) if key != attribute]
    if _true(dict(re.findall(r'\s([\w:]+)="([^"]*)"', found.group(0))).get(attribute), attribute == "autoRecover") \
            == _true(value, attribute == "autoRecover"):
        return text
    if value is not None:
        pairs.insert(0, (attribute, value))
    if not pairs:
        return text[:found.start()] + text[found.end():]
    return text[:found.start()] + f"<{name}" + "".join(f' {key}="{held}"' for key, held in pairs) + "/>" \
        + text[found.end():]
