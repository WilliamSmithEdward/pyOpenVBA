"""Hyperlinks: Worksheet.Hyperlinks, Range.Hyperlinks, the Hyperlink object, and what edits do to links.

Measured in live Excel (scripts/measure_hyperlinks.py, tests/fixtures/hyperlinks/):

- Hyperlinks.Add puts a link on each area of its anchor and answers the
  last. One over the very cells of a link the sheet has takes its place in
  the list; over any other cells it goes at the end, overlapping or not.
  Every cell of the anchor is given the style named Hyperlink -- the
  built-in one, made the first time, or one a macro made by that name --
  and the first cell the text to show, or, when it is empty, the address
  as given: a number stays a number.
- A link is named by the text to show it was given, or by its address and
  subaddress, "address - subaddress", which follows a new address; one
  read from a file without a display is named by its cell's text until a
  new address names it. TextToDisplay reads the first cell's text, "" for
  a number or a formula. SubAddress takes the address away; Address keeps
  the subaddress unless the new one has a # of its own.
- EmailSubject is the subject of a mailto address, which setting it
  writes in, "" for no address, and fails for another address. Set on a
  link with no address, it leaves Excel refusing every change after, so
  the model refuses it.
- A range's collection holds the links inside one of its areas, and, for
  a range of one area, the link it lies in. Clear, ClearContents, an empty
  value and ClearHyperlinks take away the links a range holds this way,
  as does a copy landing on them; Delete takes them away with their
  cells' formats, keeping merges. A copy brings the links it meets over
  the cells it copies, at the end of the list; a cut moves them in place.
  Inserts and deletes move a link as they move a reference, and a sort
  moves a link one line high with its line.
- On a protected sheet a link needs AllowInsertingHyperlinks and unlocked
  cells, and a locked cell keeps its link; a link's properties can still
  be set.
- A formula whose first function is HYPERLINK gives its cells the
  Hyperlink style.
"""

from __future__ import annotations

import uuid
from dataclasses import dataclass, replace
from typing import TYPE_CHECKING, Final

from pyopenvba._a1 import Area
from pyopenvba.apps.excel._model import ExcelObject, Range
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, VBAObject, member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, NOTHING, NULL, VBAInt, error, to_integer, to_text

if TYPE_CHECKING:
    from collections.abc import Callable

    from pyopenvba.apps.excel._model import Worksheet

_CREATOR: Final = 1480803660
#: The error Excel's hyperlinks fail with (E_FAIL): a subject for an address that is not mailto, the name of a
#: link with nothing to be named by, an anchor of Nothing.
_FAILED: Final = -2147467259


@dataclass(eq=False)
class Link:
    """One hyperlink: the cells it is on, where it goes, its tip, and what it is named. Two links are the same link
    only when they are one object."""

    area: Area
    #: As Excel keeps it once given (see _hyperlinks_file.tidied); "" for none.
    address: str = ""
    sub: str = ""
    tip: str = ""
    #: The name TextToDisplay or the file's display gave it, which a new address keeps when ``given``; for a
    #: link read without one, its cell's text, which a new address replaces. None: named after its address.
    display: str | None = None
    given: bool = False
    uid: str = ""


def derived(address: str, sub: str) -> str:
    """The name a link's address and subaddress give it."""
    return f"{address} - {sub}" if address and sub else address or sub


def name_of(link: Link) -> str:
    """Hyperlink.Name; a link with nothing to be named by fails."""
    name = link.display if link.display is not None else derived(link.address, link.sub)
    if link.display is None and not name:
        raise error(_FAILED, "Method 'Name' of object 'Hyperlink' failed")
    return name


def shown_text(sheet: Worksheet, link: Link) -> str:
    """TextToDisplay: the text the link's first cell holds, "" for a number, a formula or nothing, even a formula
    that gives text (tests/fixtures/hyperlinks/formulas.xlsx)."""
    cell = sheet.cells_.get((link.area.top, link.area.left))
    if cell is None or cell.formula:
        return ""
    return cell.value if isinstance(cell.value, str) else ""


def _changed(sheet: Worksheet) -> None:
    sheet.links_changed = True
    sheet.touched()


def _new_uid(sheet: Worksheet) -> str:
    return "{" + str(uuid.uuid4()).upper() + "}" if sheet.declares_revisions() else ""


# --- which links a range holds ---------------------------------------------------------------------------------


def _inside(inner: Area, outer: Area) -> bool:
    return outer.top <= inner.top and inner.bottom <= outer.bottom and outer.left <= inner.left \
        and inner.right <= outer.right


def _meets(one: Area, other: Area) -> bool:
    return not (one.bottom < other.top or other.bottom < one.top or one.right < other.left or other.right < one.left)


def held(link: Link, areas: list[Area]) -> bool:
    """Whether a range of these areas holds the link: the link is inside one of them, or the range is one area
    inside the link."""
    return any(_inside(link.area, area) for area in areas) or (len(areas) == 1 and _inside(areas[0], link.area))


def links_of(sheet: Worksheet, areas: list[Area] | None) -> list[Link]:
    return [link for link in sheet.hyperlinks if areas is None or held(link, areas)]


def touches(sheet: Worksheet, areas: list[Area]) -> bool:
    """Whether any link meets any of the areas."""
    return any(_meets(link.area, area) for link in sheet.hyperlinks for area in areas)


def _refuse_locked(sheet: Worksheet, links: list[Link]) -> None:
    from pyopenvba.apps.excel._protection import PROTECTED, any_locked, enforced

    if links and enforced(sheet) is not None and any_locked(sheet, [link.area for link in links]):
        raise error(1004, PROTECTED)


def _take(sheet: Worksheet, links: list[Link], *, formats: bool) -> None:
    """Take links off the sheet; with ``formats``, their cells go back to the default format as ClearFormats puts
    them, merges kept."""
    from pyopenvba.apps.excel import _row_formats

    if not links:
        return
    for link in links:
        sheet.hyperlinks.remove(link)
        if formats and not _row_formats.clear_area(sheet, link.area):
            default = sheet.book.stylesheet.default
            for row in range(link.area.top, link.area.bottom + 1):
                for column in range(link.area.left, link.area.right + 1):
                    sheet.restyle(row, column, default)
    _changed(sheet)


def cleared(sheet: Worksheet, areas: list[Area]) -> None:
    """Clear, ClearContents or an empty value: the links the range holds go, their cells' formats as they are."""
    if sheet.hyperlinks:
        _take(sheet, [link for link in sheet.hyperlinks if held(link, areas)], formats=False)


def empties(value: object) -> bool:
    """Whether a value written to cells empties them, which takes their links away as ClearContents does."""
    return value is EMPTY or value == ""


# --- adding a link ------------------------------------------------------------------------------------------


def _text_argument(value: object) -> str:
    """A SubAddress, ScreenTip or TextToDisplay for Add: a string, or error 5."""
    if value is MISSING:
        return ""
    if not isinstance(value, str):
        raise error(5, "Invalid procedure call or argument")
    return value


def add(sheet: Worksheet, anchor: object, address: object, sub: object, tip: object, text: object) -> Hyperlink:
    """Hyperlinks.Add: a link on each area of the anchor, which may be on another sheet; the last one."""
    from pyopenvba.apps.excel._hyperlinks_file import given
    from pyopenvba.apps.excel._protection import PROTECTED, any_locked, enforced

    if anchor is NOTHING:
        raise error(_FAILED, "Method 'Add' of object 'Hyperlinks' failed")
    if not isinstance(anchor, Range):
        if isinstance(anchor, VBAObject):
            raise VBAUnsupportedError("a hyperlink on a shape, or on anything but a range, is not implemented")
        raise error(13, "Type mismatch")
    if address is MISSING:
        raise error(450, "Wrong number of arguments or invalid property assignment")
    if address is NULL:
        raise error(13, "Type mismatch")
    sub_text, tip_text, shown = _text_argument(sub), _text_argument(tip), _text_argument(text)
    target = anchor.sheet
    if enforced(target) is not None and ("AllowInsertingHyperlinks" not in target.protection_allows
                                         or any_locked(target, anchor.areas)):
        raise error(1004, PROTECTED)
    written, from_hash, hashed = given(to_text(address))
    link: Link | None = None
    for area in anchor.areas:
        block = Area(area.top, area.left, area.bottom, area.right)
        link = Link(block, written, sub_text or (from_hash if hashed else ""), tip_text, uid=_new_uid(target))
        if shown:
            link.display, link.given = shown, True
        same = next((index for index, one in enumerate(target.hyperlinks) if one.area == block), None)
        if same is None:
            target.hyperlinks.append(link)
        else:
            target.hyperlinks[same] = link
        _dress(target, link, address, shown)
    _changed(target)
    assert link is not None
    return Hyperlink(target, link)


def _dress(sheet: Worksheet, link: Link, address: object, shown: str) -> None:
    """A new link's cells: the Hyperlink style on each, and on the first the text to show, or where it is empty
    the link's name -- the address given, where that was a number or True or False."""
    from pyopenvba.apps.excel import _merges
    from pyopenvba.apps.excel._cell_styles import give_link_style

    give_link_style(Range(sheet, [link.area]))
    row, column = link.area.top, link.area.left
    if not _merges.writable(sheet, row, column):
        return
    cell = sheet.cells_.get((row, column))
    if shown:
        value: object = shown
    elif cell is None or (cell.value is EMPTY and not cell.formula):
        value = derived(link.address, link.sub) if isinstance(address, str) else address
    else:
        return
    if value != "":
        Range(sheet, [Area(row, column, row, column)]).vba_set("Value", value)


# --- moving with the cells ---------------------------------------------------------------------------------------


def moved(sheet: Worksheet, rewrite: Callable[[str], str]) -> None:
    """Carry each link through an insert or a delete as a reference is carried; one whose cells go goes."""
    from pyopenvba._a1 import parse_area

    changed = False
    for link in list(sheet.hyperlinks):
        text = rewrite("=" + link.area.address(absolute=False))
        try:
            area = parse_area(text.removeprefix("="), sheet="")
        except ValueError:
            sheet.hyperlinks.remove(link)
            changed = True
            continue
        area = Area(area.top, area.left, area.bottom, area.right)
        if area != link.area:
            link.area = area
            changed = True
    if changed:
        sheet.links_changed = True


def copied(source: Worksheet, area: Area, target: Worksheet, landing: Area) -> None:
    """A copy of ``area`` landing at ``landing``: the links there go, and each link the copy meets comes along over
    the cells of it copied, at the end of the list."""
    down, across = landing.top - area.top, landing.left - area.left
    pieces = [(link, Area(max(link.area.top, area.top) + down, max(link.area.left, area.left) + across,
                          min(link.area.bottom, area.bottom) + down, min(link.area.right, area.right) + across))
              for link in source.hyperlinks if _meets(link.area, area)]
    gone = [link for link in target.hyperlinks if held(link, [landing])]
    if not pieces and not gone:
        return
    for link in gone:
        target.hyperlinks.remove(link)
    for link, placed in pieces:
        target.hyperlinks.append(replace(link, area=placed, uid=_new_uid(target)))
    _changed(target)


def cut(source: Worksheet, area: Area, target: Worksheet, down: int, across: int) -> None:
    """A cut: the links inside ``area`` move ``down`` and ``across``, keeping their place in the list on their own
    sheet, and the links where they land go."""
    moving = [link for link in source.hyperlinks if _inside(link.area, area)]
    landing = Area(area.top + down, area.left + across, area.bottom + down, area.right + across)
    gone = [link for link in target.hyperlinks if link not in moving and held(link, [landing])]
    if not moving and not gone:
        return
    for link in gone:
        target.hyperlinks.remove(link)
    for link in moving:
        link.area = Area(link.area.top + down, link.area.left + across, link.area.bottom + down,
                         link.area.right + across)
        if target is not source:
            source.hyperlinks.remove(link)
            target.hyperlinks.append(link)
    _changed(source)
    _changed(target)


def sorted_lines(sheet: Worksheet, block: Area, destination: dict[int, int], *, across: bool) -> None:
    """A sort of ``block``'s rows, or its columns ``across``: a link inside it one line high moves with its line;
    one over several lines stays."""
    for link in sheet.hyperlinks:
        if not _inside(link.area, block):
            continue
        first, last = (link.area.left, link.area.right) if across else (link.area.top, link.area.bottom)
        if first != last or first not in destination or destination[first] == first:
            continue
        line = destination[first]
        link.area = Area(link.area.top, line, link.area.bottom, line) if across \
            else Area(line, link.area.left, line, link.area.right)
        sheet.links_changed = True


def sheet_copied(source: Worksheet, copy: Worksheet) -> None:
    """A copy of a sheet has its links, in their order."""
    copy.hyperlinks = [replace(link, uid=_new_uid(copy)) for link in source.hyperlinks]
    copy.links_changed = bool(copy.hyperlinks)


def styles_as_link(sheet: Worksheet, formula: str) -> bool:
    """Whether a formula written to cells gives them the Hyperlink style: its first function is HYPERLINK."""
    from pyopenvba.apps.excel._formula_format import first_call

    return first_call(sheet, formula) == "HYPERLINK"


# --- the object model ---------------------------------------------------------------------------------------------


def _mailto(address: str) -> bool:
    return address[:7].casefold() == "mailto:"


def _subject(address: str) -> str:
    """The subject a mailto address names, "" for none."""
    _, _, query = address.partition("?")
    for pair in query.split("&"):
        key, equals, value = pair.partition("=")
        if equals and key.casefold() == "subject":
            return value
    return ""


def _with_subject(address: str, subject: str) -> str:
    """A mailto address with its subject set: in place of the one it names, after the rest, or taken out."""
    head, question, query = address.partition("?")
    pairs = [pair for pair in query.split("&") if pair] if question else []
    at = next((index for index, pair in enumerate(pairs) if pair.partition("=")[0].casefold() == "subject"), None)
    if at is None:
        if subject:
            pairs.append(f"subject={subject}")
    elif subject:
        pairs[at] = f"{pairs[at].partition('=')[0]}={subject}"
    else:
        del pairs[at]
    return head + ("?" + "&".join(pairs) if pairs else "")


class Hyperlink(ExcelObject):
    """One link on a sheet's cells."""

    vba_type_name = "Hyperlink"

    def __init__(self, sheet: Worksheet, link: Link) -> None:
        self.sheet = sheet
        self.link = link

    def _set(self, change: Callable[[Link], None]) -> None:
        change(self.link)
        _changed(self.sheet)

    @member
    def Address(self) -> object:
        return self.link.address

    @setter("Address")
    def _set_address(self, value: object) -> None:
        from pyopenvba.apps.excel._hyperlinks_file import given

        address, sub, hashed = given(to_text(value))

        def change(link: Link) -> None:
            link.address = address
            if hashed:
                link.sub = sub
            if not link.given:
                link.display = None
        self._set(change)

    @member
    def SubAddress(self) -> object:
        return self.link.sub

    @setter("SubAddress")
    def _set_sub_address(self, value: object) -> None:
        text = to_text(value)

        def change(link: Link) -> None:
            link.address, link.sub = "", text
            if not link.given:
                link.display = None
        self._set(change)

    @member
    def ScreenTip(self) -> object:
        return self.link.tip

    @setter("ScreenTip")
    def _set_screen_tip(self, value: object) -> None:
        text = to_text(value)
        self._set(lambda link: setattr(link, "tip", text))

    @member
    def TextToDisplay(self) -> object:
        return shown_text(self.sheet, self.link)

    @setter("TextToDisplay")
    def _set_text_to_display(self, value: object) -> None:
        from pyopenvba.apps.excel import _merges

        if value == "":
            return
        name = to_text(value)
        row, column = self.link.area.top, self.link.area.left
        if _merges.writable(self.sheet, row, column):
            Range(self.sheet, [Area(row, column, row, column)]).vba_set("Value", value)
        self.link.display, self.link.given = name, True
        _changed(self.sheet)

    @member
    def Name(self) -> object:
        return name_of(self.link)

    @setter("Name")
    def _set_name(self, value: object) -> None:
        raise error(450, "Wrong number of arguments or invalid property assignment")

    @member
    def EmailSubject(self) -> object:
        if not self.link.address:
            return ""
        if not _mailto(self.link.address):
            raise error(_FAILED, "Method 'EmailSubject' of object 'Hyperlink' failed")
        return _subject(self.link.address)

    @setter("EmailSubject")
    def _set_email_subject(self, value: object) -> None:
        if not self.link.address:
            raise VBAUnsupportedError("EmailSubject on a hyperlink with no address, which leaves Excel refusing "
                                      "every change after it, is not implemented")
        if not _mailto(self.link.address):
            raise error(_FAILED, "Method 'EmailSubject' of object 'Hyperlink' failed")
        subject = to_text(value)

        def change(link: Link) -> None:
            link.address = _with_subject(link.address, subject)
            if not link.given:
                link.display = None
        self._set(change)

    @member
    def Type(self) -> object:
        return VBAInt(0, "Long")  # msoHyperlinkRange

    @member
    def Range(self) -> object:
        return Range(self.sheet, [self.link.area])

    @member
    def Shape(self) -> object:
        raise error(1004, "Application-defined or object-defined error")

    @method
    def Delete(self) -> object:
        _refuse_locked(self.sheet, [self.link])
        if self.link in self.sheet.hyperlinks:
            _take(self.sheet, [self.link], formats=True)
        return EMPTY

    @member
    def Parent(self) -> object:
        return Range(self.sheet, [self.link.area])

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @member
    def Creator(self) -> object:
        return VBAInt(_CREATOR, "Long")


class Hyperlinks(VBACollection, ExcelObject):
    """A sheet's links, or the ones a range holds, in the order the sheet lists them."""

    vba_type_name = "Hyperlinks"

    def __init__(self, sheet: Worksheet, areas: list[Area] | None = None) -> None:
        self.sheet = sheet
        self.areas = None if areas is None else [Area(one.top, one.left, one.bottom, one.right) for one in areas]

    def vba_items(self) -> list[object]:
        return [Hyperlink(self.sheet, link) for link in links_of(self.sheet, self.areas)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            key = index.casefold()
            for item in items:
                assert isinstance(item, Hyperlink)
                link = item.link
                name = link.display if link.display is not None else derived(link.address, link.sub)
                if name.casefold() == key:
                    return item
            raise error(9, "Subscript out of range")
        position = int(to_integer(index, "Long"))
        if 1 <= position <= len(items):
            return items[position - 1]
        raise error(9, "Subscript out of range")

    @method
    def Add(self, Anchor: object = MISSING, Address: object = MISSING, SubAddress: object = MISSING,
            ScreenTip: object = MISSING, TextToDisplay: object = MISSING) -> object:
        return add(self.sheet, Anchor, Address, SubAddress, ScreenTip, TextToDisplay)

    @method
    def Delete(self) -> object:
        links = links_of(self.sheet, self.areas)
        _refuse_locked(self.sheet, links)
        _take(self.sheet, links, formats=True)
        return EMPTY

    @member
    def Parent(self) -> object:
        return self.sheet

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @member
    def Creator(self) -> object:
        return VBAInt(_CREATOR, "Long")


def clear_links(target: Range) -> None:
    """Range.ClearHyperlinks: the links the range holds go, their cells' formats as they are."""
    links = links_of(target.sheet, list(target.areas))
    _refuse_locked(target.sheet, links)
    if links:
        _take(target.sheet, links, formats=False)
