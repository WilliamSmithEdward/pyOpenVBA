"""Protection.AllowEditRanges: the ranges of a protected sheet a macro may still write to, as Excel keeps and saves them.

Measured in live Excel (scripts/measure_allow_edit_ranges.py, tests/fixtures/allow_edit_ranges.json):

- A title starts with a letter, is at most 255 characters, is text, and is not there already in any case. Adding, retitling
  and deleting on a protected sheet fail. Every error of Add's title, and asking for an item that is not there, is 1004.
- A protected sheet takes a value in any cell of a range with no password, or one its Unprotect opened with the password.
  The cell still reads Locked, and formats are still refused.
- A range follows its cells through inserts, deletes and cuts, and a sheet's copy has copies of its ranges.
- The sheet part keeps them as protectedRanges, after sheetProtection, a password as a SHA-512 hash with a new salt.
- A file's older 16-bit password stays as it is until its Unprotect opens it, which writes it as SHA-512
  (tests/test_excel_allow_edit_ranges.py).

Users takes Everyone, for whom Excel saves a fixed security descriptor. Any other account, which depends on the machine,
and a right other than edit report themselves.
"""
from __future__ import annotations

import base64
import os
import re
from dataclasses import dataclass, field
from html import escape
from typing import TYPE_CHECKING

from pyopenvba._a1 import parse_area
from pyopenvba.apps.excel._model import ExcelObject, Range
from pyopenvba.apps.excel._protection import FOLLOWING_PROTECTION, enforced, legacy_hash, spun
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, member, method, ref_setter, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, NOTHING, error, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet

_ELEMENT = re.compile(r"<protectedRanges>.*?</protectedRanges>|<protectedRanges/>", re.DOTALL)
_AFTER = re.compile(FOLLOWING_PROTECTION.pattern.replace("protectedRanges|", ""))
_EVERYONE = "O:WDG:WDD:(A;;CC;;;WD)"
_SPINS = 100000


def _refused() -> Exception:
    return error(1004, "Application-defined or object-defined error")


@dataclass
class Kept:
    """One range: its title, its cells, its password's hash, salt and spin count, its security descriptor, and whether
    its Unprotect has opened it."""

    title: str
    cells: Range
    hashed: tuple[str, str, int] | None = None
    descriptor: str = ""
    opened: bool = False
    #: The older 16-bit hash a file may hold in place of ``hashed``.
    legacy: str = ""


@dataclass
class SheetRanges:
    items: list[Kept] = field(default_factory=lambda: [])
    changed: bool = False


def _attributes(text: str) -> dict[str, str]:
    from html import unescape

    return {name: unescape(value) for name, value in re.findall(r'(\w+)="([^"]*)"', text)}


def state(sheet: Worksheet) -> SheetRanges:
    if sheet.allow_edit_ranges is None:
        found = SheetRanges()
        package = sheet.book.package
        if package is not None and sheet.part_name and package.has(sheet.part_name):
            element = _ELEMENT.search(package.read(sheet.part_name).decode("utf-8", errors="replace"))
            for text in re.findall(r"<protectedRange\b[^>]*/>", element.group(0) if element else ""):
                fields = _attributes(text)
                areas = [parse_area(one) for one in fields.get("sqref", "").split()]
                hashed = (fields["hashValue"], fields["saltValue"], int(fields.get("spinCount", "0"))) \
                    if "hashValue" in fields else None
                found.items.append(Kept(fields.get("name", ""), Range(sheet, areas), hashed,
                                        fields.get("securityDescriptor", ""), legacy=fields.get("password", "")))
        sheet.allow_edit_ranges = found
    return sheet.allow_edit_ranges


def editable(sheet: Worksheet, row: int, column: int) -> bool:
    """Whether a range of a protected sheet lets a macro write to a locked cell."""
    if sheet.allow_edit_ranges is None and (sheet.book.package is None or not sheet.part_name):
        return False
    for kept in state(sheet).items:
        if (kept.hashed is None and not kept.legacy) or kept.opened:
            if kept.cells.sheet is sheet and any(area.contains(row, column) for area in kept.cells.areas):
                return True
    return False


def _guard(sheet: Worksheet) -> None:
    if enforced(sheet) is not None:
        raise _refused()


def _valid(sheet: Worksheet, title: object, own: Kept | None = None) -> str:
    if not isinstance(title, str) or not 1 <= len(title) <= 255 or re.match(r"[^\W\d]", title) is None \
            or any(one is not own and one.title.casefold() == title.casefold() for one in state(sheet).items):
        raise _refused()
    return title


def _hashed(password: str) -> tuple[str, str, int]:
    salt = os.urandom(16)
    return spun(password, salt, _SPINS), base64.b64encode(salt).decode("ascii"), _SPINS


class AllowEditRanges(VBACollection, ExcelObject):
    vba_type_name = "AllowEditRanges"

    def __init__(self, sheet: Worksheet) -> None:
        self.sheet = sheet

    def vba_items(self) -> list[object]:
        return [AllowEditRange(self.sheet, kept) for kept in state(self.sheet).items]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        kept = state(self.sheet).items
        if isinstance(index, str):
            for one in kept:
                if one.title.casefold() == index.casefold():
                    return AllowEditRange(self.sheet, one)
            raise _refused()
        place = int(to_integer(index, "Long"))
        if not 1 <= place <= len(kept):
            raise _refused()
        return AllowEditRange(self.sheet, kept[place - 1])

    @method
    def Add(self, Title: object = MISSING, Range: object = MISSING, Password: object = MISSING) -> object:  # noqa: N803
        from pyopenvba.apps.excel._model import Range as Cells

        if Title is MISSING or Range is MISSING:
            raise error(450, "Wrong number of arguments or invalid property assignment")
        if Range is NOTHING:
            raise error(-2147417848, "Method 'Add' of object 'AllowEditRanges' failed")
        if not isinstance(Range, Cells):
            raise error(13, "Type mismatch")
        _guard(self.sheet)
        made = Kept(_valid(self.sheet, Title), Cells(Range.sheet, list(Range.areas)),
                    None if Password is MISSING else _hashed(to_text(Password)))
        kept = state(self.sheet)
        kept.items.append(made)
        kept.changed = True
        self.sheet.touched()
        return AllowEditRange(self.sheet, made)


class AllowEditRange(ExcelObject):
    vba_type_name = "AllowEditRange"

    def __init__(self, sheet: Worksheet, kept: Kept) -> None:
        self.sheet, self.kept = sheet, kept

    def _alive(self) -> None:
        if self.kept not in state(self.sheet).items:
            raise _refused()

    def _changed(self) -> None:
        state(self.sheet).changed = True
        self.sheet.touched()

    @member
    def Title(self) -> object:
        self._alive()
        return self.kept.title

    @setter("Title")
    def _set_title(self, value: object) -> None:
        self._alive()
        _guard(self.sheet)
        self.kept.title = _valid(self.sheet, value, self.kept)
        self._changed()

    @member
    def Range(self) -> object:
        self._alive()
        return self.kept.cells

    @ref_setter("Range")
    def _set_range(self, value: object) -> None:
        from pyopenvba.apps.excel._model import Range as Cells

        self._alive()
        _guard(self.sheet)
        if not isinstance(value, Cells):
            raise error(13, "Type mismatch")
        self.kept.cells = Cells(value.sheet, list(value.areas))
        self._changed()

    @member
    def Users(self) -> object:
        self._alive()
        return UserAccessList(self)

    @method
    def ChangePassword(self, Password: object = MISSING) -> object:
        if Password is MISSING:
            raise error(450, "Wrong number of arguments or invalid property assignment")
        self._alive()
        _guard(self.sheet)
        self.kept.hashed = _hashed(to_text(Password))
        self.kept.legacy = ""
        self.kept.opened = False
        self._changed()
        return EMPTY

    @method
    def Unprotect(self, Password: object = MISSING) -> object:
        """Opens a range with a password on a protected sheet, for the session."""
        self._alive()
        hashed = self.kept.hashed
        if self.kept.legacy and enforced(self.sheet) is not None and Password is not MISSING:
            if legacy_hash(to_text(Password)) != self.kept.legacy.upper():
                raise _refused()
            self.kept.hashed, self.kept.legacy, self.kept.opened = _hashed(to_text(Password)), "", True
            self._changed()
            return EMPTY
        if enforced(self.sheet) is None or hashed is None or Password is MISSING:
            raise _refused()
        if spun(to_text(Password), base64.b64decode(hashed[1]), hashed[2]) != hashed[0]:
            raise _refused()
        self.kept.opened = True
        return EMPTY

    @method
    def Delete(self) -> object:
        self._alive()
        _guard(self.sheet)
        state(self.sheet).items.remove(self.kept)
        self._changed()
        return EMPTY


class UserAccessList(VBACollection, ExcelObject):
    vba_type_name = "UserAccessList"

    def __init__(self, owner: AllowEditRange) -> None:
        self.owner = owner

    def vba_items(self) -> list[object]:
        descriptor = self.owner.kept.descriptor
        if descriptor and descriptor != _EVERYONE:
            raise VBAUnsupportedError("a protected range's users other than Everyone are not implemented")
        return [UserAccess()] if descriptor else []

    @method
    def Add(self, Name: object = MISSING, AllowEdit: object = MISSING) -> object:
        if Name is MISSING or AllowEdit is MISSING:
            raise error(450, "Wrong number of arguments or invalid property assignment")
        if to_text(Name) != "Everyone" or AllowEdit is not True:
            raise VBAUnsupportedError("a protected range's users other than Everyone allowed to edit are not implemented")
        self.owner.kept.descriptor = _EVERYONE
        state(self.owner.sheet).changed = True
        self.owner.sheet.touched()
        return UserAccess()


class UserAccess(ExcelObject):
    vba_type_name = "UserAccess"

    @member
    def Name(self) -> object:
        return "Everyone"

    @member
    def AllowEdit(self) -> object:
        return True


def copied(source: Worksheet, copy: Worksheet) -> None:
    """A sheet's copy has copies of its ranges."""
    kept = state(source)
    if kept.items:
        copy.allow_edit_ranges = SheetRanges([Kept(one.title, Range(copy, list(one.cells.areas)), one.hashed, one.descriptor, legacy=one.legacy)
                                              for one in kept.items], True)


def written(sheet: Worksheet, xml: str) -> str:
    """The sheet part with its ranges as the model has them."""
    kept = sheet.allow_edit_ranges
    if kept is None or not kept.changed:
        return xml
    xml = _ELEMENT.sub("", xml)
    parts: list[str] = []
    for one in kept.items:
        if not one.cells.areas:
            raise VBAUnsupportedError("saving a protected range whose cells have gone is not implemented")
        password = f'password="{one.legacy}" ' if one.legacy else "" if one.hashed is None else \
            f'algorithmName="SHA-512" hashValue="{one.hashed[0]}" saltValue="{one.hashed[1]}" spinCount="{one.hashed[2]}" '
        sqref = " ".join(area.address(absolute=False) for area in one.cells.areas)
        security = f' securityDescriptor="{one.descriptor}"' if one.descriptor else ""
        parts.append(f'<protectedRange {password}sqref="{sqref}" name="{escape(one.title)}"{security}/>')
    if parts:
        following = _AFTER.search(xml, max(xml.find("</sheetData>"), 0))
        assert following is not None
        xml = xml[:following.start()] + f'<protectedRanges>{"".join(parts)}</protectedRanges>' + xml[following.start():]
    kept.changed = False
    return xml

