"""Members of Workbook, Worksheet and PageSetup a macro meets with no screen, server or mail system: the theme's colours
and fonts, custom XML parts, read-only access, and what answers the same on any workbook that shares and reviews nothing.

Measured in live Excel (scripts/measure_workbook_members.py, tests/fixtures/workbook_members.json).

A change to a theme's colour reaches every cell using it and is saved. A change to a theme's font, a custom XML part
added and then saved, and going back to read and write access, where Excel reopens the file, report themselves.
"""
from __future__ import annotations

import re
import uuid
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._model import ExcelObject
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import MISSING, NOTHING, VBAArray, VBAInt, error, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Workbook
    from pyopenvba.apps.excel._page_setup import PageSetup
    from pyopenvba.powerquery._opc import OpcFile

_APP_DEFINED = "Application-defined or object-defined error"
#: Places in the file's clrScheme, by msoThemeColorSchemeIndex from 1, and in Colors.theme, which starts with the lights.
_SCHEME = ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink",
           "folHlink")
_THEME_PLACE = (1, 0, 3, 2, 4, 5, 6, 7, 8, 9, 10, 11)
_XML_FAILED = -2147467259
_BUILT_IN_XML = ("http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
                 "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
                 "http://schemas.microsoft.com/office/2006/coverPageProps")


def failed(member_name: str) -> Exception:
    return error(1004, f"{member_name} method of Workbook class failed")


def app_defined() -> Exception:
    return error(1004, _APP_DEFINED)


class _Owned(ExcelObject):
    """An object of a workbook that answers its Parent and Application."""

    def __init__(self, owner: object, book: Workbook) -> None:
        self.owner, self.book = owner, book

    @member
    def Parent(self) -> object:
        return self.owner

    @member
    def Application(self) -> object:
        return self.book.application


class OfficeTheme(_Owned):
    vba_type_name = "OfficeTheme"

    def __init__(self, book: Workbook) -> None:
        super().__init__(book, book)

    @member
    def ThemeColorScheme(self) -> object:
        return ThemeColorScheme(self, self.book)

    @member
    def ThemeFontScheme(self) -> object:
        return ThemeFontScheme(self, self.book)

    @member
    def ThemeEffectScheme(self) -> object:
        return ThemeEffectScheme(self, self.book)


class ThemeEffectScheme(_Owned):
    vba_type_name = "ThemeEffectScheme"


class ThemeColorScheme(_Owned):
    vba_type_name = "ThemeColorScheme"

    @member
    def Count(self) -> object:
        return VBAInt(len(_SCHEME), "Long")

    @method
    def Colors(self, Index: object = MISSING) -> object:
        place = int(to_integer(Index, "Long"))
        if not 1 <= place <= len(_SCHEME):
            raise error(5, "The specified value is out of range.")
        return ThemeColor(self, self.book, place)


class ThemeColor(_Owned):
    vba_type_name = "ThemeColor"

    def __init__(self, owner: object, book: Workbook, place: int) -> None:
        super().__init__(owner, book)
        self.place = place

    @member(default=True)
    def RGB(self) -> object:
        from pyopenvba.apps.excel._styles import bgr

        return VBAInt(bgr(self.book.stylesheet.colors.theme[_THEME_PLACE[self.place - 1]]), "Long")

    @setter("RGB")
    def _set_rgb(self, value: object) -> None:
        number = int(to_integer(value, "Long")) & 0xFFFFFF
        colors = self.book.stylesheet.colors
        spelled = f"{number & 255:02X}{(number >> 8) & 255:02X}{(number >> 16) & 255:02X}"
        theme = list(colors.theme)
        theme[_THEME_PLACE[self.place - 1]] = spelled
        colors.theme = tuple(theme)
        self.book.theme_colours_changed = True
        self.book.saved = False


class ThemeFontScheme(_Owned):
    vba_type_name = "ThemeFontScheme"

    @member
    def MajorFont(self) -> object:
        return ThemeFonts(self, self.book, 0)

    @member
    def MinorFont(self) -> object:
        return ThemeFonts(self, self.book, 1)


class ThemeFonts(_Owned):
    vba_type_name = "ThemeFonts"

    def __init__(self, owner: object, book: Workbook, which: int) -> None:
        super().__init__(owner, book)
        self.which = which

    @member
    def Count(self) -> object:
        return VBAInt(3, "Long")

    @method(default=True)
    def Item(self, Index: object = MISSING) -> object:
        place = int(to_integer(Index, "Long"))
        if not 1 <= place <= 3:
            raise error(5, "The specified value is out of range.")
        return ThemeFont(self, self.book, self.which, place)


class ThemeFont(_Owned):
    vba_type_name = "ThemeFont"

    def __init__(self, owner: object, book: Workbook, which: int, place: int) -> None:
        super().__init__(owner, book)
        self.which, self.place = which, place

    @member
    def Name(self) -> object:
        # msoThemeLatin, msoThemeComplexScript and msoThemeEastAsian: a new workbook names only the first.
        return self.book.stylesheet.theme_fonts[self.which] if self.place == 1 else ""

    @setter("Name")
    def _set_name(self, value: object) -> None:
        raise VBAUnsupportedError("changing a theme's font is not implemented")


def write_theme(book: Workbook, package: OpcFile) -> None:
    """The colours a macro changed, written into the file's theme part."""
    if not book.theme_colours_changed:
        return
    part = next(name for name in package.names() if re.fullmatch(r"xl/theme/theme\d+\.xml", name))
    text = package.read(part).decode("utf-8")
    for place, name in enumerate(_SCHEME):
        spelled = book.stylesheet.colors.theme[_THEME_PLACE[place]]
        text = re.sub(rf"(<a:{name}>).*?(</a:{name}>)", rf'\g<1><a:srgbClr val="{spelled}"/>\g<2>', text, count=1,
                      flags=re.DOTALL)
    package.write(part, text.encode("utf-8"))
    book.theme_colours_changed = False


def apply_theme(book: Workbook, file_name: object) -> None:
    from pathlib import Path

    if file_name is MISSING:
        raise error(450, "Wrong number of arguments or invalid property assignment")
    book.application.outside("Workbook.ApplyTheme")
    text = to_text(file_name)
    if not text or not Path(text).is_file():
        raise app_defined()
    raise VBAUnsupportedError("Workbook.ApplyTheme of a theme file is not implemented")


class CustomXMLPart(_Owned):
    vba_type_name = "CustomXMLPart"

    def __init__(self, owner: object, book: Workbook, namespace: str, text: str | None, built_in: bool) -> None:
        super().__init__(owner, book)
        self.namespace, self.text, self.built_in = namespace, text, built_in
        self.identity = "{" + str(uuid.uuid4()).upper() + "}"

    @member
    def BuiltIn(self) -> object:
        return self.built_in

    @member
    def NamespaceURI(self) -> object:
        return self.namespace

    @member
    def XML(self) -> object:
        if self.text is None:
            raise VBAUnsupportedError("the XML of a workbook's built-in custom XML parts is not implemented")
        return self.text

    @member
    def ID(self) -> object:
        return self.identity

    @method
    def Delete(self) -> object:
        if self.built_in:
            raise error(-536602603, "This custom XML part is required for Office to function properly and cannot be "
                                    "deleted.")
        parts = _custom_parts(self.book)
        if self in parts:
            parts.remove(self)
        return None


def _custom_parts(book: Workbook) -> list[CustomXMLPart]:
    """The workbook's parts, its three built in first and then those its file holds."""
    parts = book.custom_xml_parts_
    if parts is None:
        parts = [CustomXMLPart(book, book, namespace, None, True) for namespace in _BUILT_IN_XML]
        package = book.package
        if package is not None:
            for name in sorted(package.names()):
                if re.fullmatch(r"customXml/item\d+\.xml", name):
                    text = package.read(name).decode("utf-8-sig", errors="replace")
                    parts.append(CustomXMLPart(book, book, _namespace_of(text), text, False))
        book.custom_xml_parts_ = parts
        book.custom_xml_read = len(parts)
    return parts


def _namespace_of(text: str) -> str:
    from pyopenvba.apps.excel._charts_file import xml_tree

    if not text:
        return ""
    try:
        root = xml_tree(text)
    except Exception as failure:
        raise error(_XML_FAILED, "Invalid at the top level of the document.\n") from failure
    return root.tag[1:].partition("}")[0] if root.tag.startswith("{") else ""


class CustomXMLParts(VBACollection, _Owned):
    vba_type_name = "CustomXMLParts"

    def __init__(self, book: Workbook, chosen: list[CustomXMLPart] | None = None) -> None:
        _Owned.__init__(self, book, book)
        self.chosen = chosen

    def vba_items(self) -> list[object]:
        return list(_custom_parts(self.book) if self.chosen is None else self.chosen)

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            raise error(9, "Subscript out of range")
        place = int(to_integer(index, "Long"))
        if place < 1:
            raise error(5, "This operation failed because one or more changes made during events were invalid.")
        if place > len(items):
            raise error(9, "Subscript out of range")
        return items[place - 1]

    @method
    def Add(self, XML: object = MISSING, SchemaCollection: object = MISSING) -> object:
        if SchemaCollection is not MISSING:
            raise VBAUnsupportedError("CustomXMLParts.Add with a schema collection is not implemented")
        text = "" if XML is MISSING else to_text(XML)
        made = CustomXMLPart(self.book, self.book, _namespace_of(text), re.sub(r"\s+/>", "/>", text), False)
        _custom_parts(self.book).append(made)
        self.book.saved = False
        return made

    @method
    def SelectByNamespace(self, NamespaceURI: object = MISSING) -> object:
        wanted = to_text(NamespaceURI)
        return CustomXMLParts(self.book, [part for part in _custom_parts(self.book) if part.namespace == wanted])

    @method
    def SelectByID(self, Id: object = MISSING) -> object:
        wanted = to_text(Id)
        return next((part for part in _custom_parts(self.book) if part.identity == wanted), NOTHING)


def refuse_added_parts(book: Workbook) -> None:
    """Report a custom XML part a save would have to write."""
    parts = book.custom_xml_parts_
    if parts is not None and len(parts) > book.custom_xml_read:
        raise VBAUnsupportedError("saving a custom XML part a macro added is not implemented")


def change_file_access(book: Workbook, mode: object) -> object:
    """Workbook.ChangeFileAccess: a saved workbook goes read-only, its cells as they stand and marked saved. A workbook
    with no file, one writable already asked to be, and any other mode fail."""
    if mode is MISSING:
        raise error(449, "Argument not optional")
    try:
        wanted = int(to_integer(mode, "Long"))
    except Exception as failure:
        raise failed("ChangeFileAccess") from failure
    if not book.path or wanted not in (2, 3) or wanted == 2 and not book.read_only:
        raise failed("ChangeFileAccess")
    if wanted == 2:
        raise VBAUnsupportedError("ChangeFileAccess back to read and write, for which Excel reopens the file, is not "
                                  "implemented")
    book.read_only = True
    book.saved = True
    return True


def user_status(book: Workbook) -> object:
    """Workbook.UserStatus: this user, the time and xlUserStatus's exclusive, in an array from 1 by 1 to 3. A workbook opened read-only has
    no user holding it, and Excel answers 1004: new, saved and opened read-write it answers (tests/fixtures/workbook_reads.json)."""
    from datetime import datetime

    if book.read_only:
        raise error(1004, _APP_DEFINED)
    now = datetime.now().strftime("%m/%d/%Y %H:%M:%S")
    return VBAArray([(1, 1), (1, 3)], items=[book.application.user_name, now, VBAInt(1, "Long")])


class WebOptions(_Owned):
    vba_type_name = "WebOptions"

    def __init__(self, book: Workbook) -> None:
        super().__init__(book, book)

    @member
    def Encoding(self) -> object:
        return VBAInt(1252, "Long")

    @member
    def PixelsPerInch(self) -> object:
        return VBAInt(96, "Long")

    @member
    def RelyOnCSS(self) -> object:
        return True

    @member
    def OrganizeInFolder(self) -> object:
        return True

    @member
    def UseLongFileNames(self) -> object:
        return True

    @member
    def DownloadComponents(self) -> object:
        return False

    @member
    def RelyOnVML(self) -> object:
        return False

    @member
    def AllowPNG(self) -> object:
        return True

    @member
    def ScreenSize(self) -> object:
        return VBAInt(4, "Long")

    @member
    def TargetBrowser(self) -> object:
        return VBAInt(4, "Long")

    @member
    def LocationOfComponents(self) -> object:
        return ""

    @member
    def FolderSuffix(self) -> object:
        return "_files"


class _Counted(VBACollection, _Owned):
    """A workbook's collection of objects the model holds none of."""

    by_index = 9

    def __init__(self, book: Workbook) -> None:
        _Owned.__init__(self, book, book)

    def vba_items(self) -> list[object]:
        return []

    def vba_lookup(self, index: object, items: list[object]) -> object:
        raise error(self.by_index)


class ModelTables(_Counted):
    vba_type_name = "ModelTables"


class ModelRelationships(_Counted):
    vba_type_name = "ModelRelationships"


class SignatureSet(_Counted):
    vba_type_name = "SignatureSet"
    by_index = 440


class Model(_Owned):
    """The workbook's data model, which holds nothing in a workbook the model reads."""

    vba_type_name = "Model"

    def __init__(self, book: Workbook) -> None:
        super().__init__(book, book)

    @member
    def ModelTables(self) -> object:
        return ModelTables(self.book)

    @member
    def ModelRelationships(self) -> object:
        return ModelRelationships(self.book)


class SmartDocument(_Owned):
    vba_type_name = "SmartDocument"

    def __init__(self, book: Workbook) -> None:
        super().__init__(book, book)


class DocumentInspectors(_Owned):
    vba_type_name = "DocumentInspectors"

    def __init__(self, book: Workbook) -> None:
        super().__init__(book, book)

    @member
    def Count(self) -> object:
        return VBAInt(5, "Long")


class Graphic(ExcelObject):
    """A picture in a header or a footer: of a sheet with none, every member fails."""

    vba_type_name = "Graphic"

    def __init__(self, page_setup: PageSetup) -> None:
        self.page_setup = page_setup

    @member
    def Parent(self) -> object:
        return self.page_setup

    @member
    def Filename(self) -> object:
        raise app_defined()

    @setter("Filename")
    def _set_filename(self, value: object) -> None:
        from pathlib import Path

        self.page_setup.sheet.book.application.outside("A picture in a header or a footer")
        if Path(to_text(value)).is_file():
            raise VBAUnsupportedError("a picture in a header or a footer is not implemented")
        raise app_defined()

    @member
    def Height(self) -> object:
        raise app_defined()

    @member
    def Width(self) -> object:
        raise app_defined()

    @member
    def LockAspectRatio(self) -> object:
        raise app_defined()
