"""Members of Workbook, Worksheet and PageSetup a macro meets with no screen, server or mail system: the theme's colours
and fonts, custom XML parts, read-only access, and what answers the same on any workbook that shares and reviews nothing.

Measured in live Excel (scripts/measure_workbook_members.py, tests/fixtures/workbook_members.json).

A change to a theme's colour reaches every cell using it and is saved. A change to a theme's font, a custom XML part
added and then saved, and going back to read and write access, where Excel reopens the file, report themselves.
"""
from __future__ import annotations

import re
import uuid
from typing import TYPE_CHECKING, ClassVar

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

    def vba_default_call(self, args: list[object], named: dict[str, object]) -> object:
        """An index straight on the scheme, as `ThemeColorScheme(5).RGB` writes it, is Colors(5); one out of range is -2147024809 there,
        where Colors says 5 (tests/fixtures/theme_scheme_default.json, tests/fixtures/workbook_members.json)."""
        if len(args) != 1 or named:
            raise error(450, "Wrong number of arguments or invalid property assignment")
        place = int(to_integer(args[0], "Long"))
        if not 1 <= place <= len(_SCHEME):
            raise error(-2147024809, "The parameter is incorrect.")
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
    """A custom XML part: its tree, what reads and selects in it, and the prefixes its XPath knows
    (see _custom_xml)."""

    vba_type_name = "CustomXMLPart"
    vba_library = "office"
    #: Its node events, which no edit a macro makes raises, as Excel raises none (tests/fixtures/custom_xml_c.json):
    #: a WithEvents variable binds to them and hears nothing.
    raised_events: ClassVar[frozenset[str]] = frozenset({"NodeAfterDelete", "NodeAfterInsert", "NodeAfterReplace"})

    def vba_no_member(self, name: str) -> Exception:
        from pyopenvba.apps.excel._custom_xml import no_member

        return no_member(self, name) or super().vba_no_member(name)

    def __init__(self, owner: object, book: Workbook, text: str | None, built_in: bool, *, namespace: str = "",
                 identity: str = "") -> None:
        from pyopenvba.apps.excel._custom_xml import mappings_for, namespaces
        from pyopenvba.apps.excel._custom_xml_tree import parsed

        super().__init__(owner, book)
        self.built_in, self.built_in_namespace = built_in, namespace
        self.document = parsed(text) if text else None
        self.mappings: list[tuple[str, str]] = [] if built_in else mappings_for(self.document)
        #: The namespaces of the document as it was loaded, which its schemas follow (see _custom_xml.namespaces).
        self.schemas: list[str] = namespaces(self.document)
        self.identity = identity or "{" + str(uuid.uuid4()).upper() + "}"

    @property
    def namespace(self) -> str:
        if self.built_in:
            return self.built_in_namespace
        root = self.document.root() if self.document is not None else None
        return root.uri if root is not None else ""

    def edited(self) -> None:
        self.book.custom_xml_changed = True
        self.book.saved = False

    def _own(self) -> None:
        if self.built_in:
            raise VBAUnsupportedError("the nodes of a workbook's built-in custom XML parts are not implemented")

    @member
    def BuiltIn(self) -> object:
        return self.built_in

    @member
    def NamespaceURI(self) -> object:
        return self.namespace

    @member
    def XML(self) -> object:
        from pyopenvba.apps.excel._custom_xml_tree import serialized

        self._own()
        return "" if self.document is None else serialized(self.document)

    @member
    def ID(self) -> object:
        return self.identity

    @member
    def DocumentElement(self) -> object:
        from pyopenvba.apps.excel._custom_xml import view

        self._own()
        return view(self, None if self.document is None else self.document.root())

    @member
    def NamespaceManager(self) -> object:
        from pyopenvba.apps.excel._custom_xml import CustomXMLPrefixMappings

        self._own()
        return CustomXMLPrefixMappings(self)

    @member
    def Errors(self) -> object:
        from pyopenvba.apps.excel._custom_xml import CustomXMLValidationErrors

        return CustomXMLValidationErrors(self)

    @member
    def SchemaCollection(self) -> object:
        from pyopenvba.apps.excel._custom_xml import CustomXMLSchemaCollection

        return CustomXMLSchemaCollection(self)

    @setter("SchemaCollection", objects=True)
    def _set_schema_collection(self, value: object) -> None:
        # Measured late bound: Set is error 438 (tests/fixtures/custom_xml_c.json).
        raise error(438, "Object doesn't support this property or method")

    @method
    def SelectNodes(self, XPath: object = MISSING) -> object:
        from pyopenvba.apps.excel._custom_xml import CustomXMLNodes, select
        from pyopenvba.apps.excel._custom_xml_tree import Node

        self._own()
        return CustomXMLNodes(self, select(self, self.document or Node(9), XPath))

    @method
    def SelectSingleNode(self, XPath: object = MISSING) -> object:
        from pyopenvba.apps.excel._custom_xml import select, view
        from pyopenvba.apps.excel._custom_xml_tree import Node

        self._own()
        found = select(self, self.document or Node(9), XPath)
        return view(self, found[0] if found else None)

    @method
    def AddNode(self, Parent: object = MISSING, Name: object = MISSING, NamespaceURI: object = MISSING,
                NextSibling: object = MISSING, NodeType: object = MISSING, NodeValue: object = MISSING) -> object:
        """A node made under Parent. Late bound, a NextSibling left out is error 13
        (tests/fixtures/custom_xml.json); typed, it puts the node last."""
        from pyopenvba.apps.excel._custom_xml import CustomXMLNode, add
        from pyopenvba.interpreter._objects import BINDING

        self._own()
        if not isinstance(Parent, CustomXMLNode) or Parent.part is not self or (BINDING.late and NextSibling is MISSING):
            raise error(13, "Type mismatch")
        before = None
        if isinstance(NextSibling, CustomXMLNode):
            if NextSibling.node.parent is not Parent.node:
                raise VBAUnsupportedError("a NextSibling that is not one of the parent's children is not implemented")
            before = NextSibling.node
        add(self, Parent.node, Name, NamespaceURI, before, NodeType, NodeValue)
        return None

    def _loaded(self, text: str) -> object:
        from pyopenvba.apps.excel._custom_xml import mappings_for, namespaces
        from pyopenvba.apps.excel._custom_xml_tree import parsed

        if self.document is not None:
            raise error(-536602616, "This custom XML part has already been loaded.")
        document = parsed(text)
        mappings = mappings_for(document)
        self.document, self.mappings, self.schemas = document, mappings, namespaces(document)
        self.edited()
        return True

    @method
    def LoadXML(self, XML: object = MISSING) -> object:
        self._own()
        return self._loaded(to_text(XML))

    @method
    def Load(self, FilePath: object = MISSING) -> object:
        """The part read from a file, through the sandbox's gate; a file not there is error 432."""
        import os

        from pyopenvba.interpreter._files import final_path

        self._own()
        self.book.application.outside("CustomXMLPart.Load")
        if self.document is not None:
            raise error(-536602616, "This custom XML part has already been loaded.")
        path = final_path(to_text(FilePath))
        if not os.path.isfile(path):
            raise error(432, "File name or class name not found during Automation operation")
        with open(path, "rb") as source:
            data = source.read()
        return self._loaded(data.decode("utf-8-sig", errors="replace"))

    @method
    def Delete(self) -> object:
        if self.built_in:
            raise error(-536602603, "This custom XML part is required for Office to function properly and cannot be "
                                    "deleted.")
        parts = _custom_parts(self.book)
        if self in parts:
            _raised(self.book, "PartBeforeDelete", [self])
            parts.remove(self)
            self.edited()
        return None


def _raised(book: Workbook, event: str, args: list[object]) -> None:
    """One of CustomXMLParts' events, for a macro listening to the workbook's parts (tests/fixtures/custom_xml_c.json):
    Application.EnableEvents has no say in it."""
    interpreter = book.application.interpreter
    if interpreter is not None:
        interpreter.raise_event(parts_of(book), event, args)


def parts_of(book: Workbook) -> CustomXMLParts:
    """Workbook.CustomXMLParts, the same object each time, so a WithEvents variable holding it hears its events."""
    if book.custom_xml_view is None:
        book.custom_xml_view = CustomXMLParts(book)
    return book.custom_xml_view


def _custom_parts(book: Workbook) -> list[CustomXMLPart]:
    """The workbook's parts, its three built in first and then those its file holds, each with the ID its properties
    give it."""
    parts = book.custom_xml_parts_
    if parts is None:
        parts = [CustomXMLPart(book, book, None, True, namespace=namespace) for namespace in _BUILT_IN_XML]
        package = book.package
        if package is not None:
            items = sorted((name for name in package.names() if re.fullmatch(r"customXml/item\d+\.xml", name)),
                           key=lambda name: int(re.findall(r"\d+", name)[0]))
            for name in items:
                text = package.read(name).decode("utf-8-sig", errors="replace")
                properties = name.replace("/item", "/itemProps")
                identity = ""
                if package.has(properties):
                    found = re.search(r'itemID="([^"]+)"', package.read(properties).decode("utf-8", errors="replace"))
                    identity = found.group(1) if found else ""
                parts.append(CustomXMLPart(book, book, text, False, identity=identity))
        book.custom_xml_parts_ = parts
    return parts


class CustomXMLParts(VBACollection, _Owned):
    vba_type_name = "CustomXMLParts"
    vba_library = "office"
    raised_events: ClassVar[frozenset[str]] = frozenset({"PartAfterAdd", "PartAfterLoad", "PartBeforeDelete"})

    def vba_no_member(self, name: str) -> Exception:
        from pyopenvba.apps.excel._custom_xml import no_member

        return no_member(self, name) or super().vba_no_member(name)

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
        """A new part, after which the workbook's parts raise PartAfterAdd and PartAfterLoad, as Excel's do even for
        an empty part (tests/fixtures/custom_xml_c.json)."""
        if SchemaCollection is not MISSING:
            raise VBAUnsupportedError("CustomXMLParts.Add with a schema collection is not implemented")
        text = "" if XML is MISSING else to_text(XML)
        made = CustomXMLPart(self.book, self.book, text, False)
        _custom_parts(self.book).append(made)
        made.edited()
        _raised(self.book, "PartAfterAdd", [made])
        _raised(self.book, "PartAfterLoad", [made])
        return made

    @method
    def SelectByNamespace(self, NamespaceURI: object = MISSING) -> object:
        wanted = to_text(NamespaceURI)
        return CustomXMLParts(self.book, [part for part in _custom_parts(self.book) if part.namespace == wanted])

    @method
    def SelectByID(self, Id: object = MISSING) -> object:
        wanted = to_text(Id)
        return next((part for part in _custom_parts(self.book) if part.identity == wanted), NOTHING)


def write_parts(book: Workbook, package: OpcFile) -> None:
    """The custom XML parts a macro added, loaded, changed or deleted, written (see _custom_xml)."""
    from pyopenvba.apps.excel._custom_xml import write

    write(book, package)


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


#: The cryptographic providers Windows ships, and the algorithm names SetPasswordEncryptionOptions takes for each, as Excel answered once on
#: this machine (tests/fixtures/workbook_encryption_options.json). No live probe names a provider again (tests/test_probes_leave_crypto_alone.py).
_PROVIDERS: dict[str, frozenset[str]] = {
    "Microsoft Base Cryptographic Provider v1.0": frozenset({"RC4", "RC2", "DES"}),
    "Microsoft Enhanced Cryptographic Provider v1.0": frozenset({"RC4", "RC2", "DES", "3DES"}),
    "Microsoft Strong Cryptographic Provider": frozenset({"RC4", "RC2", "DES", "3DES"}),
    "Microsoft RSA SChannel Cryptographic Provider": frozenset({"RC4", "RC2", "DES", "3DES", "AES 128", "AES 256"}),
    "Microsoft Enhanced RSA and AES Cryptographic Provider": frozenset({"RC4", "RC2", "DES", "3DES", "AES 128", "AES 192", "AES 256"}),
    "Microsoft Base DSS Cryptographic Provider": frozenset(),
    "Microsoft Base DSS and Diffie-Hellman Cryptographic Provider": frozenset({"RC4", "RC2", "DES"}),
    "Microsoft Enhanced DSS and Diffie-Hellman Cryptographic Provider": frozenset({"RC4", "RC2", "DES", "3DES"}),
    "Microsoft DH SChannel Cryptographic Provider": frozenset({"RC4", "RC2", "DES", "3DES"}),
}
#: A provider whose answers depend on a smartcard: Excel asked the machine for one, and what it answered is not a rule.
_HARDWARE = "Microsoft Base Smart Card Crypto Provider"


def encryption_options(held: tuple[str, str, int, bool], provider: object, algorithm: object, length: object,
                       properties: object) -> tuple[str, str, int, bool]:
    """Workbook.SetPasswordEncryptionOptions: the options it sets, or the error Excel gives, keeping what the book held. Text for the key length
    or the file properties is 13. A provider or algorithm left out, unknown, or spelled in another case is 5, as is an algorithm the provider
    does not offer; a leading space before the algorithm is dropped. A key length is 0 or a positive multiple of 8, or 5. Typed and late
    answer the same (tests/fixtures/workbook_encryption_options.json)."""
    from pyopenvba.interpreter._values import to_bool

    bits = 0 if length is MISSING else int(to_integer(length, "Long"))
    files = held[3] if properties is MISSING else to_bool(properties)
    name = provider if isinstance(provider, str) else None
    if name == _HARDWARE:
        raise VBAUnsupportedError(f"SetPasswordEncryptionOptions with {_HARDWARE}, which asks the machine for a smartcard, is not implemented")
    method = algorithm.lstrip(" ") if isinstance(algorithm, str) else None
    if name not in _PROVIDERS or method not in _PROVIDERS[name] or bits < 0 or bits % 8:
        raise error(5, "Invalid procedure call or argument")
    return name, method, bits, files
