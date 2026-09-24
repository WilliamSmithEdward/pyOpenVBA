"""Shared VBA library reference records and file-facade operations."""
from __future__ import annotations

import struct
from dataclasses import dataclass
from uuid import UUID

from pyopenvba.exceptions import VBAProjectError
from pyopenvba.vba import VBAReference, encoding_for_codepage

MSFORMS_GUID = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}"
#: The Microsoft Forms library as the editor's reference names it, path included, whatever FM20.DLL's real
#: location: the host resolves the library by GUID and version (tests/fixtures/form_reference.json).
MSFORMS_LIBID = ("*\\G{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0#C:\\WINDOWS\\system32\\FM20.DLL"
                 "#Microsoft Forms 2.0 Object Library")
#: A control reference's twiddled libid, which is the null one.
_NULL_LIBID = "*\\G{00000000-0000-0000-0000-000000000000}#0.0#0##"
VBA_GUID = "{000204EF-0000-0000-C000-000000000046}"


@dataclass(frozen=True)
class Library:
    name: str
    guid: str
    major: int
    minor: int
    path: str
    description: str


_OFFICE = "C:\\Program Files\\Microsoft Office\\root\\Office16\\"
LIBRARIES = {
    "excel": Library("Excel", "{00020813-0000-0000-C000-000000000046}", 1, 9,
                     _OFFICE + "EXCEL.EXE", "Microsoft Excel 16.0 Object Library"),
    "word": Library("Word", "{00020905-0000-0000-C000-000000000046}", 8, 7,
                    _OFFICE + "MSWORD.OLB", "Microsoft Word 16.0 Object Library"),
    "powerpoint": Library("PowerPoint", "{91493440-5A91-11CF-8700-00AA0060263B}", 2, 0x12,
                          _OFFICE + "MSPPT.OLB", "Microsoft PowerPoint 16.0 Object Library"),
    "access": Library("Access", "{4AFFC9A0-5F99-101B-AF4E-00AA003F0F07}", 9, 0,
                      _OFFICE + "MSACC.OLB", "Microsoft Access 16.0 Object Library"),
}


@dataclass
class ReferenceSpan:
    reference: VBAReference
    start: int
    end: int


def reference_spans(raw: bytes) -> list[ReferenceSpan]:
    """Group name/body/extended records, retaining untouched byte spans."""
    result: list[ReferenceSpan] = []
    pos, code_page = 0, 1252
    name = unicode_name = ""
    start: int | None = None
    control: ReferenceSpan | None = None
    while pos + 6 <= len(raw):
        rid, size = struct.unpack_from("<HI", raw, pos)
        if rid == 0x000F:
            return result
        end = pos + (12 if rid == 0x0009 else 6 + size)
        if end > len(raw):
            raise VBAProjectError("truncated reference section")
        data = raw[pos + 6:end]
        encoding = encoding_for_codepage(code_page)
        if rid == 3 and len(data) >= 2:
            code_page = int.from_bytes(data[:2], "little")
        if rid in {0x16, 0x3E}:
            if control is None:
                start = pos if start is None else start
                if rid == 0x16:
                    name = data.decode(encoding, errors="replace")
                else:
                    unicode_name = data.decode("utf-16-le", errors="replace")
            else:
                control.end = end
        elif rid in {0xD, 0xE, 0x33, 0x2F}:
            if rid == 0x2F and control is not None:
                control.end = end
            else:
                libid = data if rid == 0x33 else data[4:4 + int.from_bytes(data[:4], "little")]
                reference = VBAReference(name=unicode_name or name, name_unicode=unicode_name,
                                         kind={0xD: "registered", 0xE: "project", 0x33: "control", 0x2F: "control"}[rid],
                                         libid=libid.decode(encoding, errors="replace"))
                if rid == 0xE:
                    offset = 4 + len(libid)
                    length = int.from_bytes(data[offset:offset + 4], "little")
                    reference.libid_secondary = data[offset + 4:offset + 4 + length].decode(encoding, errors="replace")
                span = ReferenceSpan(reference, pos if start is None else start, end)
                result.append(span)
                control = span if rid in {0x33, 0x2F} else None
            start, name, unicode_name = None, "", ""
        elif rid == 0x30 and control is not None:
            control.end = end
            control = None
        pos = end
    raise VBAProjectError("the dir stream has no PROJECTMODULES record")


def module_offset(raw: bytes) -> int:
    reference_spans(raw)  # validate before making any edit
    pos = 0
    while pos + 6 <= len(raw):
        rid, size = struct.unpack_from("<HI", raw, pos)
        if rid == 0xF:
            return pos
        pos += 12 if rid == 9 else 6 + size
    raise VBAProjectError("the dir stream has no PROJECTMODULES record")


def record(rid: int, body: bytes) -> bytes:
    return struct.pack("<HI", rid, len(body)) + body


def registered_block(reference: VBAReference, code_page: int) -> bytes:
    encoding = encoding_for_codepage(code_page)
    name, libid = reference.name.encode(encoding), reference.libid.encode(encoding)
    return (record(0x16, name) + record(0x3E, reference.name.encode("utf-16-le"))
            + record(0xD, struct.pack("<I", len(libid)) + libid + bytes(6)))


def is_msforms(reference: VBAReference) -> bool:
    """Whether ``reference`` is the Microsoft Forms library, by name or GUID."""
    return reference.name.lower() == "msforms" or reference.guid.upper() == MSFORMS_GUID


def msforms_block(code_page: int) -> bytes:
    """The Microsoft Forms reference the editor declares with a project's first UserForm.

    Excel, Word and PowerPoint write the same control reference
    (tests/fixtures/form_reference.json): the name, the original libid,
    the null twiddled libid, the name again, and the extended record
    with the type library's GUID and cookie 1. Their extended libid
    names the .exd cache the editor keeps in the saving user's Temp
    folder, under a GUID of that cache's own; here it repeats the
    original libid, which the editor compiles against and Excel keeps
    when it saves the file again, so no user's folder is written.
    """
    encoding = encoding_for_codepage(code_page)
    name = record(0x16, "MSForms".encode(encoding)) + record(0x3E, "MSForms".encode("utf-16-le"))
    original, twiddled = MSFORMS_LIBID.encode(encoding), _NULL_LIBID.encode(encoding)
    extended = (struct.pack("<I", len(original)) + original + bytes(6) + UUID(MSFORMS_GUID).bytes_le
                + struct.pack("<I", 1))
    return (name + record(0x33, original) + record(0x2F, struct.pack("<I", len(twiddled)) + twiddled + bytes(6))
            + name + record(0x30, extended))


class ReferenceManager:
    """One public reference surface; subclasses supply only storage hooks."""

    _reference_error: type[Exception] = VBAProjectError

    def _reference_data(self) -> tuple[bytes, int]:
        raise NotImplementedError

    def _write_reference_data(self, raw: bytes) -> None:
        raise NotImplementedError

    def _reference_host(self) -> str:
        raise NotImplementedError

    def _reference_forms(self) -> list[str]:
        return []

    def references(self) -> list[VBAReference]:
        """Declared libraries in priority order; implicit VBA/host libraries are excluded."""
        return [span.reference for span in reference_spans(self._reference_data()[0])]

    def add_reference(self, name: str, guid: str | None = None, major: int = 1, minor: int = 0,
                      *, path: str = "", description: str = "", lcid: int = 0) -> VBAReference:
        """Add an Office library by name, or a custom registered library by GUID.

        An existing GUID is a no-op. Save the file to persist the change.
        Paths are resolution hints; this does not install or load a library.
        """
        name = name.strip()
        if name.lower() in {self._reference_host(), "vba"}:
            raise self._reference_error(f"{name} is an implicit project library")
        known = LIBRARIES.get(name.lower())
        if guid is None:
            if known is None:
                raise self._reference_error(f"unknown library {name!r}; provide its GUID")
            name, guid, major, minor = known.name, known.guid, known.major, known.minor
            path, description = path or known.path, description or known.description
        try:
            guid = "{" + str(UUID(guid)).upper() + "}"
        except ValueError as exc:
            raise self._reference_error(f"invalid library GUID: {guid!r}") from exc
        host = LIBRARIES[self._reference_host()]
        if guid in {host.guid, VBA_GUID} or name.lower() in {host.name.lower(), "vba"}:
            raise self._reference_error(f"{name} is an implicit project library")
        if not name or any(c in name + path + description for c in "\x00\r\n#"):
            raise self._reference_error("invalid reference name, path or description")
        if min(major, minor, lcid) < 0 or max(major, minor) > 65535:
            raise self._reference_error("invalid reference version or locale")
        raw, code_page = self._reference_data()
        for existing in self.references():
            if existing.guid.upper() == guid:
                return existing
            if existing.name.lower() == name.lower():
                raise self._reference_error(f"the project already references {name!r} with a different GUID")
        reference = VBAReference(name=name, name_unicode=name, kind="registered",
                                 libid=f"*\\G{guid}#{major:x}.{minor:x}#{lcid}#{path}#{description}")
        block = registered_block(reference, code_page)
        at = module_offset(raw)
        self._write_reference_data(raw[:at] + block + raw[at:])
        return reference

    def remove_reference(self, name: str) -> bool:
        """Remove a declared reference by name or GUID; missing references are a no-op."""
        wanted = name.strip().lower()
        if not wanted:
            raise self._reference_error("a reference name or GUID is required")
        known = LIBRARIES.get(wanted)
        try:
            guid = "{" + str(UUID(known.guid if known else wanted)) + "}"
        except ValueError:
            guid = ""
        raw, _ = self._reference_data()
        cuts = [span for span in reference_spans(raw)
                if span.reference.name.lower() == wanted or (guid and span.reference.guid.lower() == guid)]
        if not cuts:
            return False
        if any(is_msforms(s.reference) for s in cuts):
            forms = self._reference_forms()
            if forms:
                raise self._reference_error(f"Microsoft Forms is required by UserForms: {', '.join(forms)}")
        parts: list[bytes] = []
        kept = 0
        for span in cuts:
            parts.append(raw[kept:span.start])
            kept = span.end
        parts.append(raw[kept:])
        self._write_reference_data(b"".join(parts))
        return True

    def _ensure_forms_reference(self) -> None:
        """Declare Microsoft Forms as the editor does with a project's first UserForm, unless it is declared.

        remove_reference refuses to take the library out while a form is
        there; this is the other half of that rule.
        """
        raw, code_page = self._reference_data()
        if any(is_msforms(span.reference) for span in reference_spans(raw)):
            return
        at = module_offset(raw)
        self._write_reference_data(raw[:at] + msforms_block(code_page) + raw[at:])

    def drop_reference(self, name: str) -> None:
        """Compatibility spelling: remove, raising when the reference is absent."""
        if not self.remove_reference(name):
            raise self._reference_error(f"the project has no reference named {name!r}")
