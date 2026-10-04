"""Explicit file diagnostics and conservative, content-preserving ZIP repair.

These checks do not execute VBA, verify signatures, or prove that Office can
open a document. Access diagnostics cover the catalog and VBA, not every row
or allocation map; use AccessDatabase.compact_and_repair for compaction.
"""

from __future__ import annotations

# The checker deliberately inspects CFB directory sizes and host format
# metadata, which are internal to the parsers it validates.
# pyright: reportPrivateUsage=false

import io
import posixpath
import struct
import zipfile
import zlib
from dataclasses import dataclass
from pathlib import Path
from typing import Literal, NoReturn
from urllib.parse import unquote, urlsplit
from xml.parsers import expat

from pyopenvba.cfb import CFB
from pyopenvba.exceptions import NoVBAProjectError, PyOpenVBAError

_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships}"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types}"
_DEFAULT_PART = 256 * 1024 * 1024
_DEFAULT_TOTAL = 1024 * 1024 * 1024


@dataclass(frozen=True)
class FileIssue:
    """A stable diagnostic code, severity, location, and explanation."""

    code: str
    severity: Literal["error", "warning"]
    location: str
    message: str
    repairable: bool = False


@dataclass(frozen=True)
class FileCheckReport:
    """Results for the on-disk snapshot, not unsaved in-memory edits.

    ``ok`` requires no errors and complete coverage of the supported checks.
    Warnings explain unsupported coverage. This is not Office certification.
    """

    path: Path
    issues: tuple[FileIssue, ...]
    checked_parts: int
    complete: bool

    @property
    def errors(self) -> tuple[FileIssue, ...]:
        return tuple(issue for issue in self.issues if issue.severity == "error")

    @property
    def warnings(self) -> tuple[FileIssue, ...]:
        return tuple(issue for issue in self.issues if issue.severity == "warning")

    @property
    def ok(self) -> bool:
        return self.complete and not self.errors


@dataclass(frozen=True)
class FileRepairResult:
    """The original findings, changes actually made, and verified output."""

    output: Path
    before: FileCheckReport
    after: FileCheckReport
    changes: tuple[FileIssue, ...]


class FileRepairError(PyOpenVBAError):
    """Repair refused because content or unsupported structures need guessing."""

    def __init__(self, message: str, report: FileCheckReport) -> None:
        super().__init__(message)
        self.report = report


def _xml(data: bytes) -> tuple[str, list[tuple[str, dict[str, str]]]]:
    """Parse well-formed XML without DTDs or entity expansion; retain children."""
    parser = expat.ParserCreate(namespace_separator="}")
    root = ""
    depth = 0
    children: list[tuple[str, dict[str, str]]] = []

    def start(name: str, attrs: dict[str, str]) -> None:
        nonlocal root, depth
        if depth == 0:
            root = name
        elif depth == 1:
            children.append((name, attrs))
        depth += 1

    def end(_name: str) -> None:
        nonlocal depth
        depth -= 1

    def refuse_dtd(*_args: object) -> NoReturn:
        raise ValueError("DTDs and declared entities are not supported")

    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.StartDoctypeDeclHandler = refuse_dtd
    parser.EntityDeclHandler = refuse_dtd
    parser.ExternalEntityRefHandler = refuse_dtd
    parser.SetParamEntityParsing(expat.XML_PARAM_ENTITY_PARSING_NEVER)
    parser.Parse(data, True)
    return root, children


def _verify_deflate(raw: bytes, info: zipfile.ZipInfo, limit: int) -> None:
    """Do not trust file_size: ZipExtFile can trim output to that size."""
    name_size, extra_size = struct.unpack_from("<HH", raw, info.header_offset + 26)
    start = info.header_offset + 30 + name_size + extra_size
    end = start + info.compress_size
    if end > len(raw):
        raise ValueError("Truncated compressed payload")
    decoder = zlib.decompressobj(-15)
    size = 0
    crc = 0
    for offset in range(start, end, 1024 * 1024):
        chunk = raw[offset:min(end, offset + 1024 * 1024)]
        while chunk:
            output = decoder.decompress(chunk, min(1024 * 1024, limit - size + 1))
            size += len(output)
            if size > limit:
                raise ValueError("Actual decompressed payload exceeds byte limit")
            crc = zlib.crc32(output, crc)
            chunk = decoder.unconsumed_tail
            if decoder.unused_data:
                raise ValueError("Trailing bytes within compressed payload")
    if not decoder.eof or size != info.file_size or crc != info.CRC:
        raise ValueError("Actual deflate payload size or CRC disagrees with central directory")


def _cfb_check(raw: bytes) -> None:
    cfb = CFB.from_bytes(raw)
    entries = cfb._directory
    if not entries or entries[0].obj_type != 5 or sum(entry.obj_type == 5 for entry in entries) != 1:
        raise ValueError("CFB directory must have one root entry at index zero")
    active = {index for index, entry in enumerate(entries) if entry.obj_type in {1, 2, 5}}
    states: dict[int, int] = {}
    # The reader tolerates broken directory pointers; diagnostics must not.
    for first in active:
        stack = [(first, False)]
        while stack:
            index, leaving = stack.pop()
            if leaving:
                states[index] = 2
                continue
            if states.get(index) == 1:
                raise ValueError("Cycle in CFB directory pointers")
            if states.get(index) == 2:
                continue
            states[index] = 1
            stack.append((index, True))
            entry = entries[index]
            for pointer in (entry.left_sibling_id, entry.right_sibling_id, entry.child_id):
                if pointer == 0xFFFFFFFF:
                    continue
                if pointer not in active:
                    raise ValueError("CFB directory pointer references a missing entry")
                stack.append((pointer, False))
    # Include orphan directory entries, and compare against declared sizes:
    # the normal reader intentionally returns the bytes it can retrieve.
    for entry in cfb._directory:
        if entry.obj_type == 2 and len(cfb._read_stream(entry)) != entry.size:
            raise ValueError(f"Truncated CFB stream: {entry.name}")
    pending: list[tuple[str, ...]] = [()]
    seen: set[tuple[str, ...]] = set()
    # Bound traversal even if a malformed directory cycles through storages.
    while pending:
        path = pending.pop()
        if path in seen or len(path) > 128 or len(seen) > 100_000:
            raise ValueError("CFB directory traversal is cyclic or exceeds limits")
        seen.add(path)
        streams = cfb.list_streams_at(path)
        storages = cfb.list_storages_at(path)
        names = [name.casefold() for name in streams + storages]
        if len(names) != len(set(names)):
            raise ValueError("Duplicate CFB directory names within one storage")
        for name in streams:
            cfb.get_stream_at(path, name)
        pending.extend((*path, name) for name in storages)


def _host_check(path: Path, raw: bytes, issues: list[FileIssue]) -> bool:
    # Imports here keep the diagnostics module independent of the public facade.
    from pyopenvba.excel import ExcelFile
    from pyopenvba.powerpoint import PowerPointFile
    from pyopenvba.word import WordFile
    from pyopenvba.vba import parse_vba_project

    plain_formats = {ExcelFile: {".xlsx", ".xltx"}, WordFile: {".docx", ".dotx"},
                     PowerPointFile: {".pptx", ".potx", ".ppsx"}}
    for host_type in (ExcelFile, WordFile, PowerPointFile):
        if path.suffix.lower() in host_type._zip_formats | host_type._cfb_formats | plain_formats[host_type]:
            # Use the immutable snapshot, without reopening a potentially changed file.
            if raw.startswith(b"PK"):
                with zipfile.ZipFile(io.BytesIO(raw)) as archive:
                    if host_type._vba_entry not in archive.namelist():
                        return True
                    project_raw = archive.read(host_type._vba_entry)
            else:
                project_raw = raw
                if host_type is PowerPointFile:
                    from pyopenvba._ppt_container import extract_vba_storage
                    try:
                        project_raw = extract_vba_storage(CFB.from_bytes(raw))
                    except NoVBAProjectError:
                        return True
            cfb = CFB.from_bytes(project_raw)
            if not raw.startswith(b"PK") and host_type._project_storage is not None:
                if host_type._project_storage.casefold() not in {name.casefold() for name in cfb.list_storages_at()}:
                    return True
            _cfb_check(project_raw)
            project = parse_vba_project(cfb)
            for module in project.modules:
                # Source is lazy: force every compressed chunk to be decoded.
                _ = module.source
            from pyopenvba.forms import read_forms
            root = () if raw.startswith(b"PK") or host_type._project_storage is None else (host_type._project_storage,)
            read_forms(cfb, code_page=project.code_page, root=root)
            for message in project.validate(cfb):
                issues.append(FileIssue("vba.structure", "error", path.name, message))
            return True
    return False


def _package_xml(parts: dict[str, bytes], names: set[str], issues: list[FileIssue]) -> None:
    parsed: dict[str, tuple[str, list[tuple[str, dict[str, str]]]]] = {}
    for name, data in parts.items():
        if name.lower().endswith((".xml", ".rels")):
            try:
                parsed[name] = _xml(data)
            except (expat.ExpatError, ValueError) as exc:
                issues.append(FileIssue("opc.xml", "error", name, str(exc)))
    if "[Content_Types].xml" not in names:
        issues.append(FileIssue("opc.content_types", "error", "[Content_Types].xml", "Missing content types part"))
    if "_rels/.rels" not in names:
        issues.append(FileIssue("opc.relationships", "error", "_rels/.rels", "Missing package relationships part"))
    for name, (root, children) in parsed.items():
        if name.endswith(".rels"):
            if root != _REL_NS + "Relationships":
                issues.append(FileIssue("opc.relationships", "error", name, "Invalid relationships root"))
                continue
            if name == "_rels/.rels":
                folder = ""
                main = [attrs for tag, attrs in children if tag == _REL_NS + "Relationship" and attrs.get("Type") in {
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                    "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument",
                }]
                if len(main) != 1 or main[0].get("TargetMode", "Internal") != "Internal":
                    issues.append(FileIssue("opc.main_relationship", "error", name, "Expected one internal Office document relationship"))
            else:
                parent, leaf = posixpath.split(name)
                if posixpath.basename(parent) != "_rels":
                    issues.append(FileIssue("opc.relationships", "error", name, "Invalid relationships part path"))
                    continue
                folder = posixpath.dirname(parent)
                source = posixpath.join(folder, leaf[:-5])
                if source not in names:
                    issues.append(FileIssue("opc.relationship_source", "error", name, f"Missing source part: {source}"))
            ids: set[str] = set()
            for tag, attrs in children:
                if tag != _REL_NS + "Relationship":
                    issues.append(FileIssue("opc.relationships", "error", name, "Unexpected relationship element"))
                    continue
                identity = attrs.get("Id", "")
                target = attrs.get("Target", "")
                if not identity or identity in ids or not target or not attrs.get("Type"):
                    issues.append(FileIssue("opc.relationships", "error", name, "Missing attributes or duplicate relationship ID"))
                ids.add(identity)
                if attrs.get("TargetMode") == "External":
                    continue
                if attrs.get("TargetMode", "Internal") != "Internal":
                    issues.append(FileIssue("opc.relationships", "error", name, "Invalid TargetMode"))
                uri = urlsplit(target)
                resolved = posixpath.normpath(posixpath.join(folder, unquote(uri.path)))
                if uri.path.startswith("/"):
                    resolved = posixpath.normpath(unquote(uri.path).lstrip("/"))
                if uri.scheme or uri.netloc or resolved.startswith("../") or resolved not in names:
                    issues.append(FileIssue("opc.relationship_target", "error", name, f"Missing or invalid internal target: {target}"))
        elif name == "[Content_Types].xml":
            if root != _CT_NS + "Types":
                issues.append(FileIssue("opc.content_types", "error", name, "Invalid content types root"))
                continue
            defaults: set[str] = set()
            overrides: set[str] = set()
            for tag, attrs in children:
                if tag == _CT_NS + "Default":
                    key = attrs.get("Extension", "").lower()
                    seen = defaults
                elif tag == _CT_NS + "Override":
                    key = unquote(attrs.get("PartName", "")).lstrip("/")
                    seen = overrides
                    if key not in names:
                        issues.append(FileIssue("opc.content_types", "error", name, f"Override refers to missing part: {key}"))
                else:
                    issues.append(FileIssue("opc.content_types", "error", name, "Unexpected content type element"))
                    continue
                if not key or key in seen or not attrs.get("ContentType"):
                    issues.append(FileIssue("opc.content_types", "error", name, "Missing attributes or duplicate content type"))
                seen.add(key)
            for part in names - {"[Content_Types].xml"}:
                if not part.endswith("/") and part not in overrides and part.rsplit(".", 1)[-1].lower() not in defaults:
                    issues.append(FileIssue("opc.content_type_missing", "error", part, "No content type declaration"))


def _check(raw: bytes, path: Path, max_part_bytes: int, max_total_bytes: int) -> FileCheckReport:
    issues: list[FileIssue] = []
    complete = True
    checked = 0
    try:
        if raw.startswith(b"PK"):
            parts: dict[str, bytes] = {}
            total = 0
            with zipfile.ZipFile(io.BytesIO(raw)) as archive:
                infos = archive.infolist()
                names = {info.filename for info in infos}
                if len(names) != len(infos):
                    issues.append(FileIssue("zip.duplicate", "error", path.name, "Duplicate ZIP part names"))
                for info in infos:
                    name = info.filename
                    if (info.orig_filename != name or "\\" in name or name.startswith("/")
                            or posixpath.normpath(name.rstrip("/")) != name.rstrip("/")):
                        issues.append(FileIssue("opc.part_name", "error", name, "Ambiguous or invalid package part name"))
                    if info.compress_type not in {zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED} or info.flag_bits & 1:
                        issues.append(FileIssue("check.compression", "warning", name, "Encrypted or unsupported compression; payload not checked"))
                        complete = False
                        continue
                    if info.file_size > max_part_bytes or total + info.file_size > max_total_bytes:
                        issues.append(FileIssue("check.limit", "warning", name, "Decompression byte limit exceeded"))
                        complete = False
                        continue
                    total += info.file_size
                    try:
                        # Read in bounded chunks; do not trust declared uncompressed sizes.
                        chunks: list[bytes] = []
                        count = 0
                        with archive.open(info) as stream:
                            while True:
                                chunk = stream.read(min(1024 * 1024, max_part_bytes - count + 1))
                                if not chunk:
                                    break
                                count += len(chunk)
                                if count > max_part_bytes:
                                    raise ValueError("Decompressed part exceeds byte limit")
                                chunks.append(chunk)
                        data = b"".join(chunks)
                        if info.compress_type == zipfile.ZIP_DEFLATED:
                            _verify_deflate(raw, info, max_part_bytes)
                        elif info.compress_type == zipfile.ZIP_STORED and info.compress_size != info.file_size:
                            raise ValueError("Stored payload compressed and uncompressed sizes disagree")
                        checked += 1
                        # Keep only XML for cross-part checks; binary payloads need no tree.
                        if name.lower().endswith((".xml", ".rels")):
                            parts[name] = data
                        if data.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
                            _cfb_check(data)
                        header = raw[info.header_offset:info.header_offset + 30]
                        if len(header) != 30 or header[:4] != b"PK\x03\x04":
                            raise ValueError("Missing local ZIP header")
                        flags, method = struct.unpack_from("<HH", header, 6)
                        crc, compressed, size = struct.unpack_from("<III", header, 14)
                        name_size, extra_size = struct.unpack_from("<HH", header, 26)
                        extra_start = info.header_offset + 30 + name_size
                        extra = raw[extra_start:extra_start + extra_size]
                        extra_offset = 0
                        zip64 = False
                        while extra_offset < len(extra):
                            if extra_offset + 4 > len(extra):
                                raise ValueError("Truncated local ZIP extra-field header")
                            kind, length = struct.unpack_from("<HH", extra, extra_offset)
                            extra_offset += 4 + length
                            if extra_offset > len(extra):
                                raise ValueError("Truncated local ZIP extra-field payload")
                            zip64 |= kind == 1
                        if flags != info.flag_bits or method != info.compress_type:
                            issues.append(FileIssue("zip.local_header", "error", name, "Local flags or compression method disagree with central directory"))
                        if flags & 8:
                            if crc not in {0, info.CRC} or compressed not in {0, 0xFFFFFFFF, info.compress_size} or size not in {0, 0xFFFFFFFF, info.file_size}:
                                issues.append(FileIssue("zip.local_header", "error", name, "Local data-descriptor metadata disagrees with central directory"))
                            offset = info.header_offset + 30 + name_size + extra_size + info.compress_size
                            expected = struct.pack("<III", info.CRC, info.compress_size, info.file_size) if max(info.compress_size, info.file_size) <= 0xFFFFFFFF else b""
                            if zip64 or compressed == 0xFFFFFFFF or size == 0xFFFFFFFF or not expected:
                                issues.append(FileIssue("check.zip64", "warning", name, "ZIP64 payload checked; data descriptor not checked"))
                                complete = False
                            elif raw[offset:offset + 12] != expected and raw[offset:offset + 16] != b"PK\x07\x08" + expected:
                                issues.append(FileIssue("zip.descriptor", "error", name, "Missing or inconsistent ZIP data descriptor"))
                        else:
                            if zip64 or compressed == 0xFFFFFFFF or size == 0xFFFFFFFF:
                                issues.append(FileIssue("check.zip64", "warning", name, "ZIP64 payload checked; local size metadata not checked"))
                                complete = False
                            elif (crc, compressed, size) != (info.CRC, info.compress_size, info.file_size):
                                issues.append(FileIssue("zip.local_metadata", "error", name, "Local CRC or sizes disagree with verified central directory", True))
                    except Exception as exc:
                        issues.append(FileIssue("zip.part", "error", name, str(exc)))
                _package_xml(parts, names, issues)
            if complete:
                if not _host_check(path, raw, issues):
                    issues.append(FileIssue("check.vba_scope", "warning", path.name, "VBA validation is unavailable for this extension"))
                    complete = False
        elif raw.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
            _cfb_check(raw)
            checked = 1
            if not _host_check(path, raw, issues):
                issues.append(FileIssue("check.vba_scope", "warning", path.name, "CFB streams checked; VBA validation is unavailable for this extension"))
                complete = False
        elif path.suffix.lower() in {".accdb", ".mdb"}:
            issues.append(FileIssue("check.access_scope", "warning", path.name, "Access allocation maps and table rows are not covered; use the Access APIs for database inspection and compaction"))
            complete = False
            from pyopenvba.access import AccessDatabase
            with AccessDatabase(raw) as database:
                database.table_names(include_system=True)
                if database.has_vba_project():
                    for module in database.vba_project().modules:
                        _ = module.source
                checked = 1
        else:
            issues.append(FileIssue("file.format", "error", path.name, "Not a supported Office ZIP, CFB, or Access container"))
            complete = False
    except Exception as exc:
        issues.append(FileIssue("file.parse", "error", path.name, str(exc)))
        complete = False
    return FileCheckReport(path, tuple(issues), checked, complete)


def _snapshot(path: Path, max_part_bytes: int, max_total_bytes: int) -> bytes:
    if max_part_bytes <= 0 or max_total_bytes <= 0:
        raise ValueError("Byte limits must be positive")
    with path.open("rb") as stream:
        raw = stream.read(max_total_bytes + 1)
    if len(raw) > max_total_bytes:
        raise ValueError("Input file exceeds max_total_bytes")
    return raw


def check_file(path: str | Path, *, max_part_bytes: int = _DEFAULT_PART,
               max_total_bytes: int = _DEFAULT_TOTAL) -> FileCheckReport:
    """Check a disk snapshot without modifying it or executing VBA.

    ZIP parts are decompressed and CRC checked, XML rejects DTDs, internal
    relationships and content types are checked, and supported VBA projects
    are parsed and validated. Legacy CFB streams are read. Access coverage
    is partial and explicitly reported. Missing files and invalid limits raise;
    malformed contents are returned as diagnostics. Limits can be raised for
    trusted large files; they bound input and uncompressed payload sizes.
    """
    source = Path(path)
    return _check(_snapshot(source, max_part_bytes, max_total_bytes), source,
                  max_part_bytes, max_total_bytes)


def repair_file(path: str | Path, *, output: str | Path,
                max_part_bytes: int = _DEFAULT_PART,
                max_total_bytes: int = _DEFAULT_TOTAL) -> FileRepairResult:
    """Repair stale ZIP local-header CRC/size fields into a new file.

    Payloads must pass central-directory CRC verification and all supported
    checks. No parts, VBA, signatures, or compressed bytes are changed. Other
    damage or incomplete checking raises FileRepairError with its report.
    The output must not exist (including the original file); it is written
    only after checking the candidate. A healthy file is copied unchanged.
    This does not rebuild databases or recover missing/corrupt content.
    """
    source, target = Path(path), Path(output)
    if source.resolve() == target.resolve():
        raise ValueError("Repair output must differ from the original file")
    if target.exists():
        raise FileExistsError(target)
    raw = _snapshot(source, max_part_bytes, max_total_bytes)
    before = _check(raw, source, max_part_bytes, max_total_bytes)
    if not before.complete or any(not issue.repairable for issue in before.errors):
        raise FileRepairError("File has damage or unsupported checks that cannot be safely repaired", before)
    changes = before.errors
    candidate = bytearray(raw)
    if changes:
        with zipfile.ZipFile(io.BytesIO(raw)) as archive:
            names = {issue.location for issue in changes}
            for info in archive.infolist():
                if info.filename in names:
                    struct.pack_into("<III", candidate, info.header_offset + 14,
                                     info.CRC, info.compress_size, info.file_size)
    after = _check(bytes(candidate), source, max_part_bytes, max_total_bytes)
    if not after.ok:
        raise FileRepairError("Candidate failed verification; no output written", after)
    stream = target.open("xb")
    try:
        with stream:
            stream.write(candidate)
    except BaseException:
        target.unlink(missing_ok=True)
        raise
    after = FileCheckReport(target, after.issues, after.checked_parts, after.complete)
    return FileRepairResult(target, before, after, changes)
