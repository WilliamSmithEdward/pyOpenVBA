"""File checking catches damage fast saves preserve; repair never guesses."""

from __future__ import annotations

import io
import struct
import zipfile
from pathlib import Path
from typing import IO, Any

import pytest

from pyopenvba import ExcelFile, FileRepairError, check_file, repair_file

CT = b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="bin" ContentType="application/octet-stream"/></Types>'
REL = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{}</Relationships>'
MAIN_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
ROOT_REL = REL.format(f'<Relationship Id="main" Type="{MAIN_TYPE}" Target="doc.xml"/>').encode()


def package(path: Path, **parts: bytes) -> bytes:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_STORED) as archive:
        archive.writestr("[Content_Types].xml", CT)
        archive.writestr("doc.xml", b"<document/>")
        if "_rels/.rels" not in parts:
            archive.writestr("_rels/.rels", ROOT_REL)
        for name, data in parts.items():
            archive.writestr(name, data)
    return path.read_bytes()


def damage_header(path: Path, offset: int, value: int) -> bytes:
    raw = bytearray(path.read_bytes())
    with zipfile.ZipFile(io.BytesIO(raw)) as archive:
        header = archive.getinfo("doc.xml").header_offset
    struct.pack_into("<I", raw, header + offset, value)
    path.write_bytes(raw)
    return bytes(raw)


@pytest.mark.parametrize("offset", [14, 18, 22])
def test_repair_verified_local_metadata(tmp_path: Path, offset: int) -> None:
    path, output = tmp_path / "book.xlsm", tmp_path / "fixed.xlsm"
    original = package(path)
    damaged = damage_header(path, offset, 123)
    report = check_file(path)
    assert not report.ok and report.complete
    assert len(report.errors) == 1
    assert report.errors[0].code == "zip.local_metadata"
    assert report.errors[0].repairable
    result = repair_file(path, output=output)
    assert result.after.ok and result.after.path == output
    assert result.changes == report.errors
    assert output.read_bytes() == original
    assert path.read_bytes() == damaged


def test_corrupt_payload_refused_without_output(tmp_path: Path) -> None:
    path, output = tmp_path / "book.xlsm", tmp_path / "fixed.xlsm"
    raw = bytearray(package(path))
    position = raw.index(b"<document/>")
    raw[position] ^= 1
    path.write_bytes(raw)
    assert any(issue.code == "zip.part" for issue in check_file(path).errors)
    with pytest.raises(FileRepairError) as failure:
        repair_file(path, output=output)
    assert failure.value.report.errors
    assert not output.exists()
    assert path.read_bytes() == raw


def test_healthy_copy_and_no_overwrite(tmp_path: Path) -> None:
    path, output = tmp_path / "book.xlsm", tmp_path / "copy.xlsm"
    raw = package(path)
    assert check_file(path).ok
    assert repair_file(path, output=output).changes == ()
    assert output.read_bytes() == raw
    with pytest.raises(FileExistsError):
        repair_file(path, output=output)
    with pytest.raises(ValueError):
        repair_file(path, output=path)


@pytest.mark.parametrize("target", ["missing.xml", "../../doc.xml", "https://example.com/doc.xml"])
def test_dangling_relationship_refused(tmp_path: Path, target: str) -> None:
    path = tmp_path / "book.xlsm"
    relationship = f'<Relationship Id="rId1" Type="type" Target="{target}"/>'
    package(path, **{"_rels/.rels": REL.format(relationship).encode()})
    assert any(issue.code == "opc.relationship_target" for issue in check_file(path).errors)
    with pytest.raises(FileRepairError):
        repair_file(path, output=tmp_path / "out.xlsm")


def test_external_and_percent_encoded_relationships(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsm"
    relationships = f'<Relationship Id="rId1" Type="type" Target="https://example.com" TargetMode="External"/><Relationship Id="rId2" Type="{MAIN_TYPE}" Target="/doc%2Exml"/>'
    package(path, **{"_rels/.rels": REL.format(relationships).encode()})
    assert check_file(path).ok


@pytest.mark.parametrize("xml", [b"<broken>", b'<!DOCTYPE a [<!ENTITY e "hello">]><a>&e;</a>', '<!DOCTYPE a><a/>'.encode("utf-16")])
def test_xml_damage_and_dtd_rejected(tmp_path: Path, xml: bytes) -> None:
    path = tmp_path / "book.xlsm"
    package(path, **{"bad.xml": xml})
    assert any(issue.code == "opc.xml" for issue in check_file(path).errors)


def test_missing_content_type_and_duplicate_part(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsm"
    package(path, **{"unknown.xyz": b"payload"})
    assert any(issue.code == "opc.content_type_missing" for issue in check_file(path).errors)
    with pytest.warns(UserWarning):
        with zipfile.ZipFile(path, "a") as archive:
            archive.writestr("doc.xml", b"<document/>")
    assert any(issue.code == "zip.duplicate" for issue in check_file(path).errors)
    with pytest.raises(FileRepairError):
        repair_file(path, output=tmp_path / "out.xlsm")


def test_limit_reports_incomplete_and_refuses_repair(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsm"
    package(path)
    report = check_file(path, max_part_bytes=10)
    assert not report.complete and not report.ok
    assert any(issue.code == "check.limit" for issue in report.warnings)
    with pytest.raises(FileRepairError):
        repair_file(path, output=tmp_path / "out.xlsm", max_part_bytes=10)
    with pytest.raises(ValueError):
        check_file(path, max_total_bytes=1)
    with pytest.raises(ValueError):
        check_file(path, max_part_bytes=0)


def test_malformed_file_report_and_missing_path_exception(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsm"
    path.write_bytes(b"PK\x03\x04broken")
    assert not check_file(path).ok
    with pytest.raises(FileNotFoundError):
        check_file(tmp_path / "absent.xlsm")


@pytest.mark.parametrize("suffix", [".xlsm", ".docm", ".pptm"])
def test_real_office_templates(tmp_path: Path, suffix: str) -> None:
    from pyopenvba import PowerPointFile, WordFile
    host = {".xlsm": ExcelFile, ".docm": WordFile, ".pptm": PowerPointFile}[suffix]
    path = tmp_path / ("new" + suffix)
    with host.create_new(path):
        pass
    report = check_file(path)
    assert report.ok, report.issues


def test_access_coverage_is_explicit(tmp_path: Path) -> None:
    from pyopenvba import AccessDatabase
    path = tmp_path / "new.accdb"
    with AccessDatabase.create_new(path):
        pass
    report = check_file(path)
    assert not report.complete and not report.ok
    assert not report.errors
    assert report.checked_parts == 1
    assert report.warnings[0].code == "check.access_scope"
    with pytest.raises(FileRepairError):
        repair_file(path, output=tmp_path / "out.accdb")


def test_mixed_damage_cannot_be_partially_repaired(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsm"
    package(path, **{"bad.xml": b"<broken>"})
    damaged = damage_header(path, 14, 123)
    with pytest.raises(FileRepairError):
        repair_file(path, output=tmp_path / "out.xlsm")
    assert path.read_bytes() == damaged
    assert not (tmp_path / "out.xlsm").exists()


def test_compressed_corruption_and_local_name_mismatch(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsm"
    package(path)
    raw = bytearray(path.read_bytes())
    with zipfile.ZipFile(io.BytesIO(raw)) as archive:
        info = archive.getinfo("doc.xml")
    raw[info.header_offset + 30] ^= 1
    path.write_bytes(raw)
    assert any(issue.code == "zip.part" for issue in check_file(path).errors)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CT)
        archive.writestr("doc.xml", b"<document>" + b"a" * 10000 + b"</document>")
        archive.writestr("_rels/.rels", ROOT_REL)
    raw = bytearray(path.read_bytes())
    with zipfile.ZipFile(io.BytesIO(raw)) as archive:
        info = archive.getinfo("doc.xml")
    raw[info.header_offset + 30 + len(info.filename) + 3] ^= 0x80
    path.write_bytes(raw)
    assert any(issue.code == "zip.part" for issue in check_file(path).errors)


class Unseekable(io.BytesIO):
    def seekable(self) -> bool:
        return False

    def seek(self, offset: int, whence: int = 0) -> int:
        raise io.UnsupportedOperation("unseekable")


def test_data_descriptor_checked(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsm"
    stream = Unseekable()
    with zipfile.ZipFile(stream, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CT)
        archive.writestr("doc.xml", b"<document/>")
        archive.writestr("_rels/.rels", ROOT_REL)
    raw = bytearray(stream.getvalue())
    path.write_bytes(raw)
    assert check_file(path).ok
    offset = raw.index(b"PK\x07\x08") + 4
    raw[offset] ^= 1
    path.write_bytes(raw)
    assert any(issue.code == "zip.descriptor" for issue in check_file(path).errors)
    with pytest.raises(FileRepairError):
        repair_file(path, output=tmp_path / "out.xlsm")


def test_zip64_partial_check_refuses_repair(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsm"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", CT)
        with archive.open("doc.xml", "w", force_zip64=True) as stream:
            stream.write(b"<document/>")
        archive.writestr("_rels/.rels", ROOT_REL)
    report = check_file(path)
    assert not report.complete and not report.errors
    assert any(issue.code == "check.zip64" for issue in report.warnings)
    with pytest.raises(FileRepairError):
        repair_file(path, output=tmp_path / "out.xlsm")


def test_cfb_directory_cycle_and_truncated_stream(tmp_path: Path) -> None:
    from pyopenvba.cfb import CFB
    from pyopenvba._templates import EMPTY_XLSM_BYTES
    with zipfile.ZipFile(io.BytesIO(EMPTY_XLSM_BYTES)) as archive:
        original = archive.read("xl/vbaProject.bin")
    cfb = CFB.from_bytes(original)
    # Change a directory entry's declared size; its chain cannot supply it.
    cfb.write_stream("PROJECT", b"x" * 5000)
    raw = bytearray(cfb.to_bytes())
    directory_sector = struct.unpack_from("<I", raw, 48)[0]
    entry_start = (directory_sector + 1) * 512
    path = tmp_path / "project.bin"
    # Root child points back to root: reader traversal tolerates this.
    struct.pack_into("<I", raw, entry_start + 76, 0)
    path.write_bytes(raw)
    assert any("Cycle" in issue.message for issue in check_file(path).errors)
    raw = bytearray(original)
    directory_sector = struct.unpack_from("<I", raw, 48)[0]
    entry_start = (directory_sector + 1) * 512
    for index in range(4):
        if raw[entry_start + index * 128 + 66] == 2:
            struct.pack_into("<I", raw, entry_start + index * 128 + 120, 100000)
            break
    else:
        pytest.fail("Template has no stream in its first directory sector")
    path.write_bytes(raw)
    assert not check_file(path).ok


def test_standard_extension_and_unknown_extension(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    package(path)
    assert check_file(path).ok
    other = tmp_path / "other.zip"
    other.write_bytes(path.read_bytes())
    assert not check_file(other).complete


def test_repaired_office_templates_preserve_all_other_bytes(tmp_path: Path) -> None:
    from pyopenvba import PowerPointFile, WordFile
    for suffix, host in ((".xlsm", ExcelFile), (".docm", WordFile), (".pptm", PowerPointFile)):
        path = tmp_path / ("book" + suffix)
        output = tmp_path / ("repaired" + suffix)
        with host.create_new(path):
            pass
        original = path.read_bytes()
        raw = bytearray(original)
        with zipfile.ZipFile(io.BytesIO(raw)) as archive:
            info = archive.infolist()[0]
        struct.pack_into("<I", raw, info.header_offset + 14, info.CRC ^ 1)
        path.write_bytes(raw)
        result = repair_file(path, output=output)
        assert len(result.changes) == 1
        assert result.after.ok
        assert output.read_bytes() == original
        assert path.read_bytes() == raw
        with host(output) as document:
            assert document.has_vba_project()


def test_output_creation_race_does_not_overwrite(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    path, output = tmp_path / "book.xlsm", tmp_path / "out.xlsm"
    package(path)
    original_open = Path.open

    def racing_open(self: Path, mode: str = "r", *args: Any, **kwargs: Any) -> IO[Any]:
        if self == output and mode == "xb":
            output.write_bytes(b"another writer")
        return original_open(self, mode, *args, **kwargs)

    monkeypatch.setattr(Path, "open", racing_open)
    with pytest.raises(FileExistsError):
        repair_file(path, output=output)
    assert output.read_bytes() == b"another writer"


def test_lazy_vba_source_damage_is_not_missed(tmp_path: Path) -> None:
    from pyopenvba.cfb import CFB
    from pyopenvba._templates import EMPTY_XLSM_BYTES
    from pyopenvba.vba import parse_vba_project
    path = tmp_path / "damaged.xlsm"
    with zipfile.ZipFile(io.BytesIO(EMPTY_XLSM_BYTES)) as archive:
        cfb = CFB.from_bytes(archive.read("xl/vbaProject.bin"))
        module = parse_vba_project(cfb).modules[0]
        original = cfb.get_stream_in_storage("VBA", module.stream_name)
        # First source chunk remains readable, but a later chunk is invalid.
        cfb.write_stream(module.stream_name, original + b"\x00\x00")
        with zipfile.ZipFile(path, "w") as output:
            for info in archive.infolist():
                output.writestr(info, cfb.to_bytes() if info.filename == "xl/vbaProject.bin" else archive.read(info))
    report = check_file(path)
    assert not report.ok and report.errors
    with pytest.raises(FileRepairError):
        repair_file(path, output=tmp_path / "out.xlsm")


def test_false_central_size_does_not_hide_payload(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsm"
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CT)
        archive.writestr("payload.bin", b"content that must not disappear")
        archive.writestr("doc.xml", b"<document/>")
        archive.writestr("_rels/.rels", ROOT_REL)
    raw = bytearray(path.read_bytes())
    with zipfile.ZipFile(io.BytesIO(raw)) as archive:
        info = archive.getinfo("payload.bin")
    offset = raw.index(b"PK\x01\x02", raw.index(b"PK\x01\x02") + 4)
    struct.pack_into("<I", raw, offset + 16, 0)  # central CRC
    struct.pack_into("<I", raw, offset + 24, 0)  # central uncompressed size
    struct.pack_into("<I", raw, info.header_offset + 14, 0)
    struct.pack_into("<I", raw, info.header_offset + 22, 0)
    path.write_bytes(raw)
    # zipfile alone accepts the zero-sized prefix and hides real payload data.
    with zipfile.ZipFile(path) as archive:
        assert archive.read("payload.bin") == b""
    assert any(issue.code == "zip.part" for issue in check_file(path).errors)
    with pytest.raises(FileRepairError):
        repair_file(path, output=tmp_path / "out.xlsm")


def test_unsupported_compression_is_incomplete(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsm"
    package(path)
    with zipfile.ZipFile(path, "a", compression=zipfile.ZIP_BZIP2) as archive:
        archive.writestr("payload.bin", b"data")
    report = check_file(path)
    assert not report.complete and not report.ok
    assert any(issue.code == "check.compression" for issue in report.warnings)
    with pytest.raises(FileRepairError):
        repair_file(path, output=tmp_path / "out.xlsm")


def test_missing_package_relationships(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsm"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", CT)
        archive.writestr("doc.xml", b"<document/>")
    assert any(issue.code == "opc.relationships" for issue in check_file(path).errors)
