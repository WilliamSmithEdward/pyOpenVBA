"""A save carries the parts it does not change over as they were stored.

``save()`` used to read every part out of the package and deflate it
again.  It now copies an unchanged part's stored bytes
(:mod:`pyopenvba._package_copy`), and still writes with ``zipfile`` a
package that cannot be copied that way.  These tests hold the two
writers to the same result, over every zip-based file with a project in
the repository, and pin which packages each one takes.
"""

from __future__ import annotations

import io
import struct
import warnings
import zipfile
from pathlib import Path

import pytest

import pyopenvba._host as host_module
from pyopenvba import ExcelFile, PowerPointFile, WordFile
from pyopenvba._host import VBAHostFile
from pyopenvba._package_copy import copy_package
from pyopenvba.powerquery._opc import OpcFile

ROOT = Path(__file__).parent.parent
HOSTS: dict[str, type[VBAHostFile]] = {
    ".xlsm": ExcelFile, ".xlsb": ExcelFile, ".xlam": ExcelFile,
    ".docm": WordFile, ".pptm": PowerPointFile,
}


def _packages() -> list[Path]:
    """Every zip-based Office file in the repository that these classes save."""
    found: list[Path] = []
    for folder in ("tests", "demo", "examples"):
        for path in sorted((ROOT / folder).rglob("*")):
            if path.suffix.lower() in HOSTS and zipfile.is_zipfile(path):
                found.append(path)
    return found


PACKAGES = _packages()
MARK = "\r\n' edited by test_package_copy\r\n"


def _edit(book: VBAHostFile) -> None:
    """Change the project: its first module, or a first module in a project made for a file without one."""
    if not book.has_vba_project():
        book.add_vba_project().add_module("Added", "Public Sub Run()\r\nEnd Sub\r\n")
        return
    project = book.vba_project()
    if project.modules:
        name = project.modules[0].name
        book.set_module(name, book.get_module(name) + MARK)
    else:
        project.add_module("Added", "Public Sub Run()\r\nEnd Sub\r\n")


def _save(path: Path, out: Path, *, copying: bool, monkeypatch: pytest.MonkeyPatch,
          edit: bool = True) -> bytes | type[BaseException]:
    """The file ``save()`` writes for ``path``, by the copying writer or by ``zipfile``; or the error it raised."""
    taken: list[bool] = []

    def writer(raw: bytes, infos: list[zipfile.ZipInfo], edits: dict[str, bytes | None],
               added: dict[str, bytes]) -> bytes | None:
        packed = copy_package(raw, infos, edits, added) if copying else None
        taken.append(packed is not None)
        return packed

    monkeypatch.setattr(host_module, "copy_package", writer)
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            with HOSTS[path.suffix.lower()](path) as book:
                if edit:
                    _edit(book)
                book.save(out, allow_protected=True, allow_invalidate_signature=True)
    except Exception as exc:
        return type(exc)
    if copying and taken:
        assert taken == [True], f"{path.name} was not written by the copying writer"
    return out.read_bytes()


def _stored(raw: bytes) -> dict[str, tuple[bytes, bytes]]:
    """Each part's bytes as the package stores them, with the extra field of its header."""
    return {entry.name: (entry.body, entry.local_extra) for entry in OpcFile.parse(raw).entries}


def test_there_are_packages_to_test() -> None:
    assert len(PACKAGES) >= 40
    assert {path.suffix.lower() for path in PACKAGES} == set(HOSTS) - {".xlam"}


@pytest.mark.parametrize("path", PACKAGES, ids=lambda path: path.name)
def test_both_writers_write_the_same_package(path: Path, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """Part for part: the same names in the same order, the same content, compression method and date."""
    copied = _save(path, tmp_path / ("copied" + path.suffix), copying=True, monkeypatch=monkeypatch)
    rewritten = _save(path, tmp_path / ("rewritten" + path.suffix), copying=False, monkeypatch=monkeypatch)
    if not isinstance(copied, bytes) or not isinstance(rewritten, bytes):
        assert copied is rewritten   # a file that refuses the save refuses it either way
        return
    with zipfile.ZipFile(io.BytesIO(copied)) as fast, zipfile.ZipFile(io.BytesIO(rewritten)) as slow:
        assert fast.testzip() is None
        assert fast.namelist() == slow.namelist()
        for ours, theirs in zip(fast.infolist(), slow.infolist(), strict=True):
            assert fast.read(ours) == slow.read(theirs), ours.filename
            assert (ours.compress_type, ours.date_time) == (theirs.compress_type, theirs.date_time), ours.filename


@pytest.mark.parametrize("path", PACKAGES, ids=lambda path: path.name)
def test_a_part_keeps_the_header_the_file_gave_it(path: Path, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """Every part the file had, changed or not, keeps its attributes, flags, versions, extra field and comment.

    ``zipfile`` writes a header of its own for each part: it gives a part
    whose attributes are zero, as Office writes them, a file mode, and
    drops the flags, the extra field and the comment.
    """
    copied = _save(path, tmp_path / ("copied" + path.suffix), copying=True, monkeypatch=monkeypatch)
    if not isinstance(copied, bytes):
        return
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(io.BytesIO(copied)) as saved:
        before = {info.filename: info for info in source.infolist()}
        kept = [info for info in saved.infolist() if info.filename in before]
        assert kept
        for info in kept:
            was = before[info.filename]
            assert (
                info.compress_type, info.date_time, info.external_attr, info.internal_attr, info.create_system,
                info.create_version, info.extract_version, info.flag_bits, info.extra, info.comment,
            ) == (
                was.compress_type, was.date_time, was.external_attr, was.internal_attr, was.create_system,
                was.create_version, was.extract_version, was.flag_bits & ~0x08, was.extra, was.comment,
            ), info.filename


@pytest.mark.parametrize("path", PACKAGES, ids=lambda path: path.name)
def test_an_unchanged_part_keeps_its_stored_bytes(path: Path, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """What the edit does not reach is byte for byte what the file held: nothing was deflated again."""
    copied = _save(path, tmp_path / ("copied" + path.suffix), copying=True, monkeypatch=monkeypatch)
    rewritten = _save(path, tmp_path / ("rewritten" + path.suffix), copying=False, monkeypatch=monkeypatch)
    if not isinstance(copied, bytes) or not isinstance(rewritten, bytes):
        return
    before, after, slow = _stored(path.read_bytes()), _stored(copied), _stored(rewritten)
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(io.BytesIO(copied)) as saved:
        unchanged = [name for name in saved.namelist()
                     if name in before and source.read(name) == saved.read(name)]
        changed = [name for name in saved.namelist() if name not in unchanged]
    assert unchanged, "the edit changed every part"
    for name in unchanged:
        assert after[name] == before[name], name
    # A part that did change is deflated by zlib as zipfile deflates it: the same stored bytes either way.
    for name in changed:
        assert after[name][0] == slow[name][0], name


@pytest.mark.parametrize("path", [p for p in PACKAGES if p.name == "test_macro_workbook.xlsm"], ids=lambda p: p.name)
def test_a_save_with_no_edit_keeps_every_other_part(path: Path, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    saved = _save(path, tmp_path / "noop.xlsm", copying=True, monkeypatch=monkeypatch, edit=False)
    assert isinstance(saved, bytes)
    before, after = _stored(path.read_bytes()), _stored(saved)
    assert list(after) == list(before)
    assert {name for name in before if after[name] != before[name]} <= {"xl/vbaProject.bin"}


# ---------------------------------------------------------------------------
# copy_package itself
# ---------------------------------------------------------------------------

def _zip(parts: dict[str, bytes], *, method: int = zipfile.ZIP_DEFLATED) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", method) as package:
        for name, data in parts.items():
            package.writestr(name, data)
    return buffer.getvalue()


def _infos(raw: bytes) -> list[zipfile.ZipInfo]:
    with zipfile.ZipFile(io.BytesIO(raw)) as package:
        return package.infolist()


PARTS = {"a.xml": b"<a/>" * 200, "b.bin": bytes(range(256)) * 8, "c.xml": b"<c/>"}


def test_a_part_is_replaced_dropped_and_added() -> None:
    raw = _zip(PARTS)
    packed = copy_package(raw, _infos(raw), {"a.xml": b"new", "c.xml": None}, {"d.xml": b"<d/>"})
    assert packed is not None
    with zipfile.ZipFile(io.BytesIO(packed)) as package:
        assert package.testzip() is None
        assert package.namelist() == ["a.xml", "b.bin", "d.xml"]
        assert package.read("a.xml") == b"new"
        assert package.read("b.bin") == PARTS["b.bin"]
        assert package.read("d.xml") == b"<d/>"
        assert package.getinfo("d.xml").date_time == (1980, 1, 1, 0, 0, 0)
    assert _stored(packed)["b.bin"] == _stored(raw)["b.bin"]


def test_an_edit_naming_no_part_changes_nothing() -> None:
    raw = _zip(PARTS)
    packed = copy_package(raw, _infos(raw), {"absent.bin": b"x"}, {})
    assert packed is not None
    assert _stored(packed) == _stored(raw)


def test_a_stored_part_stays_stored() -> None:
    raw = _zip(PARTS, method=zipfile.ZIP_STORED)
    packed = copy_package(raw, _infos(raw), {"a.xml": b"new"}, {})
    assert packed is not None
    with zipfile.ZipFile(io.BytesIO(packed)) as package:
        assert package.getinfo("a.xml").compress_type == zipfile.ZIP_STORED
        assert package.read("a.xml") == b"new"


class _Unseekable(io.RawIOBase):
    """A stream that cannot seek, which makes ``zipfile`` write each part's sizes after its bytes."""

    def __init__(self) -> None:
        self.data = bytearray()

    def writable(self) -> bool:
        return True

    def write(self, b: object) -> int:
        self.data += bytes(b)  # type: ignore[arg-type]
        return len(bytes(b))  # type: ignore[arg-type]


def test_sizes_written_after_a_part_move_into_its_header() -> None:
    """Such a part's header holds no sizes.  Carried over, it holds them, and says so."""
    stream = _Unseekable()
    with zipfile.ZipFile(stream, "w", zipfile.ZIP_DEFLATED) as package:  # type: ignore[arg-type]
        for name, data in PARTS.items():
            package.writestr(name, data)
    raw = bytes(stream.data)
    assert all(info.flag_bits & 0x08 for info in _infos(raw))

    packed = copy_package(raw, _infos(raw), {"a.xml": b"new"}, {})
    assert packed is not None
    with zipfile.ZipFile(io.BytesIO(packed)) as package:
        assert package.testzip() is None
        assert not any(info.flag_bits & 0x08 for info in package.infolist())
        assert package.read("b.bin") == PARTS["b.bin"]
        for info in package.infolist():
            # The local header, which is where a reader that streams the file looks.
            crc, packed_size, size = struct.unpack_from("<III", packed, info.header_offset + 14)
            assert (crc, packed_size, size) == (info.CRC, info.compress_size, info.file_size)


def test_a_part_compressed_another_way_is_copied_but_not_written() -> None:
    pytest.importorskip("bz2")
    raw = _zip(PARTS, method=zipfile.ZIP_BZIP2)
    assert copy_package(raw, _infos(raw), {"a.xml": b"new"}, {}) is None
    packed = copy_package(raw, _infos(raw), {}, {})
    assert packed is not None
    with zipfile.ZipFile(io.BytesIO(packed)) as package:
        assert package.read("a.xml") == PARTS["a.xml"]


def test_packages_left_to_zipfile() -> None:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")   # zipfile warns of the duplicate name
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as package:
            package.writestr("a.xml", b"one")
            package.writestr("a.xml", b"two")
    twice = buffer.getvalue()
    assert copy_package(twice, _infos(twice), {}, {}) is None

    # The bytes that open a ZIP64 locator, anywhere in the file, are taken for one.
    zip64 = _zip({"a.bin": b"PK\x06\x07" + bytes(16)}, method=zipfile.ZIP_STORED)
    assert copy_package(zip64, _infos(zip64), {}, {}) is None

    raw = _zip(PARTS)
    assert copy_package(raw, _infos(raw)[:-1], {}, {}) is None   # the readers disagree on the parts
    assert copy_package(b"not a zip", [], {}, {}) is None
    assert copy_package(raw[:-30], _infos(raw), {}, {}) is None  # no end record

    encrypted = bytearray(raw)
    for info in _infos(raw):
        encrypted[info.header_offset + 6] |= 0x01
    central = raw.rfind(b"PK\x01\x02")
    encrypted[central + 8] |= 0x01
    assert copy_package(bytes(encrypted), _infos(raw), {}, {}) is None


def _central(raw: bytes, name: str) -> int:
    """Where the central directory's record for ``name`` starts."""
    at = raw.index(b"PK")
    while raw[at + 46 : at + 46 + len(name)] != name.encode():
        at = raw.index(b"PK", at + 4)
    return at


def test_directories_that_disagree_are_left_to_zipfile() -> None:
    """The part's own header and the central directory each describe it, and a copy needs both to be right."""
    raw = _zip(PARTS)
    infos = _infos(raw)
    assert copy_package(raw, infos, {}, {}) is not None

    # The part's own header names it differently.
    renamed = bytearray(raw)
    renamed[infos[1].header_offset + 30] ^= 0x01
    assert copy_package(bytes(renamed), infos, {}, {}) is None
    with zipfile.ZipFile(io.BytesIO(bytes(renamed))) as package, pytest.raises(zipfile.BadZipFile):
        package.read("b.bin")

    # Two entries of the directory point at the same bytes.
    overlapping = bytearray(raw)
    struct.pack_into("<I", overlapping, _central(raw, "c.xml") + 42, infos[0].header_offset)
    struct.pack_into("<III", overlapping, _central(raw, "c.xml") + 16,
                     infos[0].CRC, infos[0].compress_size, infos[0].file_size)
    assert copy_package(bytes(overlapping), _infos(bytes(overlapping)), {}, {}) is None


def test_a_size_held_in_a_zip64_record_is_left_to_zipfile() -> None:
    """A part may keep its size in a ZIP64 record in a package with no ZIP64 end records to give it away."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as package:
        with package.open("a.xml", "w", force_zip64=True) as part:
            part.write(PARTS["a.xml"])
        package.writestr("b.bin", PARTS["b.bin"])
    raw = buffer.getvalue()
    assert b"PK" not in raw
    assert copy_package(raw, _infos(raw), {}, {}) is None
    assert copy_package(raw, _infos(raw), {"a.xml": b"new"}, {}) is None


def test_a_package_left_to_zipfile_still_saves(tmp_path: Path) -> None:
    """A workbook holding the bytes of a ZIP64 locator goes the old way, and saves as it did."""
    source = ROOT / "demo" / "test_macro_workbook.xlsm"
    odd, out = tmp_path / "odd.xlsm", tmp_path / "out.xlsm"
    with zipfile.ZipFile(source) as original, zipfile.ZipFile(odd, "w", zipfile.ZIP_DEFLATED) as copy:
        for info in original.infolist():
            copy.writestr(info, original.read(info.filename))
        copy.writestr(zipfile.ZipInfo("docProps/odd.bin"), b"PK\x06\x07" + bytes(16))
    raw = odd.read_bytes()
    assert copy_package(raw, _infos(raw), {}, {}) is None
    with ExcelFile(odd) as book:
        name = book.module_names()[0]
        book.set_module(name, book.get_module(name) + MARK)
        book.save(out)
    with ExcelFile(out) as book, zipfile.ZipFile(out) as package:
        assert book.get_module(name).endswith(MARK)
        assert package.read("docProps/odd.bin") == b"PK\x06\x07" + bytes(16)


def test_a_damaged_part_the_save_does_not_change_is_carried_over(tmp_path: Path) -> None:
    """An unchanged part is not read, so it is not checked either: it leaves as it came.

    ``zipfile`` read each part to deflate it again, which refused the
    whole save with ``BadZipFile`` over a part the edit had nothing to do
    with.  The part is as damaged after the save as before it, and the
    project beside it is saved.
    """
    source = ROOT / "demo" / "test_macro_workbook.xlsm"
    raw = bytearray(source.read_bytes())
    victim = max((info for info in _infos(bytes(raw)) if info.filename != "xl/vbaProject.bin"),
                 key=lambda info: info.compress_size)
    name_length, extra_length = struct.unpack_from("<HH", raw, victim.header_offset + 26)
    middle = victim.header_offset + 30 + name_length + extra_length + victim.compress_size // 2
    raw[middle] ^= 0xFF
    raw[middle + 1] ^= 0xFF
    damaged, out = tmp_path / "damaged.xlsm", tmp_path / "out.xlsm"
    damaged.write_bytes(bytes(raw))

    with ExcelFile(damaged) as book:
        name = book.module_names()[0]
        book.set_module(name, book.get_module(name) + MARK)
        book.save(out)
    assert _stored(out.read_bytes())[victim.filename] == _stored(bytes(raw))[victim.filename]
    with ExcelFile(out) as book:
        assert book.get_module(name).endswith(MARK)
    with zipfile.ZipFile(out) as package, pytest.raises(zipfile.BadZipFile):
        package.read(victim.filename)
