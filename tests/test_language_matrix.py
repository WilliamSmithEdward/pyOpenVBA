"""Per-code-page language matrix: non-ASCII text through the real engine.

Ported from xlide_vscode's ``tests/vbaLanguageMatrix.test.ts`` (see
GitHub issue #13).  One native-language sample per supported code page,
each run through a full write -> read -> list -> validate cycle on a
workbook whose PROJECTCODEPAGE is that page.

Why a per-page sweep rather than spot checks: every bug in this class --
here and in the TypeScript port -- survived because the test corpus was
ASCII/cp1252-authored.  Writing this matrix immediately found cp1258
Vietnamese being destroyed on encode by a codec layer that already had
passing code-page tests.

The first assertion is the load-bearing one.  With ``errors="replace"``
a wall of ``?`` round-trips happily, so "decode(encode(x)) == x" alone
passes while the text is being destroyed; assert zero substitution bytes.

NFC normalization matters only for cp1258: that page stores stacked
diacritics as a precomposed base plus a combining mark, so the decoded
text is canonically equivalent but not identical.  See
:func:`pyopenvba.vba.encode_mbcs`.
"""

from __future__ import annotations

import io
import shutil
import struct
import unicodedata
import zipfile
from pathlib import Path

import pytest

from pyopenvba import ExcelFile
from pyopenvba.access import AccessDatabase
from pyopenvba.access._vba import STORAGE_TABLE, encoding_of
from pyopenvba.access_read import AccessReader
from pyopenvba.cfb import CFB
from pyopenvba.vba import (
    VBAModuleKind,
    _encoding_for_codepage,  # pyright: ignore[reportPrivateUsage]
    compress,
    decompress,
    encode_mbcs,
)

_VBA_ENTRY = "xl/vbaProject.bin"
_ACCESS_TEMPLATE = (
    Path(__file__).parents[1]
    / "src"
    / "pyopenvba"
    / "_templates"
    / "blank_files"
    / "blank_database.accdb"
)

# (code page, language label, native sample text)
LANGUAGE_MATRIX: list[tuple[int, str, str]] = [
    (874, "Thai", "ทดสอบภาษาไทย"),
    (932, "Japanese (Shift-JIS)", "テスト用モジュール"),
    (936, "Chinese Simplified (GBK)", "中文测试模块"),
    (949, "Korean (EUC-KR)", "한국어 테스트"),
    (950, "Chinese Traditional (Big5)", "繁體中文測試"),
    (1250, "Central European", "Příliš žluťoučký kůň Zażółć gęślą"),
    (1251, "Cyrillic", "Проверка русского текста"),
    (1252, "Western European", "déjà vu € œuvre Straße"),
    (1253, "Greek", "Δοκιμή ελληνικού κειμένου"),
    (1254, "Turkish", "Türkçe deneme ğüşiöç İı"),
    (1255, "Hebrew", "בדיקת עברית"),
    (1256, "Arabic", "اختبار العربية"),
    (1257, "Baltic", "Lietuviškas tekstas ąčęėįšųū"),
    (1258, "Vietnamese", "Tiếng Việt thử nghiệm"),
    (10000, "Mac Roman", "déjà vu café œuvre"),
    (20866, "Russian (KOI8-R)", "Тест КОИ-8"),
    (21866, "Ukrainian (KOI8-U)", "Тест української ґї"),
    (28592, "ISO-8859-2", "Zažil žluťoučký Zażółć"),
    (28595, "ISO-8859-5", "Проверка ИСО"),
    (54936, "GB18030", "中文 GB18030 测试"),
    (65001, "UTF-8", "любой текст 中文 déjà ทดสอบ"),
]

# Native module names, for the pages where VBA hosts routinely see them.
NAME_MATRIX: list[tuple[int, str, str]] = [
    (1251, "Cyrillic", "МодульТест"),
    (932, "Japanese", "モジュール"),
    (936, "Chinese Simplified", "测试模块"),
]

_IDS = [f"cp{cp}-{label}" for cp, label, _ in LANGUAGE_MATRIX]
_NAME_IDS = [f"cp{cp}-{label}" for cp, label, _ in NAME_MATRIX]


def _patch_dir_code_page(dir_raw: bytes, code_page: int) -> bytes:
    """Rewrite the PROJECTCODEPAGE (0x0003) record of a decompressed dir stream.

    Walks the ``[id u16][size u32][data]`` record sequence.  PROJECTVERSION
    (0x0009) is special-cased: its size slot is a reserved marker and its
    payload is a fixed 10 bytes ([MS-OVBA] 2.3.4.2.1.11).
    """
    buf = bytearray(dir_raw)
    pos = 0
    while pos + 6 <= len(buf):
        record_id = struct.unpack_from("<H", buf, pos)[0]
        if record_id == 0x0009:
            pos += 12
            continue
        size = struct.unpack_from("<I", buf, pos + 2)[0]
        if record_id == 0x0003:
            struct.pack_into("<H", buf, pos + 6, code_page)
            return bytes(buf)
        pos += 6 + size
    raise AssertionError("PROJECTCODEPAGE record not found in dir stream")


def _workbook_with_code_page(tmp_path: Path, code_page: int) -> Path:
    """Build a blank .xlsm whose project declares ``code_page``.

    Cheap by design: one template copy, one dir-stream record patched --
    no per-language binary fixtures to maintain.
    """
    target = tmp_path / f"cp{code_page}.xlsm"
    with ExcelFile.create_new(target) as wb:
        wb.save()

    original = target.read_bytes()
    with zipfile.ZipFile(io.BytesIO(original)) as zin:
        cfb = CFB.from_bytes(zin.read(_VBA_ENTRY))
        dir_raw = decompress(cfb.get_stream_in_storage("VBA", "dir"))
        cfb.write_stream_in_storage(
            "VBA", "dir", compress(_patch_dir_code_page(dir_raw, code_page))
        )
        patched_vba = cfb.to_bytes()
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as zout:
            for info in zin.infolist():
                data = (
                    patched_vba
                    if info.filename == _VBA_ENTRY
                    else zin.read(info.filename)
                )
                zout.writestr(info, data)
    target.write_bytes(buf.getvalue())

    with ExcelFile(target) as wb:
        assert wb.vba_project().code_page == code_page
    return target


def _nfc(text: str) -> str:
    return unicodedata.normalize("NFC", text)


@pytest.mark.parametrize(("code_page", "label", "sample"), LANGUAGE_MATRIX, ids=_IDS)
class TestCodecLayer:
    def test_encode_has_no_substitution_bytes(
        self, code_page: int, label: str, sample: str
    ) -> None:
        """No silent destruction: not one character becomes '?'.

        This is the assertion that catches the bug class -- a fully
        substituted string still round-trips through decode.
        """
        encoded = encode_mbcs(sample, _encoding_for_codepage(code_page))
        assert b"?" not in encoded, (
            f"{label} (cp{code_page}) lost characters on encode: {encoded!r}"
        )

    def test_round_trips_through_the_code_page(
        self, code_page: int, label: str, sample: str
    ) -> None:
        encoding = _encoding_for_codepage(code_page)
        decoded = encode_mbcs(sample, encoding).decode(encoding)
        assert _nfc(decoded) == _nfc(sample), f"{label} (cp{code_page}) round trip"


@pytest.mark.parametrize(("code_page", "label", "sample"), LANGUAGE_MATRIX, ids=_IDS)
def test_module_source_survives_full_workbook_cycle(
    tmp_path: Path, code_page: int, label: str, sample: str
) -> None:
    """End to end: push a module carrying the sample in a comment and a
    string literal, save, reopen, and read it back unchanged."""
    workbook = _workbook_with_code_page(tmp_path, code_page)
    source = (
        f"' {sample}\r\n"
        "Sub Probe()\r\n"
        f'    Dim s As String: s = "{sample}"\r\n'
        "End Sub\r\n"
    )
    with ExcelFile(workbook) as wb:
        wb.set_module("Module1", source)
        wb.save()

    with ExcelFile(workbook) as wb:
        round_tripped = wb.get_module("Module1")
        assert _nfc(sample) in _nfc(round_tripped), (
            f"{label} (cp{code_page}) source did not survive the round trip"
        )
        assert "?" not in round_tripped.replace("Probe()", ""), (
            f"{label} (cp{code_page}) source was substituted on write"
        )
        assert "Module1" in wb.module_names()
        assert wb.validate() == []


@pytest.mark.parametrize(("code_page", "label", "name"), NAME_MATRIX, ids=_NAME_IDS)
def test_native_module_names_survive_add_and_rename(
    tmp_path: Path, code_page: int, label: str, name: str
) -> None:
    """Native-language module NAMES: add, rename, and confirm the ANSI
    PROJECT declaration and the dir records all carry real bytes.

    The PROJECT stream is the one that regressed (issue #11): it is
    code-page ANSI, so a cp1252-hardcoded writer turned every non-Latin
    name into '?' while the dir stream stayed correct.
    """
    encoding = _encoding_for_codepage(code_page)
    workbook = _workbook_with_code_page(tmp_path, code_page)

    with ExcelFile(workbook) as wb:
        wb.vba_project().add_module(
            name, "Sub Probe()\r\nEnd Sub\r\n", kind=VBAModuleKind.standard
        )
        wb.save()

    with ExcelFile(workbook) as wb:
        assert name in wb.module_names(), f"{label}: name lost on add"
        project_stream = wb._get_cfb().get_stream("PROJECT")  # pyright: ignore[reportPrivateUsage]
        assert name.encode(encoding) in project_stream, (
            f"{label}: PROJECT declaration lacks the cp{code_page} name bytes"
        )
        assert b"Module=?" not in project_stream, (
            f"{label}: PROJECT declaration was substituted to '?'"
        )

    renamed = name + "2"
    with ExcelFile(workbook) as wb:
        wb.vba_project().rename_module(name, renamed)
        wb.save()

    with ExcelFile(workbook) as wb:
        assert renamed in wb.module_names(), f"{label}: name lost on rename"
        assert name not in wb.module_names()
        project_stream = wb._get_cfb().get_stream("PROJECT")  # pyright: ignore[reportPrivateUsage]
        assert renamed.encode(encoding) in project_stream
        assert wb.validate() == []


def test_unencodable_characters_degrade_only_in_the_ansi_records(
    tmp_path: Path,
) -> None:
    """Characters outside the project's page become '?' in the ANSI
    records while the UTF-16 records keep them intact -- the documented
    degradation path, pinned so it stays deliberate."""
    from pyopenvba.vba import parse_dir_stream, serialize_dir_modules_section

    workbook = _workbook_with_code_page(tmp_path, 1252)
    with ExcelFile(workbook) as wb:
        project = wb.vba_project()
        # Cyrillic in a cp1252 project: unrepresentable in ANSI.
        module = project.get_module("Module1")
        module.name = "Тест"
        module.name_unicode = "Тест"
        module.stream_name = "Module1"
        module.stream_name_unicode = "Module1"
        block = serialize_dir_modules_section(project)

    assert "Тест".encode("utf-16-le") in block, "Unicode record must be lossless"
    assert b"?" * 4 in block, "ANSI record is expected to degrade to '?'"

    # And the parser prefers the lossless Unicode record when the ANSI
    # side came back as replacement characters (issue #12).
    prefix = b"\x0f\x00\x02\x00\x00\x00\x01\x00\x13\x00\x02\x00\x00\x00\x00\x00"
    _info, modules = parse_dir_stream(prefix + block)
    assert any(m.name == "Тест" for m in modules)


# --- Access ------------------------------------------------------------------
# The same sweep against the Jet/ACE write path.  Access keeps its VBA in
# MSysAccessStorage rows rather than a CFB file, and that writer read no
# PROJECTCODEPAGE at all until issue #18: every ANSI string went through
# latin-1, so writing a module died outright on anything latin-1 cannot
# hold, and cp1252's 0x80-0x9F punctuation came back as the wrong
# characters.  These run in the same CI job as the Excel sweep above.


def _database_with_code_page(tmp_path: Path, code_page: int) -> Path:
    """A copy of the shipped Access template declaring ``code_page``.

    Same trick as :func:`_workbook_with_code_page`: patch the one dir
    record rather than carry a binary fixture per language.
    """
    target = tmp_path / f"cp{code_page}.accdb"
    shutil.copyfile(_ACCESS_TEMPLATE, target)

    db = AccessDatabase(target)
    rid, dir_stream = db._vba_dir()  # pyright: ignore[reportPrivateUsage]
    db.table(STORAGE_TABLE).update_row(
        rid, {"Lv": compress(_patch_dir_code_page(dir_stream, code_page))}
    )
    db.save()

    assert AccessReader(target).read_project_info().code_page == code_page
    return target


def _access_encoding(path: Path) -> str:
    db = AccessDatabase(path)
    return encoding_of(db._vba_dir()[1])  # pyright: ignore[reportPrivateUsage]


@pytest.mark.parametrize(("code_page", "label", "sample"), LANGUAGE_MATRIX, ids=_IDS)
def test_access_module_source_survives_full_database_cycle(
    tmp_path: Path, code_page: int, label: str, sample: str
) -> None:
    """Write a module carrying the sample, save, reopen, read it back."""
    database = _database_with_code_page(tmp_path, code_page)
    source = (
        f"' {sample}\r\n"
        "Public Sub Probe()\r\n"
        f'    Dim s As String: s = "{sample}"\r\n'
        "End Sub\r\n"
    )

    db = AccessDatabase(database)
    db.set_module_source("Module1", source)
    db.save()

    back = AccessDatabase(database).module("Module1").source
    assert _nfc(sample) in _nfc(back), (
        f"{label} (cp{code_page}) source did not survive the round trip"
    )
    assert "?" not in back.replace("Probe()", ""), (
        f"{label} (cp{code_page}) source was substituted on write"
    )


@pytest.mark.parametrize(("code_page", "label", "name"), NAME_MATRIX, ids=_NAME_IDS)
def test_access_native_module_names_survive_add_and_rename(
    tmp_path: Path, code_page: int, label: str, name: str
) -> None:
    """A native-language module name has to reach the dir stream as real
    code-page bytes, not as a row of '?'."""
    database = _database_with_code_page(tmp_path, code_page)
    encoding = _access_encoding(database)

    db = AccessDatabase(database)
    db.create_module(name, "Option Compare Database")
    db.save()

    db = AccessDatabase(database)
    assert name in [module.name for module in db.modules()], f"{label}: name lost on add"
    dir_stream = db._vba_dir()[1]  # pyright: ignore[reportPrivateUsage]
    assert encode_mbcs(name, encoding) in dir_stream, (
        f"{label}: the dir stream lacks the cp{code_page} name bytes"
    )
    assert b"?" * len(name) not in dir_stream, f"{label}: the name was substituted to '?'"

    renamed = name + "2"
    db.rename_module(name, renamed)
    db.save()

    db = AccessDatabase(database)
    names = [module.name for module in db.modules()]
    assert renamed in names, f"{label}: name lost on rename"
    assert name not in names
    assert encode_mbcs(renamed, encoding) in db._vba_dir()[1]  # pyright: ignore[reportPrivateUsage]

    db.delete_module(renamed)
    db.save()
    assert renamed not in [module.name for module in AccessDatabase(database).modules()]


def test_access_cp1252_punctuation_is_the_regression_this_found(tmp_path: Path) -> None:
    """The half of issue #18 that reaches ordinary English projects.

    latin-1 and cp1252 agree everywhere except 0x80-0x9F, which is where
    the em dash, the curly quotes, the ellipsis and the euro sign live.
    Encoding as latin-1 raised outright; the danger if it had not is that
    reading is byte-lossless either way, so an untouched round trip hides
    the damage until something re-encodes the text.
    """
    database = _database_with_code_page(tmp_path, 1252)
    sample = "em dash — quotes “q” ellipsis … euro €"
    source = f"' {sample}\r\nPublic Sub Probe()\r\nEnd Sub\r\n"

    db = AccessDatabase(database)
    db.set_module_source("Module1", source)
    db.save()

    back = AccessDatabase(database).module("Module1").source
    assert sample in back
    assert "?" not in back

    # And the bytes on disk are the single cp1252 ones, not UTF-8 or latin-1's
    # C1 controls: 0x97 em dash, 0x93/0x94 quotes, 0x85 ellipsis, 0x80 euro.
    db = AccessDatabase(database)
    rid, _dir_stream = db._vba_dir()  # pyright: ignore[reportPrivateUsage]
    del rid
    stream_name = db.module("Module1").stream_name
    row = next(
        row
        for _rid, row in db.table(STORAGE_TABLE).rows_with_ids()
        if str(row["Name"]) == stream_name
    )
    raw = row["Lv"]
    assert isinstance(raw, bytes)
    assert b"\x97" in decompress(raw)


def test_access_writes_nothing_when_the_page_cannot_hold_the_text(tmp_path: Path) -> None:
    """A code page that genuinely cannot hold a character folds it to '?',
    which is what the VBE does.  Cyrillic in a cp1252 project degrades;
    the same text in a cp1251 project does not."""
    western = _database_with_code_page(tmp_path, 1252)
    cyrillic = _database_with_code_page(tmp_path, 1251)
    source = "' Проверка\r\nPublic Sub Probe()\r\nEnd Sub\r\n"

    for database in (western, cyrillic):
        db = AccessDatabase(database)
        db.set_module_source("Module1", source)
        db.save()

    assert "?" in AccessDatabase(western).module("Module1").source
    assert "Проверка" in AccessDatabase(cyrillic).module("Module1").source


def test_unknown_code_page_warns_instead_of_silently_mojibaking() -> None:
    with pytest.warns(UserWarning, match="No Python codec for VBA code page"):
        assert _encoding_for_codepage(99999) == "latin-1"


def test_gb18030_resolves_to_its_python_codec() -> None:
    """54936 is the one page VBA hosts write that Python does not spell
    ``cp<N>``; without the alias it fell back to latin-1 mojibake."""
    assert _encoding_for_codepage(54936) == "gb18030"


def test_cp1258_vietnamese_is_the_regression_this_matrix_found() -> None:
    """Pin the specific failure that motivated the encoder: Python's
    charmap codec cannot compose Vietnamese stacked diacritics."""
    sample = "Tiếng Việt thử nghiệm"
    assert sample.encode("cp1258", errors="replace").count(b"?") == 4
    assert b"?" not in encode_mbcs(sample, "cp1258")
    # ệ is stored as precomposed ê (0xEA) plus combining dot-below (0xF2).
    assert encode_mbcs("ệ", "cp1258") == b"\xea\xf2"


def test_every_matrix_code_page_resolves_portably() -> None:
    """Guard against the platform-dependence trap.

    On Windows, CPython falls through to the OS code-page registry, so
    names like ``cp28592`` resolve there but raise LookupError on Linux
    and macOS.  A page that relies on that fallback decodes correctly on
    one platform and turns into latin-1 mojibake on another.  This test
    fails on every platform (not just the affected one) by consulting
    ``encodings.search_function``, the pure-Python registry that is
    identical everywhere.

    The cross-OS ``languages`` CI job caught exactly this on its first
    run: cp10000, cp20866, cp21866, cp28592, and cp28595 passed on
    Windows and failed on ubuntu.
    """
    import encodings

    non_portable: list[str] = []
    for code_page, label, _sample in LANGUAGE_MATRIX:
        name = _encoding_for_codepage(code_page)
        if encodings.search_function(name) is None:
            non_portable.append(f"cp{code_page} ({label}) -> {name!r}")
    assert not non_portable, (
        "these code pages resolve only via a platform-specific codec "
        "registry; add a portable alias to _CODEPAGE_ALIASES: "
        + ", ".join(non_portable)
    )


def test_alias_table_entries_are_all_portable() -> None:
    """Every alias must exist in the pure-Python codec registry."""
    import encodings

    from pyopenvba.vba import _CODEPAGE_ALIASES  # pyright: ignore[reportPrivateUsage]

    broken = [
        f"{cp} -> {name!r}"
        for cp, name in _CODEPAGE_ALIASES.items()
        if encodings.search_function(name) is None
    ]
    assert not broken, f"unresolvable codec aliases: {broken}"
