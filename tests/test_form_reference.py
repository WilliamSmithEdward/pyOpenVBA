"""add_form declares Microsoft Forms, as the editor does with a project's first UserForm.

tests/fixtures/form_reference.json is what scripts/measure_form_reference.py
saw when Excel, Word and PowerPoint each inserted a UserForm through their
editor: the same control reference in all three, record for record. The
library writes that reference, appended after the project's others, but
for its extended libid: the applications name there the .exd cache their
editor keeps in the saving user's Temp folder, and the library repeats the
original libid, which the editor accepts (tests/test_live_form_reference_gate.py).
"""

from __future__ import annotations

import json
import struct
from pathlib import Path
from typing import Any

import pytest

from pyopenvba import ExcelFile, PowerPointFile, VBAProjectError, WordFile
from pyopenvba._references import MSFORMS_GUID, MSFORMS_LIBID
from pyopenvba.cfb import CFB
from pyopenvba.vba import decompress

RECORD: dict[str, dict[str, Any]] = json.loads(
    (Path(__file__).parent / "fixtures" / "form_reference.json").read_text(encoding="utf-8"))
OPENERS: dict[str, tuple[type[ExcelFile] | type[WordFile] | type[PowerPointFile], str]] = {
    "excel": (ExcelFile, "xlsm"), "word": (WordFile, "docm"), "powerpoint": (PowerPointFile, "pptm")}


def _block(project: bytes) -> dict[str, Any]:
    """The saved project's Microsoft Forms reference, taken apart as the measurement takes it apart."""
    raw = decompress(CFB.from_bytes(project).get_stream("dir"))
    records: list[tuple[int, bytes]] = []
    at = 0
    while True:
        rid, size = struct.unpack_from("<HI", raw, at)
        if rid == 0x000F:
            break
        end = at + (12 if rid == 0x0009 else 6 + size)
        records.append((rid, raw[at + 6:end]))
        at = end
    names = [body.decode("latin-1") for rid, body in records if rid == 0x0016]
    start = next(i for i, (rid, body) in enumerate(records) if rid == 0x0016 and body == b"MSForms")
    end = next(i for i in range(start, len(records)) if records[i][0] == 0x0030) + 1
    block = records[start:end]
    twiddled = block[3][1]
    size = int.from_bytes(twiddled[:4], "little")
    extended = block[-1][1]
    ex_size = int.from_bytes(extended[:4], "little")
    return {
        "reference_names": list(dict.fromkeys(names)),
        "record_ids": [f"0x{rid:04X}" for rid, _ in block],
        "name": block[0][1].decode("latin-1"), "name_unicode": block[1][1].decode("utf-16-le"),
        "original_libid": block[2][1].decode("latin-1"),
        "twiddled_libid": twiddled[4:4 + size].decode("latin-1"), "twiddled_reserved": twiddled[4 + size:].hex(),
        "extended_libid": extended[4:4 + ex_size].decode("latin-1"),
        "extended_reserved": extended[4 + ex_size:10 + ex_size].hex(),
        "type_library_guid": extended[10 + ex_size:26 + ex_size].hex(),
        "cookie": int.from_bytes(extended[26 + ex_size:30 + ex_size], "little"),
        "extended_size": len(extended),
        "last_reference": all(rid != 0x0016 for rid, _ in records[end:]),
    }


def _saved(host: str, tmp_path: Path, *forms: str) -> Path:
    opener, suffix = OPENERS[host]
    target = tmp_path / f"form.{suffix}"
    with opener.create_new(target) as office_file:
        for form in forms:
            office_file.add_form(form)
        office_file.save()
    return target


@pytest.mark.parametrize("host", list(RECORD))
def test_the_reference_is_the_one_the_editor_writes(host: str, tmp_path: Path) -> None:
    opener, _ = OPENERS[host]
    with opener(_saved(host, tmp_path, "Wizard")) as office_file:
        written = _block(office_file.vba_project_bytes())
    measured = dict(RECORD[host])
    # The editor's extended libid names its own .exd cache under the saving user's Temp folder, so its length
    # is the user's; the library repeats the original libid, and the record is that libid, the six reserved
    # bytes, the GUID and the cookie.
    assert measured.pop("extended_libid").endswith("#%TEMP%\\VBE\\MSForms.exd#Microsoft Forms 2.0 Object Library")
    measured.pop("extended_size")
    assert written.pop("extended_libid") == MSFORMS_LIBID
    assert written.pop("extended_size") == 4 + len(MSFORMS_LIBID) + 6 + 16 + 4
    assert written == measured


@pytest.mark.parametrize("host", list(RECORD))
def test_references_lists_it_as_a_control_reference(host: str, tmp_path: Path) -> None:
    opener, _ = OPENERS[host]
    with opener(_saved(host, tmp_path, "Wizard")) as office_file:
        forms = [ref for ref in office_file.references() if ref.name == "MSForms"]
    assert [(ref.kind, ref.guid.upper()) for ref in forms] == [("control", MSFORMS_GUID)]


def test_a_second_form_adds_no_second_reference(tmp_path: Path) -> None:
    with ExcelFile(_saved("excel", tmp_path, "Wizard", "Setup")) as workbook:
        assert [ref.name for ref in workbook.references()].count("MSForms") == 1
        workbook.add_form("Third")
        assert [ref.name for ref in workbook.references()].count("MSForms") == 1


def test_a_project_that_references_it_already_is_left_alone(tmp_path: Path) -> None:
    target = tmp_path / "registered.xlsm"
    with ExcelFile.create_new(target) as workbook:
        workbook.add_reference("MSForms", MSFORMS_GUID, 2, 0)
        before = [(ref.name, ref.kind) for ref in workbook.references()]
        workbook.add_form("Wizard")
        assert [(ref.name, ref.kind) for ref in workbook.references()] == before


def test_the_reference_stays_while_a_form_does(tmp_path: Path) -> None:
    with ExcelFile(_saved("excel", tmp_path, "Wizard")) as workbook:
        with pytest.raises(VBAProjectError, match="required by UserForms: Wizard"):
            workbook.remove_reference("MSForms")
