"""The text collation the ACE engine uses for index keys.

``collation_samples.json`` holds keys the engine itself wrote: every
multi-character composition sample, every character carrying extra
weights, kana marks or an unprintable code, and the first character of
every run in the generated table -- 4 084 strings.  The encoder must
reproduce each one byte for byte.  The live gate re-measures all 63 000
code points; this file needs no Office.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.access._collation import (
    MAX_KEY_LENGTH,
    character_key,
    compose_marks,
    encode_text_key,
)
from pyopenvba.access_read import AccessError

SAMPLES = Path(__file__).parent / "live_access_test" / "collation_samples.json"


def _samples() -> dict[str, bytes]:
    return {s: bytes.fromhex(k) for s, k in json.loads(SAMPLES.read_text(encoding="utf-8")).items()}


def test_every_measured_key_is_reproduced() -> None:
    samples = _samples()
    assert len(samples) > 4000
    wrong = {s: (k, encode_text_key(s)) for s, k in samples.items() if encode_text_key(s) != k}
    assert not wrong, f"{len(wrong)} of {len(samples)} keys differ, e.g. {list(wrong.items())[:3]!r}"


def test_case_is_not_stored() -> None:
    assert encode_text_key("Access") == encode_text_key("ACCESS") == encode_text_key("access")
    assert encode_text_key("a") == bytes.fromhex("4a0100")


def test_trailing_spaces_are_trimmed_and_inner_spaces_weigh_07() -> None:
    assert encode_text_key("a   ") == encode_text_key("a")
    assert encode_text_key(" a") == bytes.fromhex("074a0100")
    assert encode_text_key("a  a") == bytes.fromhex("4a07074a0100")
    assert encode_text_key("   ") == bytes.fromhex("0100")
    assert encode_text_key("") == bytes.fromhex("0100")


def test_ignorable_characters_record_their_element_offset() -> None:
    # Hyphen after one element: 7 + 4 * 1 = 0x0b; after the two elements
    # of a sharp s: 0x0f; two hyphens at the same offset repeat it.
    assert encode_text_key("a-") == bytes.fromhex("4a01010101800b068200")
    assert encode_text_key("ß-") == bytes.fromhex("6b6b01010101800f068200")
    assert encode_text_key("a--b") == bytes.fromhex("4a4c01010101800b0682800b068200")


def test_diacritics_take_one_weight_per_element() -> None:
    assert encode_text_key("é") == bytes.fromhex("51010e00")
    assert encode_text_key("aé") == bytes.fromhex("4a5101020e00")
    assert encode_text_key("éa") == bytes.fromhex("514a010e00")
    # A combining acute composes to the same key as the precomposed letter.
    assert encode_text_key("é") == encode_text_key("é")
    # Stacked marks add: a + acute + grave is 0x0e + 0x0d.
    assert encode_text_key("á̀") == bytes.fromhex("4a011b00")


def test_kana_marks_pack_three_per_byte() -> None:
    assert encode_text_key("あ") == bytes.fromhex("7f02010101ff0280ff8000")
    assert encode_text_key("ぁ") == bytes.fromhex("7f02010101a0ff0280ff8000")
    assert encode_text_key("あああぁ") == bytes.fromhex(
        "7f027f027f027f02010101bfa0ff0280ff8000"
    )
    assert encode_text_key("あ-") == bytes.fromhex("7f02010101ff0280ff80ff01800b068200")


def test_two_byte_primary_is_one_element() -> None:
    # Cyrillic pe is two key bytes but one element: the hyphen after it
    # records offset 1, and a following diacritic gets one placeholder.
    assert encode_text_key("п-") == bytes.fromhex("794701010101800b068200")
    assert encode_text_key("пÉ") == bytes.fromhex("79475101020e00")


def test_keys_past_the_engine_limit_are_refused() -> None:
    assert len(encode_text_key("a" * 255)) == 257
    with pytest.raises(AccessError):
        encode_text_key("b" * 254 + "É")
    assert MAX_KEY_LENGTH == 509


def test_astral_characters_are_refused() -> None:
    with pytest.raises(AccessError):
        encode_text_key("\U0001F600")


def test_character_key_and_compose_marks_shapes() -> None:
    assert character_key(ord("a")).elements == ((b"J", 0),)
    assert character_key(0x20).elements == ((b"\x07", 0),)
    assert character_key(0xFFFF).elements == ()
    assert compose_marks("éx") == [("é", []), ("x", [])]
