"""Text sort keys: the bytes the ACE engine stores for a text value in an
index, for sort order 1033 version 0 (the "General" order of Jet 4 and
Access 2007 files).

The engine's rules, all measured on keys it wrote (see
``docs/access_engine.md``):

* A string is trimmed of trailing spaces and composed (a base letter plus
  a combining mark keys like the precomposed letter).
* Each character maps to zero or more *elements*.  Most letters are one
  element of one or two bytes; expansions such as sharp s (``ss``) or the
  ``ffi`` ligature are several elements; 19 585 code points are ignored.
  Case is not encoded at all.
* The key is the element bytes, ``0x01``, then up to four sections
  separated by ``0x01`` with trailing empty ones omitted, then ``0x00``:

  1. one diacritic weight per element, ``0x02`` standing in for elements
     without one, trailing stand-ins trimmed;
  2. never seen non-empty;
  3. for kana: a bit stream (``10`` then ``11`` per full-size kana and
     ``10`` per small one, cut after the last small one, zero-padded)
     followed by ``ff 02 80 ff 80``, and ``ff`` more when section 4 follows;
  4. for each ignorable-but-recorded character (hyphen, apostrophe,
     controls): ``80 <7 + 4 * elements before it> 06 <code>``.

The table itself is generated from the engine's own output by
``scripts/generate_access_collation.py`` into ``_collation_general_legacy``.
"""

from __future__ import annotations

import bisect
import unicodedata
from dataclasses import dataclass

from pyopenvba.access_read import AccessError
from pyopenvba.access import _collation_general_legacy as table

TEXT_END = 0x01
SECTION_END = 0x01
KEY_END = 0x00
WEIGHT_PLACEHOLDER = 0x02
KANA_SUFFIX = bytes((0xFF, 0x02, 0x80, 0xFF, 0x80))
KANA_BEFORE_UNPRINTABLE = 0xFF
UNPRINTABLE_LEAD = 0x80
UNPRINTABLE_BASE = 7
UNPRINTABLE_STEP = 4
UNPRINTABLE_MID = 0x06
# The engine stores at most 510 key bytes including the flag byte and cuts
# longer keys without a clean terminator; 509 is the limit on what this
# module produces.
MAX_KEY_LENGTH = 509

_RUN_STARTS = [run[0] for run in table.PRIMARY_RUNS]


@dataclass(frozen=True)
class CharacterKey:
    """What one character contributes: its elements as (primary bytes,
    diacritic weight or 0) pairs, whether it is a kana and a small one,
    and its code when it is an unprintable that is still recorded."""

    elements: tuple[tuple[bytes, int], ...]
    kana: bool
    small_kana: bool
    unprintable: int | None


_EMPTY = CharacterKey((), False, False, None)


SPACE = 0x20
SPACE_ELEMENT = (bytes((0x07,)), 0)


def character_key(code_point: int) -> CharacterKey:
    if code_point == SPACE:
        # Trailing spaces are trimmed before encoding; any other space is
        # an element of its own.  A lone space keys as empty, so the
        # generated table cannot carry this and the encoder does.
        return CharacterKey((SPACE_ELEMENT,), False, False, None)
    if code_point in table.EXPANSIONS:
        primaries: tuple[bytes, ...] = table.EXPANSIONS[code_point]
    else:
        primaries = _run_primary(code_point)
    weights = table.WEIGHTS.get(code_point, ())
    if not primaries and weights:
        # A bare combining mark: a weight with nothing under it.
        primaries = (b"",)
    elements = tuple(
        (primary, weights[i] if i < len(weights) else 0) for i, primary in enumerate(primaries)
    )
    kana = _in_ranges(code_point, table.KANA_RANGES)
    small = code_point in table.KANA_SMALL
    unprintable = table.UNPRINTABLE.get(code_point)
    if not elements and not kana and unprintable is None:
        return _EMPTY
    return CharacterKey(elements, kana, small, unprintable)


def _run_primary(code_point: int) -> tuple[bytes, ...]:
    index = bisect.bisect_right(_RUN_STARTS, code_point) - 1
    if index < 0:
        return ()
    first, count, first_key, width = table.PRIMARY_RUNS[index]
    if code_point >= first + count:
        return ()
    return ((first_key + (code_point - first)).to_bytes(width, "big"),)


def _in_ranges(code_point: int, ranges: tuple[tuple[int, int], ...]) -> bool:
    for start, end in ranges:
        if start <= code_point <= end:
            return True
    return False


def compose_marks(text: str) -> list[tuple[str, list[str]]]:
    """Group each character with the combining marks that follow it,
    folding a mark into the base when a precomposed character exists,
    which is what the engine does.  Whole-string NFC would be wrong: it
    also rewrites singletons such as U+0387 to U+00B7, and the engine
    keys those as themselves.  Marks that do not compose stay listed
    with their base; the encoder adds their weights."""
    out: list[tuple[str, list[str]]] = []
    for char in text:
        if out and unicodedata.combining(char):
            base, marks = out[-1]
            composed = unicodedata.normalize("NFC", base + char)
            if len(composed) == 1 and not marks:
                out[-1] = (composed, marks)
            else:
                marks.append(char)
            continue
        out.append((char, []))
    return out


def mark_weight(mark: str, first: bool) -> int:
    """The weight a combining mark adds to the element before it.  The
    first mark on a base that has no precomposed form takes the weight
    that mark gives any precomposed letter; a further mark adds the
    weight it carries standing alone (measured: a + acute + grave is
    0x0e + 0x0d, a + grave + grave is 0x0f + 0x0d)."""
    code_point = ord(mark)
    if first and code_point in table.ATTACHED_WEIGHTS:
        return table.ATTACHED_WEIGHTS[code_point]
    weights = table.WEIGHTS.get(code_point, ())
    return weights[0] if weights else 0


def encode_text_key(text: str) -> bytes:
    """The stored key for ``text`` in an ascending text index column,
    without the leading flag byte."""
    primary = bytearray()
    weights: list[int] = []
    kana_pairs: list[int] = []
    last_small = -1
    unprintables = bytearray()
    element_count = 0
    for char, marks in compose_marks(text.rstrip(" ")):
        code_point = ord(char)
        if code_point > 0xFFFF:
            raise AccessError(
                f"U+{code_point:X}: characters outside the Basic Multilingual Plane "
                "have not been measured in an index key"
            )
        info = character_key(code_point)
        for element_primary, weight in info.elements:
            primary += element_primary
            weights.append(weight)
            element_count += 1
        if marks and info.elements:
            composed_already = unicodedata.normalize("NFD", char) != char
            for i, mark in enumerate(marks):
                weights[-1] += mark_weight(mark, first=(i == 0 and not composed_already))
        elif marks:
            for mark in marks:
                for element_primary, weight in character_key(ord(mark)).elements:
                    primary += element_primary
                    weights.append(weight)
                    element_count += 1
        if info.kana:
            kana_pairs.append(0b10 if info.small_kana else 0b11)
            if info.small_kana:
                last_small = len(kana_pairs) - 1
        if info.unprintable is not None:
            unprintables += bytes(
                (
                    UNPRINTABLE_LEAD,
                    UNPRINTABLE_BASE + UNPRINTABLE_STEP * element_count,
                    UNPRINTABLE_MID,
                    info.unprintable,
                )
            )

    section1 = bytearray(w if w else WEIGHT_PLACEHOLDER for w in weights)
    while section1 and section1[-1] == WEIGHT_PLACEHOLDER:
        section1.pop()
    section3 = bytearray()
    if kana_pairs:
        if last_small >= 0:
            section3 += _pack_kana(kana_pairs[: last_small + 1])
        section3 += KANA_SUFFIX
        if unprintables:
            section3.append(KANA_BEFORE_UNPRINTABLE)

    key = bytearray(primary)
    key.append(TEXT_END)
    key += section1
    if section3 or unprintables:
        key += bytes((SECTION_END, SECTION_END))  # section 2 is always empty
        key += section3
    if unprintables:
        key.append(SECTION_END)
        key += unprintables
    key.append(KEY_END)
    if len(key) > MAX_KEY_LENGTH:
        raise AccessError(
            f"text key of {len(key)} bytes exceeds the engine's {MAX_KEY_LENGTH}-byte limit; "
            "the engine truncates such keys in a way that has not been reproduced"
        )
    return bytes(key)


def _pack_kana(pairs: list[int]) -> bytes:
    """Three two-bit kana codes per byte behind a ``10`` marker, the last
    byte zero-padded: measured as a0 for one small kana, b8 for a full
    then a small, bf a0 for three full then a small."""
    out = bytearray()
    for i in range(0, len(pairs), 3):
        chunk = pairs[i : i + 3] + [0] * (3 - len(pairs[i : i + 3]))
        out.append(0b10 << 6 | chunk[0] << 4 | chunk[1] << 2 | chunk[2])
    return bytes(out)
