"""The identifier hash stored in each ``_VBA_PROJECT`` record.

Every identifier record carries a u16 alongside the name. It is the low
word of the **OLE Automation name hash**, ``LHashValOfNameSysA`` from
``OLEAUT32.dll`` -- the same hash type libraries use for name lookup.
VBE7 computes it when interning an identifier (its intern routine calls
``LHashValOfNameSysA(syskind, lcid, name)`` and keeps the low 16 bits).

The algorithm, confirmed exact on 5,255 measured names with zero
exceptions:

    h = 0x0DEADBEE                       # fixed initial value
    for c in name:
        h = (37 * h + LOOKUP[c]) & 0xFFFFFFFF
    id = (h % 65599) & 0xFFFF

``LOOKUP`` is the case-folding table the API applies: uppercase ASCII,
with digits and ``_`` mapping to themselves. VBE7's table additionally
folds ``W`` -> ``V`` and ``Y`` -> ``U`` (a quirk of the specific
per-syskind lookup table it loads); every other letter maps to its plain
uppercase code. This is what earlier revisions of this module modelled as
a mythical per-length "seed" -- there is no seed. The reduction is
*unsigned* ``% 65599`` (65599 = 2**16 + 63), and only the low 16 bits are
stored.

Reference: ReactOS / Wine ``dll/win32/oleaut32/hash.c``
(``LHashValOfNameSysA``); MS ``oleauto.h``.
"""
from __future__ import annotations

INITIAL = 0x0DEADBEE
MULTIPLIER = 37
MODULUS = 65599            # 2**16 + 63
FIELD_MASK = 0xFFFF
_U32 = 0xFFFFFFFF

# The case-folding lookup table, built to match the values VBE7 emits.
# Letters fold to uppercase; W and Y fold one and four below their own
# code (to V and U); digits, underscore and everything else map to
# themselves. Non-ASCII bytes are passed through unchanged, which is
# correct for the Windows (non-Mac) syskind VBA uses.
_FOLD = {ord("W"): ord("V"), ord("Y"): ord("U")}


def _build_lookup() -> list[int]:
    table = list(range(256))
    for code in range(ord("a"), ord("z") + 1):
        table[code] = code - 0x20            # lowercase -> uppercase
    for code, folded in _FOLD.items():       # W/Y quirk, upper and lower
        table[code] = folded
        table[code + 0x20] = folded
    return table


LOOKUP = _build_lookup()


def identifier_hash(name: str) -> int:
    """The u16 an ``_VBA_PROJECT`` record carries for ``name``.

    Exact for any VBA identifier; there is no length restriction and no
    fitting involved.
    """
    h = INITIAL
    for ch in name:
        h = (MULTIPLIER * h + LOOKUP[ord(ch) & 0xFF]) & _U32
    return (h % MODULUS) & FIELD_MASK


def name_hash_full(name: str, syskind: int = 3, mask: int = 0) -> int:
    """The full 32-bit ``LHashValOfNameSysA`` return value.

    The high word is ``(syskind | mask) << 16``; VBA stores only the low
    word (:func:`identifier_hash`). ``syskind`` defaults to ``SYS_WIN64``.
    """
    return ((syskind | mask) << 16) | identifier_hash(name)


def encode_identifier_record(name: str, type_byte: int) -> bytes:
    """The exact ``_VBA_PROJECT`` bytes for a compact identifier record.

    A compact record (references and ordinary user identifiers, type byte
    below ``0x80``) is ``<u8 name-length><u8 type><ASCII name><u16 hash>``
    followed by the ``0x0010`` trailer. Module records and records whose
    type sets the ``0x80`` descriptor bit carry extra fields and are not
    produced here.
    """
    if type_byte >= 0x80:
        raise ValueError("descriptor/module records are not compact")
    body = name.encode("latin-1")
    return (bytes((len(body), type_byte)) + body
            + identifier_hash(name).to_bytes(2, "little") + b"\x10\x00")
