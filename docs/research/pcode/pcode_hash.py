"""The identifier hash stored in each `_VBA_PROJECT` record.

Every identifier record carries a u16 alongside the name. It is a hash
of the name, and it is the last piece standing between the assembler and
writing *new* identifiers rather than editing existing ones.

The mechanism is fully reverse-engineered and reproduces **every** one
of 1,117 measured ids exactly (see ``docs/pcode_reference.md`` section
4.1 and ``hash_probe.py``):

    h = SEED_BY_LENGTH[len(name)]          # 32-bit, per name length
    for c in name:
        h = (h * 37 + charval(c)) & 0xFFFFFFFF
    id = (signed32(h) % 65599) & 0xFFFF

where ``signed32`` reinterprets the 32-bit accumulator as a signed
integer before the modulo (so the reduction of a value >= 2**31 differs
from the unsigned one -- this is what makes long names diverge from a
naive ``* 37 mod`` hash), and ``charval`` is the uppercased ASCII code
with two fixed folds, ``W -> V`` and ``Y -> U``.

Every constant here was read off measured data, not assumed:

* multiplier 37: the per-position weight of a unit change in a character
  is exactly ``37**k mod 65599`` for the low positions;
* signed 32-bit accumulator: the weights start diverging from
  ``37**k mod 65599`` at exactly ``k = 6``, where ``37**6`` first
  exceeds ``2**31``, and the divergence matches the signed reduction to
  the bit;
* modulus 65599 (``2**16 + 63``): scanning every modulus from 65,500 to
  66,600 puts a sharp maximum here;
* the ``& 0xFFFF``: the field is 16 bits, so a reduced value in
  ``[65536, 65598]`` is stored truncated (``0x10013 -> 0x0013``);
* uppercase with ``W -> V`` / ``Y -> U``: the character sweep gives each
  letter its uppercased ASCII contribution, with ``W`` taking ``V``'s
  value and ``Y`` taking ``U``'s. Digits and ``_`` contribute their own
  ASCII codes.

**Open:** ``SEED_BY_LENGTH`` is a genuine per-length constant, but its
single underlying value is *not recoverable from the ids alone* -- the
reduction collapses the seed's high bits, so thousands of distinct
32-bit seeds reproduce every id of a given length identically. The table
below holds one measured representative per length (1..30), enough to
generate a correct id for any identifier up to 30 characters; longer
lengths extend it by probing (``hash_probe.py``).
"""
from __future__ import annotations

MULTIPLIER = 37
MODULUS = 65599            # 2**16 + 63
FIELD_MASK = 0xFFFF
_U32 = 0xFFFFFFFF
_SIGN = 0x80000000
_WRAP = 0x100000000

# Two letters do not contribute their own code point: W hashes as V, Y as
# U. Reproducible in the raw stream, not a parser artifact.
_FOLD = {ord("W"): ord("V"), ord("Y"): ord("U")}

# One measured 32-bit seed representative per name length. See module
# docstring: the true single seed is output-underdetermined, so these are
# representatives, each of which reproduces every id of its length.
SEED_BY_LENGTH: dict[int, int] = {
    1: 0x94C1BEEB, 2: 0x987F28BD, 3: 0xD986C65E, 4: 0xE4B4C7D2,
    5: 0x5129AE63, 6: 0x12892FB7, 7: 0x8E6E83F7, 8: 0x5154F400,
    9: 0x4A821752, 10: 0x214BA79D, 11: 0x7A58B56B, 12: 0x90E05AC2,
    13: 0x838CDDF5, 14: 0x0C7A4325, 15: 0xC2B77358, 16: 0x8895F898,
    17: 0xA7FBCDC4, 18: 0x9B7E2613, 19: 0xC027CDCB, 20: 0x1EF1F0F0,
    21: 0x5B28485C, 22: 0x5736E9BF, 23: 0x9BA9ADC8, 24: 0xC8492D63,
    25: 0x522C9E2D, 26: 0x6392245B, 27: 0x0D0630D2, 28: 0xCFC863BE,
    29: 0x55FB0355, 30: 0x32843028,
}


def charval(ch: str) -> int:
    """The value a single character contributes: uppercased ASCII, folded."""
    code = ord(ch.upper())
    return _FOLD.get(code, code)


def accumulate(name: str, seed: int) -> int:
    """The raw 32-bit accumulator for ``name`` from ``seed``."""
    h = seed & _U32
    for ch in name:
        h = (h * MULTIPLIER + charval(ch)) & _U32
    return h


def reduce_field(h: int) -> int:
    """Reduce a 32-bit accumulator to the stored 16-bit id."""
    signed = h - _WRAP if h & _SIGN else h
    return (signed % MODULUS) & FIELD_MASK


def identifier_hash(name: str) -> int | None:
    """The u16 an ``_VBA_PROJECT`` record carries for ``name``.

    Returns None for a length with no measured seed representative,
    rather than guessing.
    """
    seed = SEED_BY_LENGTH.get(len(name))
    if seed is None:
        return None
    return reduce_field(accumulate(name, seed))


def fit_seed(samples: dict[str, int]) -> int | None:
    """Recover a working seed representative for one name length.

    ``samples`` maps names of a single length to their measured ids. The
    reduction is not injective in the seed, so many seeds satisfy the
    samples; this returns one that reproduces all of them, or None.
    """
    lengths = {len(n) for n in samples}
    if len(lengths) != 1:
        raise ValueError("all samples must share a length")
    length = lengths.pop()
    if not samples:
        return None
    base, base_id = next(iter(samples.items()))
    factor = pow(MULTIPLIER, length, _WRAP)
    inverse = pow(factor, -1, _WRAP)
    raw_base = accumulate(base, 0)
    for target in (base_id, base_id + (FIELD_MASK + 1)):
        for band in range(-40000, 40001):
            seed = (inverse * ((target - raw_base) % _WRAP
                               + MODULUS * band)) % _WRAP
            if reduce_field(accumulate(base, seed)) != base_id:
                continue
            if all(reduce_field(accumulate(n, seed)) == v
                   for n, v in samples.items()):
                return seed
    return None
