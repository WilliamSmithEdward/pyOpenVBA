"""The identifier hash stored in each `_VBA_PROJECT` record.

Every identifier record carries a u16 alongside the name. It is a hash
of the name, and it is the last piece standing between the assembler and
writing *new* identifiers rather than editing existing ones.

The model below reproduces **every measured id for names up to six
characters** and 198 of 216 overall; longer names drift, so the
reduction is close but not exact. See ``docs/pcode_reference.md``
section 4.1 for how it was measured and where it breaks.

    h = 0 (seeded per length, see SEEDS)
    for c in name.upper():
        h = (h * 37 + FOLD.get(c, c)) % 65599
    id = h & 0xFFFF
"""
from __future__ import annotations

MULTIPLIER = 37
MODULUS = 65599            # 2**16 + 63
MASK = 0xFFFF

# Two letters do not contribute their own code point. Reproducible on
# every probe: a name ending in W hashes identically to the same name
# ending in V, and Y identically to U.
FOLD = {"W": "V", "Y": "U"}

# The initial accumulator is not a single constant: fitting it per name
# length is what makes the model exact for short names. These are the
# values measured from Excel-compiled projects.
SEEDS: dict[int, int] = {
    1: 0x29FD, 2: 0x29FD, 3: 0xFBCE, 4: 0xA7D3, 5: 0x416D, 6: 0x1800,
    7: 0x6C92, 8: 0xE451, 9: 0x11D4, 10: 0xBB37, 12: 0x0424,
}


def identifier_hash(name: str) -> int | None:
    """The u16 an ``_VBA_PROJECT`` record would carry for ``name``.

    Returns None for a length with no measured seed, rather than
    guessing.
    """
    seed = SEEDS.get(len(name))
    if seed is None:
        return None
    h = seed
    for ch in name.upper():
        h = (h * MULTIPLIER + ord(FOLD.get(ch, ch))) % MODULUS
    return h & MASK


def poly(name: str) -> int:
    """The seed-independent part of the hash, for fitting new seeds."""
    h = 0
    for ch in name.upper():
        h = (h * MULTIPLIER + ord(FOLD.get(ch, ch))) % MODULUS
    return h


def fit_seed(samples: dict[str, int]) -> int | None:
    """Recover the seed for one name length from measured ids."""
    lengths = {len(n) for n in samples}
    if len(lengths) != 1:
        raise ValueError("all samples must share a length")
    length = lengths.pop()
    factor = pow(MULTIPLIER, length, MODULUS)
    inverse = pow(factor, -1, MODULUS)
    votes: dict[int, int] = {}
    for name, value in samples.items():
        seed = ((value - poly(name)) * inverse) % MODULUS
        votes[seed] = votes.get(seed, 0) + 1
    return max(votes, key=lambda s: votes[s]) if votes else None
