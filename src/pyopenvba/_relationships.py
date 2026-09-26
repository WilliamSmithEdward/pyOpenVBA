"""The order Office writes a part's relationships in.

Excel, and PowerPoint with it, writes each ``.rels`` part in the order a
linear-hashing table keyed by the relationship ids iterates. Measured by
saving sheets holding every count of hyperlinks from 1 to 160 and larger
counts up to 65,530, the most a sheet holds, one relationship each
(scripts/measure_relationship_order.py):

- An id's hash is its characters read as a number in base 101, the
  textbook string hash, put through one step of C's rand() generator:
  (h * 1103515245 + 12345) >> 16, in 32 bits.
- The table starts with 8 buckets. Before an insert that would put more
  than 4 ids per bucket it splits one bucket, p, the next in turn, into p
  and p + L (L the buckets before this round of splits began), the ids
  whose next hash bit is set moving to the new bucket, in their order. An
  id's bucket is its hash mod L, or mod 2L while that is below the split
  pointer.
- It is read bucket by bucket, each bucket's ids in the order they went
  in; Office puts them in as rId1, rId2 and so on.

This reproduces all 177 measured orders, and each part Office wrote in the
tests' fixtures. They reach 16,384 buckets, the hash's low 14 bits; a
part of more than 65,536 relationships, whose order they cannot vouch
for, keeps the order it has, as does one whose ids are not all rId and a
number.
"""

from __future__ import annotations

import math
import re
from typing import Final

#: The buckets the table starts with, and the ids per bucket past which it splits one.
_START: Final = 8
_LOAD: Final = 4
#: The most relationships a part can hold for the order to be the measured one: 16,384 buckets.
_MOST: Final = 65536
_RELATIONSHIP = re.compile(r"<Relationship\b[^>]*?(?:/>|>\s*</Relationship>)")
#: An id Office numbers: rId and a number, written as a number is, with no leading zero.
_ID = re.compile(r'\bId="rId(0|[1-9]\d*)"')


def id_hash(identifier: str) -> int:
    """The hash Office's table places a relationship id by: base 101, then a step of C's rand() generator."""
    h = 0
    for character in identifier:
        h = (h * 101 + ord(character)) & 0xFFFFFFFF
    return ((h * 1103515245 + 12345) & 0xFFFFFFFF) >> 16


def bucket_order(numbers: list[int]) -> list[int] | None:
    """The ids rId<number>, by their numbers, in the order Office writes them; None where the order is not
    known."""
    if not numbers or len(numbers) > _MOST or any(number < 0 for number in numbers):
        return None
    buckets = max(_START, math.ceil(len(numbers) / _LOAD))
    level = _START
    while level * 2 <= buckets:
        level *= 2
    split = buckets - level

    def bucket(number: int) -> int:
        hashed = id_hash(f"rId{number}")
        home = hashed % level
        return hashed % (2 * level) if home < split else home

    return sorted(numbers, key=lambda number: (bucket(number), number))


def in_office_order(text: str) -> str:
    """A ``.rels`` part with its relationships in the order Office writes them, where the order is known."""
    elements = _RELATIONSHIP.findall(text)
    numbers: list[int] = []
    for element in elements:
        found = _ID.search(element)
        if found is None:
            return text
        numbers.append(int(found.group(1)))
    if len(set(numbers)) != len(numbers):
        return text
    order = bucket_order(numbers)
    if order is None:
        return text
    by_number = dict(zip(numbers, elements, strict=True))
    start = text.find(elements[0])
    end = text.rfind(elements[-1]) + len(elements[-1])
    if start < 0 or "".join(elements) != text[start:end]:
        return text
    return text[:start] + "".join(by_number[number] for number in order) + text[end:]
