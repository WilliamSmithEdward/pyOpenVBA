"""Name the property codes by what they hold, not by where they sit.

A straight walk of the blob against SaveAsText drifts: the blob carries
records the text does not write.  Matching on the value instead is
decisive -- a property whose value appears exactly once on each side of
one object can only be that record -- and a name is kept only when every
object that carries it agrees and nothing already named is contradicted.
"""

import struct
import pathlib
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[3] / "src"))

from pyopenvba.access import AccessDatabase  # noqa: E402
from pyopenvba.access._designs import PROPERTY_CODES  # noqa: E402

#: Where build_rich_form.py left the database and its text export.
HERE = Path(".").resolve()
NAME_CODE = 20


def parse(lines: list[str]) -> dict[str, list[tuple[str, str]]]:
    """Each named `Begin ... End` block's `(property, value)` pairs."""
    out: dict[str, list[tuple[str, str]]] = {}
    stack: list[list[tuple[str, str]]] = []
    names: list[str | None] = []
    skip = 0
    for raw in lines:
        line = raw.strip()
        if not line:
            continue
        if skip:
            if line == "End":
                skip -= 1
            continue
        if line == "Begin" or line.startswith("Begin "):
            stack.append([])
            names.append(None)
            continue
        if line == "End":
            if stack:
                block = stack.pop()
                name = names.pop()
                if name:
                    out[name] = block
            continue
        key, sep, value = line.partition("=")
        key, value = key.strip(), value.strip()
        if not key or not sep or not stack:
            continue
        stack[-1].append((key, value))
        if key == "Name":
            names[-1] = value.strip('"')
        if value == "Begin":
            skip += 1
    return out


def readings(value: bytes) -> set[str]:
    """Every way the record's bytes could be written in the text."""
    out: set[str] = set()
    if not value:
        return out
    try:
        text = value.decode("utf-16-le")
        if text.isprintable():
            out.add(text)
            out.add(f'"{text}"')
    except UnicodeDecodeError:
        pass
    for size, code in ((1, "B"), (2, "<H"), (4, "<I")):
        if len(value) == size:
            number = struct.unpack(code, value)[0]
            out.add(str(number))
            if size == 4:
                out.add(str(struct.unpack("<i", value)[0]))
    if len(value) == 4:
        out.add(f"{struct.unpack('<f', value)[0]:g}")
    return out


def main() -> None:
    text = (HERE / "rich.txt").read_text(encoding="utf-16")
    blocks = parse(text.splitlines())

    db = AccessDatabase(HERE / "rich.accdb")
    objects = list(db.form("Form1").objects)

    votes: dict[int, dict[str, int]] = defaultdict(lambda: defaultdict(int))
    for obj in objects:
        found = [r for r in obj.records if r.code == NAME_CODE]
        if not found:
            continue
        pairs = blocks.get(found[0].value.decode("utf-16-le"))
        if not pairs:
            continue
        # A value that shows up once on each side names its record.
        for prop, written in pairs:
            hits = [r for r in obj.records if written in readings(r.value)]
            if len(hits) != 1:
                continue
            same = [p for p, w in pairs if w == written]
            if len(same) != 1:
                continue
            votes[hits[0].code][prop] += 1

    agreed = {
        next(iter(counts)): code for code, counts in votes.items() if len(counts) == 1
    }
    known = {n: c for n, c in agreed.items() if n in PROPERTY_CODES}
    wrong = {n: (c, PROPERTY_CODES[n]) for n, c in known.items() if PROPERTY_CODES[n] != c}
    fresh = {n: c for n, c in agreed.items() if n not in PROPERTY_CODES}
    # A code cannot mean two things, so a new name landing on a code we
    # already call something else is a mistake, not a discovery.
    taken = {c: n for n, c in PROPERTY_CODES.items()}
    clashes = {n: (c, taken[c]) for n, c in fresh.items() if c in taken}

    print(f"undisputed: {len(agreed)}  |  reproduced already-named: {len(known) - len(wrong)}")
    print(f"contradicted: {wrong or 'none'}")
    print(f"clashes with a code already named: {clashes or 'none'}")
    print(f"\nnew names: {len(fresh)}")
    for name, code in sorted(fresh.items(), key=lambda kv: kv[1]):
        print(f'    "{name}": {code},')

    out = HERE / "codes.txt"
    with out.open("w", encoding="utf-8") as fh:
        for name, code in sorted(fresh.items(), key=lambda kv: kv[1]):
            if name not in clashes:
                fh.write(f'    "{name}": {code},\n')
    print("\nwritten to", out)


if __name__ == "__main__":
    main()
