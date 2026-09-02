"""Generate the text collation table from keys the ACE engine wrote.

    powershell -File tests/live_access_test/dao_oracle.ps1 -Command build-collation -Path chars.accdb
    python scripts/generate_access_collation.py chars.accdb

The database holds one indexed row per BMP code point plus composition
samples.  This reads every index entry back with the engine, derives per
code point its collation elements, diacritic weights, kana kind and
unprintable code, writes ``src/pyopenvba/access/_collation_general_legacy.py``,
then re-encodes every string in the database with the freshly written
table and reports any key that does not come back byte for byte.  It also
writes ``tests/live_access_test/collation_samples.json`` -- every
multi-character sample plus every character that carries extra weights or
starts a run -- which the unit tests check without Office.

Dev-only: needs a database built by the oracle, so Windows with the
Access database engine.
"""

from __future__ import annotations

import importlib
import json
import sys
import unicodedata
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from pyopenvba.access import AccessDatabase  # noqa: E402
from pyopenvba.access._index import leaf_entries  # noqa: E402
from pyopenvba.access._rows import split_row  # noqa: E402

MODULE = ROOT / "src" / "pyopenvba" / "access" / "_collation_general_legacy.py"
SAMPLES = ROOT / "tests" / "live_access_test" / "collation_samples.json"


def read_keys(path: Path) -> dict[str, bytes]:
    db = AccessDatabase(path)
    table = db.table("Chars")
    definition = table.definition
    index = table.index("IX_Ch")
    keys: dict[str, bytes] = {}
    for entry in leaf_entries(db.store, index.real.root_page):
        raw = table.fetch_row(entry.page, entry.row)
        assert raw is not None
        text = table.decode(split_row(definition, raw))["Ch"]
        assert isinstance(text, str)
        assert entry.key[0] == 0x7F, entry.key.hex()
        keys[text] = entry.key[1:]
    return keys


def split_key(key: bytes) -> tuple[bytes, list[bytes]]:
    end = key.index(1)
    primary = key[:end]
    rest = key[end + 1 :]
    assert rest[-1] == 0, key.hex()
    rest = rest[:-1]
    return primary, rest.split(b"\x01") if rest else []


def main(argv: list[str]) -> int:
    if len(argv) != 1:
        raise SystemExit(__doc__)
    keys = read_keys(Path(argv[0]))
    singles = {ord(s): k for s, k in keys.items() if len(s) == 1}
    print(f"{len(keys)} strings, {len(singles)} single characters")

    parsed = {cp: split_key(k) for cp, k in singles.items()}
    one_byte = {p[0] for p, _ in parsed.values() if len(p) == 1}

    def elements_of(primary: bytes) -> list[bytes]:
        # A byte that is a whole one-byte code starts a one-byte element
        # unless it leads a two-byte code; the probes showed every
        # two-byte primary made of two one-byte codes is two elements.
        out: list[bytes] = []
        i = 0
        while i < len(primary):
            if primary[i] in one_byte and (i + 1 >= len(primary) or primary[i + 1] in one_byte):
                out.append(primary[i : i + 1])
                i += 1
            else:
                out.append(primary[i : i + 2])
                i += 2
        return out

    runs: list[tuple[int, int, int, int]] = []
    expansions: dict[int, tuple[bytes, ...]] = {}
    weights: dict[int, tuple[int, ...]] = {}
    kana: list[int] = []
    small: list[int] = []
    unprintable: dict[int, int] = {}
    for cp in sorted(parsed):
        primary, sections = parsed[cp]
        elements = elements_of(primary)
        if len(elements) == 1:
            width = len(elements[0])
            value = int.from_bytes(elements[0], "big")
            if runs and runs[-1][0] + runs[-1][1] == cp and runs[-1][3] == width and runs[-1][2] + runs[-1][1] == value:
                first, count, first_key, w = runs[-1]
                runs[-1] = (first, count + 1, first_key, w)
            else:
                runs.append((cp, 1, value, width))
        elif elements:
            expansions[cp] = tuple(elements)
        s1 = sections[0] if sections else b""
        if s1:
            weights[cp] = tuple(s1) + (0,) * (max(len(elements), 1) - len(s1)) if len(s1) <= max(len(elements), 1) else tuple(s1)
            weights[cp] = tuple(weights[cp][: max(len(elements), 1)])
        if len(sections) >= 3 and sections[2]:
            kana.append(cp)
            if sections[2].startswith(b"\xa0"):
                small.append(cp)
        if len(sections) >= 4 and sections[3]:
            block = sections[3]
            assert len(block) == 4 and block[0] == 0x80 and block[1] == 7 and block[2] == 6, (hex(cp), block.hex())
            unprintable[cp] = block[3]
        if len(sections) >= 2 and sections[1]:
            raise SystemExit(f"U+{cp:04X} uses section 2, which the model does not know")

    # The weight a combining mark contributes when it is the first mark on
    # a base with no precomposed form: read off any precomposed character
    # that decomposes to an unweighted base plus that mark.
    attached: dict[int, int] = {}
    for cp, ws in sorted(weights.items()):
        decomposed = unicodedata.normalize("NFD", chr(cp))
        if len(decomposed) == 2 and unicodedata.combining(decomposed[1]) and ord(decomposed[0]) not in weights:
            mark = ord(decomposed[1])
            if len(ws) == 1 and ws[0]:
                attached.setdefault(mark, ws[0])

    kana_ranges: list[tuple[int, int]] = []
    for cp in kana:
        if kana_ranges and kana_ranges[-1][1] + 1 == cp:
            kana_ranges[-1] = (kana_ranges[-1][0], cp)
        else:
            kana_ranges.append((cp, cp))

    lines = [
        '"""Collation table for sort order 1033 version 0, generated by',
        "scripts/generate_access_collation.py from keys the ACE engine wrote.",
        "Do not edit by hand; regenerate from a database the oracle built.",
        '"""',
        "",
        "from __future__ import annotations",
        "",
        "# (first code point, count, key of the first, key width in bytes):",
        "# consecutive code points whose one-element primary keys are consecutive.",
        "PRIMARY_RUNS: tuple[tuple[int, int, int, int], ...] = (",
    ]
    lines += [f"    ({first:#06x}, {count}, {key:#x}, {width})," for first, count, key, width in runs]
    lines += [")", "", "# Code points that produce several elements.", "EXPANSIONS: dict[int, tuple[bytes, ...]] = {"]
    lines += [
        f"    {cp:#06x}: ({', '.join(repr(e) for e in elements)},),"
        for cp, elements in sorted(expansions.items())
    ]
    lines += ["}", "", "# Diacritic weight per element, 0 for an element without one.", "WEIGHTS: dict[int, tuple[int, ...]] = {"]
    lines += [f"    {cp:#06x}: ({', '.join(f'{w:#04x}' for w in ws)},)," for cp, ws in sorted(weights.items())]
    lines += ["}", "", "KANA_RANGES: tuple[tuple[int, int], ...] = ("]
    lines += [f"    ({a:#06x}, {b:#06x})," for a, b in kana_ranges]
    lines += [")", "", "KANA_SMALL: frozenset[int] = frozenset(("]
    lines += [f"    {cp:#06x}," for cp in small]
    lines += ["))", "", "# Ignorable characters that are still recorded, with their code.", "UNPRINTABLE: dict[int, int] = {"]
    lines += [f"    {cp:#06x}: {code:#04x}," for cp, code in sorted(unprintable.items())]
    lines += ["}", "", "# Weight a combining mark adds as the first mark on a base with no", "# precomposed form, from precomposed letters carrying that mark.", "ATTACHED_WEIGHTS: dict[int, int] = {"]
    lines += [f"    {mark:#06x}: {w:#04x}," for mark, w in sorted(attached.items())]
    lines += ["}", ""]
    MODULE.write_text("\n".join(lines), encoding="utf-8", newline="\n")
    print(
        f"wrote {MODULE.relative_to(ROOT)}: {len(runs)} runs, {len(expansions)} expansions, "
        f"{len(weights)} weighted, {len(kana)} kana ({len(small)} small), {len(unprintable)} unprintable"
    )

    # Verify with the table just written.
    if "pyopenvba.access._collation_general_legacy" in sys.modules:
        importlib.reload(sys.modules["pyopenvba.access._collation_general_legacy"])
    from pyopenvba.access import _collation

    importlib.reload(_collation)
    failures: Counter[str] = Counter()
    shown = 0
    sample_strings: set[str] = set()
    capped = 0
    for text, expected in keys.items():
        if len(expected) >= _collation.MAX_KEY_LENGTH:
            # The engine truncates such keys and leaves bytes behind that
            # follow no rule; the encoder refuses them instead.
            capped += 1
            continue
        try:
            got = _collation.encode_text_key(text)
        except Exception as exc:  # noqa: BLE001
            got = repr(exc).encode()
        if got != expected:
            kind = "single" if len(text) == 1 else "multi"
            failures[kind] += 1
            if shown < 12:
                label = "".join(c if 32 <= ord(c) < 127 else f"<{ord(c):04X}>" for c in text)
                print(f"  MISMATCH {label!s:<24} expected {expected.hex(' ')}\n           {'':<24} got      {got.hex(' ')}")
                shown += 1
        if len(text) != 1:
            if len(expected) < _collation.MAX_KEY_LENGTH:
                sample_strings.add(text)
        elif ord(text) in weights or ord(text) in unprintable or ord(text) in kana or ord(text) in expansions:
            sample_strings.add(text)
    for first, _count, _key, _width in runs:
        sample_strings.add(chr(first))
    print(f"verification: {sum(failures.values())} mismatches {dict(failures)} over {len(keys)} strings ({capped} past the engine's key cap, skipped)")
    SAMPLES.write_text(
        json.dumps({s: keys[s].hex() for s in sorted(sample_strings)}, ensure_ascii=True, indent=0, sort_keys=True),
        encoding="utf-8",
        newline="\n",
    )
    print(f"wrote {SAMPLES.relative_to(ROOT)}: {len(sample_strings)} samples")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
