"""Measure the `_VBA_PROJECT` identifier hash and check the model.

Compiles families of names through Excel, reads back the u16 each
identifier record carries, and verifies :func:`pcode_hash.identifier_hash`
reproduces every one.

    python docs/research/pcode/hash_probe.py            # compile + check
    python docs/research/pcode/hash_probe.py --check    # check the cache

The probe covers, per name length: an all-'a' base plus single-character
sweeps at several positions (which over-determine the length's seed), a
full identifier-character sweep at a fixed position (which pins the
character map), and case variants. ``identifier_hashes.json`` caches the
measured samples so the check runs without Office.

Dev-only: measuring needs Windows, desktop Excel and `pyvbaharness`.
"""
from __future__ import annotations

import json
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from pcode_hash import fit_seed, identifier_hash

CACHE = "identifier_hashes.json"
ALNUM = "abcdefghijklmnopqrstuvwxyz0123456789_"
FIRST = "abcdefghijklmnopqrstuvwxyz_"


def probe_names() -> list[str]:
    names: list[str] = []
    for length in range(1, 31):
        base = "a" * length
        names.append(base)
        for pos in {0, 1, length - 1}:
            for ch in "abcdefgh":
                names.append(base[:pos] + ch + base[pos + 1:])
    names += ["aa" + c + "aa" for c in ALNUM]          # character map
    names += [c + "aaaa" for c in FIRST]               # legal first chars
    names += ["Alpha", "ALPHA", "alpha", "AlPhA"]      # case folding
    seen, uniq = set(), []
    for n in names:
        if n.lower() not in seen and (n[0].isalpha() or n[0] == "_"):
            seen.add(n.lower())
            uniq.append(n)
    return uniq


def _module(chunk: list[str]) -> str:
    body = "Sub A()\r\n" + "".join(f"    Dim {n}\r\n" for n in chunk)
    body += "".join(f"    {n} = 1\r\n" for n in chunk)
    return body + "End Sub\r\n"


def measure() -> dict[str, int]:
    from batch import compile_many, hdr, scratch_dir
    from pcode_names import parse_identifiers

    from pyopenvba import ExcelFile
    from pyopenvba.cfb import CFB

    names = probe_names()
    chunks = {f"h{i}": _module(names[i:i + 28])
              for i in range(0, len(names), 28)}
    compile_many({k: hdr(v) for k, v in chunks.items()})
    table: dict[str, int] = {}
    for tag in chunks:
        with ExcelFile(scratch_dir() / f"v_{tag}.xlsm") as workbook:
            cfb = CFB.from_bytes(workbook.vba_project_bytes())
        stream = cfb.get_stream_in_storage("VBA", "_VBA_PROJECT")
        for ident in parse_identifiers(stream):
            table[ident.name] = ident.id_value
    return table


def check(table: dict[str, int]) -> int:
    by_length: dict[int, dict[str, int]] = defaultdict(dict)
    for name, value in table.items():
        by_length[len(name)][name] = value
    exact = total = 0
    for length in sorted(by_length):
        samples = by_length[length]
        hits = sum(1 for n, v in samples.items() if identifier_hash(n) == v)
        exact += hits
        total += len(samples)
        note = ""
        if hits != len(samples):
            seed = fit_seed(samples)
            note = (f"   <- {hits}/{len(samples)}; a fitting representative "
                    f"is {seed:#010x}" if seed else "   <- no seed fits")
        print(f"  len {length:2}  {hits}/{len(samples)} exact{note}")
    print(f"\n{exact}/{total} identifier hashes reproduced")
    return total - exact


if __name__ == "__main__":
    cache = Path(__file__).parent / CACHE
    if "--check" in sys.argv and cache.exists():
        data = json.loads(cache.read_text())
    else:
        data = measure()
        cache.write_text(json.dumps(data, indent=0, sort_keys=True))
    raise SystemExit(1 if check(data) else 0)
