"""Measure the `_VBA_PROJECT` identifier hash and check the model.

Compiles two families of names through Excel and reads back the u16 each
identifier record carries:

* one base name per length plus a single-character bump at every
  position, which exposes each position's weight;
* a full a-z sweep through each position of a fixed name, which shows
  whether the hash is linear in the character (it is, except for W and
  Y).

    python docs/research/pcode/hash_probe.py            # compile + check
    python docs/research/pcode/hash_probe.py --check    # check cached

Dev-only: needs Windows, desktop Excel and `pyvbaharness`.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from batch import compile_many, hdr, scratch_dir
from pcode_hash import fit_seed, identifier_hash
from pcode_names import parse_identifiers

from pyopenvba import ExcelFile
from pyopenvba.cfb import CFB

LETTERS = "abcdefghijklmnopqrstuvwxyz"
CACHE = "identifier_hashes.json"


def probe_names() -> list[str]:
    names: list[str] = []
    for length in range(1, 11):                 # positional weights
        base = "q" * length
        names.append(base)
        for i in range(length):
            names.append(base[:i] + "z" + base[i + 1:])
    for pos in range(4):                        # linearity in the character
        for ch in LETTERS:
            names.append("qqqq"[:pos] + ch + "qqqq"[pos + 1:])
    names += ["alpha", "Bravo", "charlie", "delta9", "foxtrot_1", "golf",
              "hotel", "india", "juliet", "kilo", "lima"]
    seen, unique = set(), []
    for name in names:
        if name.lower() not in seen:
            seen.add(name.lower())
            unique.append(name)
    return unique


def _module(chunk: list[str]) -> str:
    body = "Sub A()\r\n"
    body += "".join(f"    Dim {n}\r\n" for n in chunk)
    body += "".join(f"    {n} = 1\r\n" for n in chunk)
    return body + "End Sub\r\n"


def measure() -> dict[str, int]:
    names = probe_names()
    chunks = {f"hash{i}": _module(names[i:i + 30])
              for i in range(0, len(names), 30)}
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
    by_length: dict[int, dict[str, int]] = {}
    for name, value in table.items():
        by_length.setdefault(len(name), {})[name] = value
    exact = total = 0
    for length in sorted(by_length):
        samples = by_length[length]
        seed = fit_seed(samples)
        hits = sum(1 for n, v in samples.items() if identifier_hash(n) == v)
        exact += hits
        total += len(samples)
        flag = "" if hits == len(samples) else "   <- model drifts"
        print(f"  len {length:2}  seed {seed:#07x}  "
              f"{hits}/{len(samples)} exact{flag}")
    print(f"\n{exact}/{total} identifier hashes reproduced")
    return total - exact


if __name__ == "__main__":
    cache = Path(__file__).parent / CACHE
    if "--check" in sys.argv and cache.exists():
        data = json.loads(cache.read_text())
    else:
        data = measure()
        cache.write_text(json.dumps(data, indent=1, sort_keys=True))
    raise SystemExit(1 if check(data) > 30 else 0)
