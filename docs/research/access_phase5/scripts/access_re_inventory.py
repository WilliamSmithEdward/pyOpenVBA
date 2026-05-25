"""Phase 4 — full LVAL row inventory across the RE corpus.

Classifies every non-tombstone LVAL row in every corpus sample to
isolate rows that are NOT:

    * MS-OVBA dir-stream catalog (Phase 3, solved)
    * MS-OVBA per-module source cache ("Attribute VB_Name = ...")
    * PROJECT plaintext (line-based ID="{...}")

The leftover rows are candidates for the authoritative p-code /
interned-string store. We tabulate by sample, then dump the
unique-content rows for diffing.
"""
from __future__ import annotations

import sys
from pathlib import Path
from collections import Counter

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from pyopenvba.access import AccessFile  # noqa: E402
from pyopenvba.vba import decompress as _ovba_decompress  # noqa: E402


CORPUS = ROOT / "tests" / "live_access_test" / "re_corpus" / "samples"


_DIR_MAGIC = b"\x01\x00\x04\x00\x00\x00"


def classify_row(row: bytes) -> tuple[str, bytes | None]:
    """Return (label, decompressed_or_None)."""
    if not row:
        return "empty", None
    # OVBA decompressible?
    if row[0] == 0x01 and len(row) >= 3:
        hdr = int.from_bytes(row[1:3], "little")
        if ((hdr >> 12) & 0x7) == 0b011:
            try:
                raw = _ovba_decompress(bytes(row), stream_name="probe")
            except Exception:
                raw = None
            if raw is not None:
                if raw.startswith(_DIR_MAGIC):
                    return "catalog_dir", raw
                if raw.startswith(b"Attribute VB_Name = "):
                    return "ovba_module", raw
                return "compressed_other", raw
    # PROJECT plaintext?
    if row[:5] == b'ID="{' or b'ID="{' in row[:16]:
        return "project_plaintext", row
    # Embedded ID="{ near start?
    if b'ID="{' in row[:64]:
        return "project_plaintext_offset", row
    # E3 source-row index?
    if b"\xE3\x00\x00\x00" in row[:32]:
        return "source_index_e3", row
    return "binary_other", row


def inventory(path: Path) -> dict[tuple[int, int], tuple[str, int, int]]:
    """Return {(page, slot): (label, row_len, raw_len_or_0)}."""
    db = AccessFile(path)
    out: dict[tuple[int, int], tuple[str, int, int]] = {}
    for page, slot, row in db._iter_lval_rows():  # pyright: ignore[reportPrivateUsage]
        label, raw = classify_row(row)
        out[(page, slot)] = (label, len(row), len(raw) if raw else 0)
    return out


def main() -> None:
    samples = sorted(CORPUS.glob("*.accdb"))
    if not samples:
        print(f"no corpus at {CORPUS}", file=sys.stderr)
        return

    # Header: per-sample row inventory
    print("=" * 80)
    print("ROW INVENTORY")
    print("=" * 80)
    for s in samples:
        inv = inventory(s)
        labels = Counter(v[0] for v in inv.values())
        print(f"\n{s.name}  ({sum(labels.values())} rows)")
        for (page, slot), (label, rlen, raw_len) in sorted(inv.items()):
            extra = f" raw={raw_len}" if raw_len else ""
            print(f"  ({page:>3},{slot})  len={rlen:<5} {label}{extra}")
        print(f"  ---- {dict(labels)}")

    # Focus: binary_other / compressed_other / source_index_e3 rows ARE
    # the candidate authoritative-store rows.
    print()
    print("=" * 80)
    print("UNKNOWN ROW LENGTHS PER SAMPLE (candidates for p-code store)")
    print("=" * 80)
    header_samples = [s.name for s in samples]
    # Per (page,slot) length matrix across samples for unknown rows
    keys: set[tuple[int, int]] = set()
    per_sample: dict[str, dict[tuple[int, int], tuple[str, int]]] = {}
    for s in samples:
        inv = inventory(s)
        per_sample[s.name] = {}
        for k, v in inv.items():
            if v[0] in ("binary_other", "source_index_e3", "compressed_other"):
                per_sample[s.name][k] = (v[0], v[1])
                keys.add(k)
    for k in sorted(keys):
        print(f"\nrow {k}:")
        for name in header_samples:
            entry = per_sample.get(name, {}).get(k)
            if entry is None:
                print(f"  {name:<48}  --")
            else:
                print(f"  {name:<48}  {entry[0]:<22} len={entry[1]}")


if __name__ == "__main__":
    main()
