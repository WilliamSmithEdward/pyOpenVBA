"""Probe: dump the OVBA-cache row prefix for several samples to RE
the per-module wrapper layout (the bytes before the OVBA 0x01 magic)."""
import struct
from pathlib import Path

from pyopenvba.access import AccessFile
from pyopenvba.vba import decompress as _ovba_decompress

CORPUS = Path("tests/live_access_test/re_corpus/samples")
SAMPLES = sorted(CORPUS.glob("0*__*.accdb"))


def hexdump_prefix(name: str, blob: bytes, ovba_off: int) -> None:
    print(f"\n=== {name} (row_len={len(blob)}, ovba_off={ovba_off}) ===")
    # Print prefix in 16-byte rows.
    prefix = blob[:ovba_off]
    for i in range(0, len(prefix), 16):
        chunk = prefix[i : i + 16]
        hexs = " ".join(f"{b:02x}" for b in chunk)
        asci = "".join(chr(b) if 32 <= b < 127 else "." for b in chunk)
        print(f"  {i:04x}  {hexs:<48}  {asci}")
    # Decode candidate u32 fields.
    if ovba_off >= 32:
        print("  ---- u32 little-endian dwords ----")
        for i in range(0, min(ovba_off, 64), 4):
            (v,) = struct.unpack_from("<I", prefix, i)
            print(f"    [{i:04x}] = 0x{v:08x} ({v})")


for sample in SAMPLES[:6]:
    db = AccessFile(sample)
    found = False
    for page, slot, row in db._iter_lval_rows():  # pyright: ignore[reportPrivateUsage]
        sigs = db._scan_ovba_signatures(row)  # pyright: ignore[reportPrivateUsage]
        for off in sigs:
            try:
                decomp = _ovba_decompress(
                    bytes(row[off:]), stream_name=f"@({page},{slot})+{off}"
                )
            except Exception:
                continue
            if not decomp.startswith(b'Attribute VB_Name = "'):
                continue
            hexdump_prefix(
                f"{sample.name} page={page} slot={slot}",
                bytes(row),
                off,
            )
            print(f"  decompressed source size: {len(decomp)}")
            print(f"  compressed OVBA blob size: {len(bytes(row)) - off}")
            found = True
            break
        if found:
            break
