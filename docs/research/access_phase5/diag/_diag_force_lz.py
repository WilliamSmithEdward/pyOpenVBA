"""Test: bypass the literal-only short-circuit in compress() and see
if the resulting bytes match Access's encoding for sample 040 M."""
from __future__ import annotations
import struct

from pyopenvba.access import AccessFile
from pyopenvba.vba import decompress, _encode_lz  # type: ignore

orig = AccessFile('tests/live_access_test/re_corpus/samples/040__sub_msgbox_hello.accdb')

for page, slot, row in orig._iter_lval_rows():
    sigs = orig._scan_ovba_signatures(row)
    for off in sigs:
        blob = bytes(row)[off:]
        try:
            d = decompress(blob, stream_name='probe')
        except Exception:
            continue
        if not d.startswith(b'Attribute VB_Name = "M"'):
            continue

        # Re-encode using LZ path directly (bypass literal-only shortcut)
        encoded = _encode_lz(d)
        # Wrap in OVBA framing
        header = 0xB000 | (len(encoded) - 1)
        ours_lz = bytes([0x01]) + struct.pack('<H', header) + encoded

        print(f'orig blob len: {len(blob)}')
        print(f'ours LZ-forced len: {len(ours_lz)}')
        print(f'byte equal: {blob == ours_lz}')

        n = min(len(blob), len(ours_lz))
        diffs = 0
        for i in range(n):
            if blob[i] != ours_lz[i]:
                if diffs < 10:
                    print(f'  diff @0x{i:04x}: orig={blob[i]:02x} ours={ours_lz[i]:02x}')
                diffs += 1
        print(f'  total diffs in first {n} bytes: {diffs}')
        if len(blob) != len(ours_lz):
            print(f'  length differs by {len(ours_lz) - len(blob)}')
        raise SystemExit
