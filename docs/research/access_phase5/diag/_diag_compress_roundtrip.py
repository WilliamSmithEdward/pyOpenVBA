from pyopenvba.access import AccessFile
from pyopenvba.vba import compress, decompress

orig = AccessFile('tests/live_access_test/re_corpus/samples/040__sub_msgbox_hello.accdb')

for page, slot, row in orig._iter_lval_rows():
    sigs = orig._scan_ovba_signatures(row)
    for off in sigs:
        blob = bytes(row)[off:]
        try:
            d = decompress(blob)
        except Exception:
            continue
        if not d.startswith(b'Attribute VB_Name = "M"'):
            continue
        # Found the real one.
        recomp = compress(d)
        print(f'page={page} slot={slot} off={off}')
        print(f'  orig blob   len: {len(blob)}')
        print(f'  recomp blob len: {len(recomp)}')
        print(f'  decomp      len: {len(d)}')
        print(f'  byte-identical:  {blob == recomp}')
        print(f'  row len {len(row)} = off({off}) + blob_len({len(blob)}) -> tail bytes after blob = {len(row) - off - len(blob)}')
        n = min(len(blob), len(recomp))
        diffs = 0
        for i in range(n):
            if blob[i] != recomp[i]:
                if diffs < 6:
                    print(f'    diff @0x{i:04x}: orig={blob[i]:02x} ours={recomp[i]:02x}')
                diffs += 1
        print(f'  total mismatching bytes (first {n}): {diffs}')
        if len(blob) != len(recomp):
            print(f'  length differs by {len(recomp) - len(blob)}')
            print(f'  orig tail: {blob[-16:].hex(" ")}')
            print(f'  ours tail: {recomp[-16:].hex(" ")}')
        # Also show full hex of both for short blobs
        if len(blob) < 300:
            print('  orig blob hex:')
            for i in range(0, len(blob), 16):
                print(f'    {i:04x}  {blob[i:i+16].hex(" ")}')
            print('  ours blob hex:')
            for i in range(0, len(recomp), 16):
                print(f'    {i:04x}  {recomp[i:i+16].hex(" ")}')
        raise SystemExit
