"""diag_C: writes the ORIGINAL OVBA blob bytes back to the same row
through _lval_write_row. Tests whether our write path corrupts the
page even when the bytes we write are identical to what's already there.

* C passes -> our write path is fine; Access requires byte-exact
  compressor parity (we'd need to refactor compress() to match Access's
  LZ77 token choices, or alternatively accept that any rewrite invalidates
  some external structure).
* C fails -> our write path itself perturbs the page (slot table /
  other rows), regardless of OVBA blob content.
"""
from __future__ import annotations

import shutil
from pathlib import Path

from pyopenvba.access import AccessFile
from pyopenvba.vba import decompress as ovba_decompress

REPO = Path(__file__).resolve().parents[1]
SAMPLE = REPO / "tests" / "live_access_test" / "re_corpus" / "samples" / "040__sub_msgbox_hello.accdb"
OUT = REPO / "demo" / "output" / "access_phase5f"


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    dst = OUT / "diag_C_byte_identical_rewrite.accdb"
    shutil.copy(SAMPLE, dst)

    db = AccessFile(dst)
    # Find the validated OVBA row for module M
    target: tuple[int, int, int, bytes] | None = None
    for page, slot, row in db._iter_lval_rows():
        sigs = db._scan_ovba_signatures(row)
        for off in sigs:
            blob = bytes(row)[off:]
            try:
                d = ovba_decompress(blob, stream_name="probe")
            except Exception:
                continue
            if d.startswith(b'Attribute VB_Name = "M"'):
                target = (page, slot, off, bytes(row))
                break
        if target is not None:
            break

    assert target is not None, "could not locate OVBA row for M"
    page, slot, off, row = target
    print(f'target: page={page} slot={slot} off={off} row_len={len(row)}')

    orig_blob = row[off:]
    # Rebuild row with same bytes (no change). Round-trip should be no-op.
    new_row = row[:off] + orig_blob
    assert new_row == row, "constructed row should equal original"
    db._lval_write_row(page, slot, new_row)
    db.save()
    print(f'wrote identical row back; saved to {dst.name}')

    # Verify file size and that read-back still finds M
    db2 = AccessFile(dst)
    for m in db2.iter_vba_modules():
        if m.name == "M":
            print(f'  read-back source: {m.source!r}')
            break

    print()
    print('VERIFY: open diag_C_byte_identical_rewrite.accdb -> open M.')
    print('  PASS = original Sub A()/MsgBox "hello" shown, no error.')
    print('  FAIL = error -> our _lval_write_row is corrupting the page')
    print('         even when writing the original bytes back.')


if __name__ == "__main__":
    main()
