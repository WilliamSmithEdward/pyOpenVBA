"""diag_D: call modify_module_cache with the IDENTICAL source text.
The blob bytes will differ (our compressor != Access's) but the
decompressed text matches exactly.

* D passes -> Access tolerates byte-different OVBA blobs as long as
  the decompressed text matches what p-code expects. The blocker is
  data-dependent (writing 'WORLD' breaks because p-code still says
  'hello'). Fix: also rewrite the p-code area, or invalidate it so
  Access recompiles.
* D fails -> Access byte-validates the OVBA blob (checksum/hash).
  Fix: produce byte-exact compressor output OR locate and patch
  the checksum.
"""
from __future__ import annotations
import shutil
from pathlib import Path

from pyopenvba.access import AccessFile

REPO = Path(__file__).resolve().parents[1]
SAMPLE = REPO / "tests" / "live_access_test" / "re_corpus" / "samples" / "040__sub_msgbox_hello.accdb"
OUT = REPO / "demo" / "output" / "access_phase5f"


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    dst = OUT / "diag_D_identical_source_rewrite.accdb"
    shutil.copy(SAMPLE, dst)

    db = AccessFile(dst)
    # Read the current source via the same path iter_vba_modules uses.
    src = db.read_vba_module("M")
    print(f'original source ({len(src)}B): {src!r}')

    # Set the module to literally the same source. Our compress() will
    # produce different bytes than Access's compress() did, but the
    # decompressed text will be identical.
    db.set_module("M", src)
    db.save()
    print(f'wrote identical source via modify_module_cache; saved to {dst.name}')

    # Verify read-back
    db2 = AccessFile(dst)
    src2 = db2.read_vba_module("M")
    print(f'read-back source ({len(src2)}B): {src2!r}')
    print(f'text round-trip match: {src == src2}')

    print()
    print('VERIFY: open diag_D_identical_source_rewrite.accdb -> click M.')
    print('  PASS = Sub A() / MsgBox "hello" shown, no error')
    print('  FAIL = error -> compressor byte-mismatch breaks Access')


if __name__ == "__main__":
    main()
