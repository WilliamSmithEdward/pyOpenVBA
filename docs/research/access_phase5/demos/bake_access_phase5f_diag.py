"""
Phase 5f diagnostic bake: produce two files that bracket the failure
mode of 01_simple_multiline_replace.accdb.

  diag_A_noop_save.accdb        copy of sample 040, opened + save()
                                with ZERO edits. Should be byte-stable.
  diag_B_same_length_replace.accdb
                                set_module to a body whose decompressed
                                payload is the same length as the
                                original (within +/-1 byte).

Open each in Access and report which (if any) throw 'Error accessing
file. Network connection may have been lost.' when you click on M.

* diag_A pass + diag_B pass -> Access tolerates our writes; the
  failure on 01 is a size-delta problem (we need to RE the length
  field that bounds the OVBA cache).
* diag_A pass + diag_B fail -> any source change breaks VBE display.
* diag_A fail -> save() alone corrupts something even without edits.
"""

from __future__ import annotations

import shutil
from pathlib import Path

from pyopenvba.access import AccessFile

REPO = Path(__file__).resolve().parents[1]
CORPUS = REPO / "tests" / "live_access_test" / "re_corpus" / "samples"
OUT = REPO / "demo" / "output" / "access_phase5f"
SAMPLE = CORPUS / "040__sub_msgbox_hello.accdb"


def bake_noop() -> Path:
    dst = OUT / "diag_A_noop_save.accdb"
    shutil.copy(SAMPLE, dst)
    db = AccessFile(dst)
    db.save()
    return dst


def bake_same_length() -> Path:
    """Original body decompressed is 90 bytes. Build a replacement
    body whose decompressed payload comes out to exactly 90 bytes
    (Attribute VB_Name = "M"\r\n = 25 bytes prefix; body = 65 bytes)."""
    dst = OUT / "diag_B_same_length_replace.accdb"
    shutil.copy(SAMPLE, dst)
    # Original body: 'Option Compare Database\r\n\r\nSub A()\r\n    MsgBox "hello"\r\nEnd Sub\r\n\r\n' = 65 bytes
    new_body = (
        'Option Compare Database\r\n'
        '\r\n'
        'Sub Z()\r\n'
        '    MsgBox "WORLD"\r\n'
        'End Sub\r\n'
        '\r\n'
    )
    assert len(new_body) == 67, f"body must be 67 bytes, got {len(new_body)}"
    db = AccessFile(dst)
    db.set_module("M", new_body)
    db.save()
    return dst


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    a = bake_noop()
    b = bake_same_length()

    # Read-back probe
    for label, path in [("A noop", a), ("B same-length", b)]:
        db = AccessFile(path)
        m = next(db.iter_vba_modules())
        decomp_len = len(m.attributes_text) + len(m.source)
        print(f'  {label}: {path.name}  decomp={decomp_len}B  source={m.source!r}')

    print()
    print("VERIFY IN ACCESS:")
    print("  1. Open diag_A_noop_save.accdb -> open module M -- should show")
    print("     the original Sub A() / MsgBox \"hello\" body, no error.")
    print("  2. Open diag_B_same_length_replace.accdb -> open module M --")
    print("     does it error, or does it show Sub Z() / MsgBox \"WORLD\"?")


if __name__ == "__main__":
    main()
