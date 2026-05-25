"""Phase 5h p-code invalidation probes.

The user observed that any LVAL-row-resizing edit on the OVBA cache
breaks Access GUI display, while same-length in-place edits work.
That points at Access caching compiled p-code in separate rows that
Access reads instead of the OVBA cache. memory/repo notes confirm:
the OVBA cache is passive; the authoritative source is rU@ p-code +
CAFE module-stream rows + plaintext B9/E3 records.

Hypothesis: if we TOMBSTONE the rU@ row and/or the CAFE module-stream
row, Access has no compiled p-code to read and might fall back to
recompiling from the OVBA cache on open.

Three diagnostics, each a copy of sample 040 with M intact:

  diag_L_tombstone_rU.accdb
    Tombstone the module-active rU@ p-code row. Keep CAFE + OVBA
    cache intact. If Access falls back to CAFE, project should
    still display M with original body.

  diag_M_tombstone_CAFE.accdb
    Tombstone the CAFE module-stream rows. Keep rU@ + OVBA cache
    intact. If Access falls back to rU@, project should still
    display M.

  diag_N_100line_plus_invalidate.accdb
    100-line rewrite of M's OVBA cache + tombstone BOTH rU@ AND
    CAFE module-stream rows. If Access falls back to OVBA cache,
    we'll see the 100-line body. This is the proof-of-recompile.
"""

from __future__ import annotations

import shutil
from pathlib import Path

from pyopenvba.access import AccessFile

# Re-use the 100-line builder from the previous diagnostic so the
# write payloads are identical.
import sys
sys.path.insert(0, str(Path(__file__).parent))
from bake_access_phase5h_diag_K_100line import _build_100_line_body

REPO = Path(__file__).resolve().parents[1]
CORPUS = REPO / "tests" / "live_access_test" / "re_corpus" / "samples"
OUT = REPO / "demo" / "output" / "access_phase5f"
SAMPLE_STD = CORPUS / "040__sub_msgbox_hello.accdb"


def _copy(name: str) -> Path:
    dst = OUT / name
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(SAMPLE_STD, dst)
    return dst


def _tombstone_active_pcode(db: AccessFile) -> tuple[int, int]:
    """Tombstone the rU@-prefixed module-active p-code row."""
    s = db.read_module_pcode_stream()
    db._lval_tombstone_slot(s.page, s.slot)
    return s.page, s.slot


def _tombstone_all_pcode(db: AccessFile) -> list[tuple[int, int]]:
    """Tombstone every rU@ row regardless of active/stub status."""
    out: list[tuple[int, int]] = []
    for s in db.iter_pcode_streams():
        db._lval_tombstone_slot(s.page, s.slot)
        out.append((s.page, s.slot))
    return out


def _tombstone_cafe_streams(db: AccessFile) -> list[tuple[int, int]]:
    """Tombstone every CAFE module-stream row."""
    out: list[tuple[int, int]] = []
    for ms in db.find_module_streams():
        db._lval_tombstone_slot(ms.page, ms.slot)
        out.append((ms.page, ms.slot))
    return out


def bake_L_tombstone_rU() -> Path:
    dst = _copy("diag_L_tombstone_rU.accdb")
    db = AccessFile(dst)
    killed = _tombstone_all_pcode(db)
    db.save()
    print(f"  diag_L tombstoned rU@ rows: {killed}")
    return dst


def bake_M_tombstone_CAFE() -> Path:
    dst = _copy("diag_M_tombstone_CAFE.accdb")
    db = AccessFile(dst)
    killed = _tombstone_cafe_streams(db)
    db.save()
    print(f"  diag_M tombstoned CAFE rows: {killed}")
    return dst


def bake_N_100line_plus_invalidate() -> Path:
    dst = _copy("diag_N_100line_plus_invalidate.accdb")
    db = AccessFile(dst)
    body = _build_100_line_body()
    db.set_module("M", body)
    killed_pcode = _tombstone_all_pcode(db)
    killed_cafe = _tombstone_cafe_streams(db)
    db.save()
    print(f"  diag_N wrote 100-line body to M")
    print(f"  diag_N tombstoned rU@ rows: {killed_pcode}")
    print(f"  diag_N tombstoned CAFE rows: {killed_cafe}")
    return dst


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    print("Baking p-code invalidation diagnostics ...")
    paths = [
        bake_L_tombstone_rU(),
        bake_M_tombstone_CAFE(),
        bake_N_100line_plus_invalidate(),
    ]
    print()
    for p in paths:
        print(f"  {p}")
        db = AccessFile(p)
        names = db.vba_module_names()
        print(f"    modules visible to pyopenvba: {names}")
    print()
    print("OPEN EACH IN ACCESS VBE AND REPORT RESULT:")
    print()
    print("  diag_L_tombstone_rU.accdb")
    print("    Expected if Access uses CAFE: M displays original body.")
    print("    If Access uses rU@: error / crash.")
    print()
    print("  diag_M_tombstone_CAFE.accdb")
    print("    Expected if Access uses rU@: M displays original body.")
    print("    If Access uses CAFE: error / crash.")
    print()
    print("  diag_N_100line_plus_invalidate.accdb")
    print("    Best case: M shows the 100-line body (Access recompiled")
    print("    from OVBA cache because both p-code sources are gone).")
    print("    Else: error / crash / empty M.")


if __name__ == "__main__":
    main()
