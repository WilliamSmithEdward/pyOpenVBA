"""Bake Phase 5h delete/add diagnostics for Access GUI verification.

If delete + add work in Access GUI, rename can be implemented as
decompose: capture source -> delete -> add(new_name, source). That
side-steps the rename-specific name-table desync that bricks
Renamed_M-style files.

Three diagnostics, all from sample 040 (single module ``M``):

  diag_H_delete_only.accdb
      delete_module("M") -> save. Access should open with no module M.

  diag_I_add_catalog_only.accdb
      add_module_catalog_entry("AddedZ") with NO body. Tests whether
      Access tolerates a dir-stream entry with no OVBA cache row.

  diag_J_delete_then_add_with_body.accdb
      Captures M's source + attributes, deletes M, appends an OVBA
      cache row for a fresh module ``ReM`` containing the captured
      source (VB_Name forced to ``ReM``), then registers ``ReM`` in
      the dir-stream + MSysObjects. This is the decompose-rename
      proof.
"""

from __future__ import annotations

import re
import shutil
from pathlib import Path

from pyopenvba.access import ACE_PAGE_SIZE, AccessError, AccessFile
from pyopenvba.vba import compress as ovba_compress

REPO = Path(__file__).resolve().parents[1]
CORPUS = REPO / "tests" / "live_access_test" / "re_corpus" / "samples"
OUT = REPO / "demo" / "output" / "access_phase5f"
SAMPLE_STD = CORPUS / "040__sub_msgbox_hello.accdb"


def _copy(name: str) -> Path:
    dst = OUT / name
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(SAMPLE_STD, dst)
    return dst


def _force_vb_name(text: str, new_name: str) -> str:
    """Rewrite Attribute VB_Name to bind ``new_name``."""
    pattern = re.compile(r'^Attribute VB_Name = "[^"]*"', re.MULTILINE)
    replacement = f'Attribute VB_Name = "{new_name}"'
    if pattern.search(text):
        return pattern.sub(replacement, text, count=1)
    return replacement + "\r\n" + text


def _append_ovba_cache_row(db: AccessFile, payload: bytes) -> tuple[int, int]:
    """Find any LVAL page with enough free space and append the OVBA
    cache row there. Returns the (page, slot)."""
    need = len(payload) + 2  # +2 for new slot table entry
    for page in db._iter_lval_pages():
        if db._lval_free_space(page) >= need:
            slot = db._lval_append_row(page, payload)
            return page, slot
    raise AccessError(
        f"_append_ovba_cache_row: no LVAL page has {need} bytes free"
    )


def bake_H_delete_only() -> Path:
    dst = _copy("diag_H_delete_only.accdb")
    db = AccessFile(dst)
    db.delete_module("M")
    db.save()
    return dst


def bake_I_add_catalog_only() -> Path:
    dst = _copy("diag_I_add_catalog_only.accdb")
    db = AccessFile(dst)
    db.add_module_catalog_entry("AddedZ")
    db.save()
    return dst


def bake_J_delete_then_add_with_body() -> Path:
    dst = _copy("diag_J_delete_then_add_with_body.accdb")
    db = AccessFile(dst)

    # 1) Capture M's full source (attributes + body).
    full_source = db.read_vba_module_with_attributes("M")
    # 2) Force VB_Name to the new module name.
    new_full = _force_vb_name(full_source, "ReM")

    # 3) Delete the old module entirely (dir-stream, OVBA cache,
    # MSysObjects).
    db.delete_module("M")

    # 4) Compress and append a fresh OVBA cache row for ReM.
    compressed = ovba_compress(new_full.encode("latin-1"))
    _page, _slot = _append_ovba_cache_row(db, compressed)

    # 5) Register ReM in the dir-stream catalog (this also adds the
    # MSysObjects parallel row).
    db.add_module_catalog_entry("ReM")

    db.save()
    return dst


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    paths = [
        bake_H_delete_only(),
        bake_I_add_catalog_only(),
        bake_J_delete_then_add_with_body(),
    ]
    print()
    print("Baked delete/add diagnostics:")
    for p in paths:
        print(f"  {p}")
        db = AccessFile(p)
        names = db.vba_module_names()
        print(f"    modules visible to pyopenvba: {names}")
        for n in names:
            try:
                body = db.get_module(n)
                preview = body.split("\r\n", 1)[0]
                print(f"      [{n}] first line: {preview!r}")
            except AccessError as e:
                print(f"      [{n}] get_module failed: {e}")
    print()
    print("OPEN EACH IN ACCESS VBE AND REPORT PASS / soft-error / crash /")
    print("'cannot read VBA project' dialog:")
    print("  diag_H_delete_only.accdb         -- expect project opens, no M")
    print("  diag_I_add_catalog_only.accdb    -- expect M + AddedZ shell or")
    print("                                      a soft error about AddedZ")
    print("  diag_J_delete_then_add_with_body -- expect ReM with M's body")


if __name__ == "__main__":
    _ = ACE_PAGE_SIZE  # silence unused-import warning if check tools complain
    main()
