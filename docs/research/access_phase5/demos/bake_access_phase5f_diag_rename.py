"""Isolation diagnostics for Phase 5g rename failure mode.

Bakes three files to triangulate where the crash originates:

  diag_E_rename_only.accdb
    Just M -> Renamed_M. No body change. If this crashes, the rename
    path itself produces a structurally inconsistent .accdb.

  diag_F_rename_then_set_same.accdb
    Rename, then set_module with the ORIGINAL body. Equivalent to
    rename + a no-op body rewrite. If diag_E passes and this fails,
    set_module after rename mis-targets the cache row.

  diag_G_set_then_rename.accdb
    set_module first (no-op body), then rename. Reverse ordering.

Open each in Access VBE and report PASS / soft-error / crash.
"""

from __future__ import annotations

import shutil
from pathlib import Path

from pyopenvba.access import AccessFile

REPO = Path(__file__).resolve().parents[1]
CORPUS = REPO / "tests" / "live_access_test" / "re_corpus" / "samples"
OUT = REPO / "demo" / "output" / "access_phase5f"
SAMPLE_STD = CORPUS / "040__sub_msgbox_hello.accdb"


def _copy(name: str) -> Path:
    dst = OUT / name
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(SAMPLE_STD, dst)
    return dst


def bake_E_rename_only() -> Path:
    dst = _copy("diag_E_rename_only.accdb")
    db = AccessFile(dst)
    db.rename_module("M", "Renamed_M")
    db.save()
    return dst


def bake_F_rename_then_set_same() -> Path:
    dst = _copy("diag_F_rename_then_set_same.accdb")
    db = AccessFile(dst)
    original = db.get_module("M")
    db.rename_module("M", "Renamed_M")
    # Strip the Attribute VB_Name header so set_module preserves
    # the (newly-renamed) attribute preamble verbatim.
    if original.startswith('Attribute VB_Name = "'):
        nl = original.find("\r\n")
        body_only = original[nl + 2 :] if nl >= 0 else ""
    else:
        body_only = original
    db.set_module("Renamed_M", body_only)
    db.save()
    return dst


def bake_G_set_then_rename() -> Path:
    dst = _copy("diag_G_set_then_rename.accdb")
    db = AccessFile(dst)
    original = db.get_module("M")
    if original.startswith('Attribute VB_Name = "'):
        nl = original.find("\r\n")
        body_only = original[nl + 2 :] if nl >= 0 else ""
    else:
        body_only = original
    db.set_module("M", body_only)
    db.rename_module("M", "Renamed_M")
    db.save()
    return dst


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    paths = [
        bake_E_rename_only(),
        bake_F_rename_then_set_same(),
        bake_G_set_then_rename(),
    ]
    print()
    print("Baked rename-isolation diagnostics:")
    for p in paths:
        print(f"  {p}")
        db = AccessFile(p)
        names = db.vba_module_names()
        print(f"    modules visible to pyopenvba: {names}")
    print()
    print("Open each in Access VBE and report PASS / soft-error / crash.")


if __name__ == "__main__":
    main()
