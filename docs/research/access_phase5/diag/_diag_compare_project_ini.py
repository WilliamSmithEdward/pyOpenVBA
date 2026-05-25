"""Compare PROJECT INI between user's ground-truth fresh_renamed_m.accdb
and our diag_E_rename_only.accdb.
"""
from __future__ import annotations
from pathlib import Path

from pyopenvba.access import AccessFile as AccessDatabase

OUT = Path(__file__).parent / "output" / "access_phase5f"


def dump(label: str, path: Path) -> None:
    print(f"\n===== {label}: {path.name} =====")
    db = AccessDatabase(path)
    found = db._find_project_ini_row()
    if found is None:
        print("  no PROJECT INI row found")
        return
    page, slot, raw = found
    print(f"  page={page} slot={slot} len={len(raw)}")
    # Decode latin-1 for readability.
    text = raw.decode("latin-1", errors="replace")
    print("  --- BEGIN ---")
    for line in text.splitlines():
        print(f"   {line}")
    print("  --- END ---")
    # Also look for stale "M" references.
    for needle in (b"Module=M\r\n", b"\r\nM=", b"Renamed_M"):
        idxs: list[int] = []
        start = 0
        while True:
            i = raw.find(needle, start)
            if i < 0:
                break
            idxs.append(i)
            start = i + 1
        print(f"  needle {needle!r} occurrences in PROJECT INI: {idxs}")


def main() -> None:
    dump("GROUND TRUTH (user)", OUT / "fresh_renamed_m.accdb")
    dump("TEMPLATE (pre-rename)",
         Path(__file__).resolve().parents[1]
         / "tests" / "live_access_test" / "re_corpus" / "samples"
         / "040__sub_msgbox_hello.accdb")
    dump("OUR DIAG_E (Phase5h)", OUT / "diag_E_rename_only.accdb")


if __name__ == "__main__":
    main()
