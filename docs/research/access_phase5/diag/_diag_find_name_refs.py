"""Find all byte locations of name references in a renamed .accdb.

Compares the user's working ground-truth (fresh_renamed_m.accdb) with our
diag_E_rename_only.accdb to expose any place that still says 'M' in ours
where the ground truth says 'Renamed_M' (or has been restructured).
"""
from __future__ import annotations

from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / "demo" / "output" / "access_phase5f"
TEMPLATE = (
    Path(__file__).resolve().parents[1]
    / "tests" / "live_access_test" / "re_corpus" / "samples"
    / "040__sub_msgbox_hello.accdb"
)


def scan_name(data: bytes, name: str) -> list[tuple[int, str, bytes]]:
    """Return [(offset, encoding, context_bytes)] for occurrences.

    Looks for the name as latin-1 (MBCS) and as UTF-16LE, length-
    prefixed or naked. Returns small surrounding context for triage.
    """
    out: list[tuple[int, str, bytes]] = []
    nb_latin = name.encode("latin-1")
    nb_utf16 = name.encode("utf-16-le")
    # Plain MBCS occurrences.
    start = 0
    while True:
        i = data.find(nb_latin, start)
        if i < 0:
            break
        # Skip if preceded/followed by alphanumeric to avoid substring
        # collisions like 'Module' or 'Form'.
        prev_ok = i == 0 or not (chr(data[i - 1]).isalnum() or data[i - 1] == ord("_"))
        end = i + len(nb_latin)
        next_ok = end >= len(data) or not (chr(data[end]).isalnum() or data[end] == ord("_"))
        if prev_ok and next_ok:
            ctx_start = max(0, i - 8)
            ctx_end = min(len(data), end + 8)
            out.append((i, "latin1", data[ctx_start:ctx_end]))
        start = i + 1
    # UTF-16LE occurrences.
    start = 0
    while True:
        i = data.find(nb_utf16, start)
        if i < 0:
            break
        ctx_start = max(0, i - 8)
        ctx_end = min(len(data), i + len(nb_utf16) + 8)
        # Word boundary: previous and next u16 should not be alphanum.
        out.append((i, "utf16", data[ctx_start:ctx_end]))
        start = i + 1
    return out


def page_slot(offset: int) -> tuple[int, int]:
    PAGE = 4096
    return offset // PAGE, offset % PAGE


def main() -> None:
    template = TEMPLATE.read_bytes()
    diag = (OUT / "diag_E_rename_only.accdb").read_bytes()
    ground = (OUT / "fresh_renamed_m.accdb").read_bytes()

    print("=" * 70)
    print("TEMPLATE 040 (pre-rename, has 'M'):")
    print("=" * 70)
    for off, enc, ctx in scan_name(template, "M"):
        p, o = page_slot(off)
        print(f"  off=0x{off:08X} page={p} pageoff=0x{o:03X} enc={enc}  ctx={ctx!r}")

    print()
    print("=" * 70)
    print("GROUND TRUTH fresh_renamed_m (has 'Renamed_M', should have NO 'M' as a standalone):")
    print("=" * 70)
    refs_m = scan_name(ground, "M")
    print(f"  standalone 'M' occurrences: {len(refs_m)}")
    for off, enc, ctx in refs_m[:30]:
        p, o = page_slot(off)
        print(f"  off=0x{off:08X} page={p} pageoff=0x{o:03X} enc={enc}  ctx={ctx!r}")
    refs_r = scan_name(ground, "Renamed_M")
    print(f"  'Renamed_M' occurrences: {len(refs_r)}")
    for off, enc, ctx in refs_r:
        p, o = page_slot(off)
        print(f"  off=0x{off:08X} page={p} pageoff=0x{o:03X} enc={enc}  ctx={ctx!r}")

    print()
    print("=" * 70)
    print("OUR DIAG_E (any 'M' here is a missed rewrite):")
    print("=" * 70)
    refs_m = scan_name(diag, "M")
    print(f"  standalone 'M' occurrences: {len(refs_m)}")
    for off, enc, ctx in refs_m[:30]:
        p, o = page_slot(off)
        print(f"  off=0x{off:08X} page={p} pageoff=0x{o:03X} enc={enc}  ctx={ctx!r}")
    refs_r = scan_name(diag, "Renamed_M")
    print(f"  'Renamed_M' occurrences: {len(refs_r)}")
    for off, enc, ctx in refs_r:
        p, o = page_slot(off)
        print(f"  off=0x{off:08X} page={p} pageoff=0x{o:03X} enc={enc}  ctx={ctx!r}")


if __name__ == "__main__":
    main()
