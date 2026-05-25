"""Phase 4 — content-based row fingerprinting across the RE corpus.

The slot allocation shuffles dramatically when code is added (rows move
between pages and slots). So we classify each row by CONTENT signatures,
then track each content class across samples to find the authoritative
p-code store.

Known content fingerprints:
* catalog_dir          : OVBA-decompresses to PROJECTSYSKIND header
* ovba_cache           : OVBA-decompresses to "Attribute VB_Name ="
* project_plaintext    : Contains `ID="{` (line-based PROJECT)
* references_libid     : Contains `*\\G{` ASCII (MS-OVBA LibID format)
* has_e3_markers       : Contains 0xE3 0x00 0x00 0x00 (comment-row index)
* has_b9_markers       : Contains 0xB9 0x00 0x?? followed by ASCII text
* has_module_name_utf16: Contains the module name UTF-16-LE
* has_proc_name        : Contains the procedure name (Sub <X>)
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from pyopenvba.access import AccessFile  # noqa: E402
from pyopenvba.vba import decompress as _ovba_decompress  # noqa: E402

CORPUS = ROOT / "tests" / "live_access_test" / "re_corpus" / "samples"

_DIR_MAGIC = b"\x01\x00\x04\x00\x00\x00"


def fingerprint(row: bytes, module_name: str, proc_name: str | None) -> list[str]:
    tags: list[str] = []
    # OVBA prefix?
    if row[:1] == b"\x01" and len(row) >= 3:
        hdr = int.from_bytes(row[1:3], "little")
        if ((hdr >> 12) & 0x7) == 0b011:
            try:
                raw = _ovba_decompress(bytes(row), stream_name="probe")
            except Exception:
                pass
            else:
                if raw.startswith(_DIR_MAGIC):
                    tags.append("catalog_dir")
                elif raw.startswith(b"Attribute VB_Name = "):
                    tags.append("ovba_cache_root")
    # Embedded OVBA further in?
    if "ovba_cache_root" not in tags:
        for j in range(0, min(len(row), 2048) - 3):
            if row[j] == 0x01:
                hdr = int.from_bytes(row[j + 1 : j + 3], "little")
                if ((hdr >> 12) & 0x7) == 0b011:
                    try:
                        raw = _ovba_decompress(bytes(row[j:]), stream_name="probe")
                    except Exception:
                        continue
                    if raw.startswith(b"Attribute VB_Name = "):
                        tags.append(f"ovba_cache_wrapped@{j}")
                        break
                    if raw.startswith(_DIR_MAGIC):
                        tags.append(f"catalog_wrapped@{j}")
                        break
    if b'ID="{' in row[:128]:
        tags.append("project_plaintext")
    if b"*\\G{" in row:
        tags.append("references_libid")
    if b"\xE3\x00\x00\x00" in row:
        tags.append("has_e3_markers")
    # b9 string literal: 0xB9 0x00 <len_byte> <ascii payload>
    if b"\xB9\x00" in row:
        tags.append("has_b9_markers")
    # Module name as ANSI / UTF-16-LE
    if module_name and module_name.encode("ascii", errors="ignore") in row:
        tags.append("has_module_name_ansi")
    if module_name:
        u16 = module_name.encode("utf-16-le", errors="ignore")
        if u16 in row:
            tags.append("has_module_name_u16")
    if proc_name and proc_name.encode("ascii", errors="ignore") in row:
        tags.append(f"has_proc_name_ansi[{proc_name}]")
    # "Attribute VB_" plaintext (uncompressed OVBA cache stub)
    if b"Attribute VB_" in row:
        tags.append("has_attribute_plaintext")
    # "Option Compare" plaintext
    if b"Option Compare" in row:
        tags.append("has_option_compare")
    return tags


def module_and_proc(bas_path: Path) -> tuple[str, str | None]:
    """Extract Attribute VB_Name and first Sub/Function name from .bas."""
    text = bas_path.read_text(encoding="utf-8", errors="replace")
    name = ""
    proc: str | None = None
    for ln in text.splitlines():
        s = ln.strip()
        if s.startswith('Attribute VB_Name = "'):
            name = s.split('"', 2)[1]
        elif s.startswith(("Sub ", "Function ", "Public Sub ", "Private Sub ")):
            head = s.split("(", 1)[0]
            proc = head.split()[-1]
            break
    return name, proc


def inventory_sample(accdb: Path) -> list[tuple[tuple[int, int], int, list[str]]]:
    bas = accdb.with_suffix(".bas")
    mod_name, proc = module_and_proc(bas) if bas.exists() else ("", None)
    db = AccessFile(accdb)
    out: list[tuple[tuple[int, int], int, list[str]]] = []
    for page, slot, row in db._iter_lval_rows():  # pyright: ignore[reportPrivateUsage]
        if page == 42:
            continue
        out.append(((page, slot), len(row), fingerprint(bytes(row), mod_name, proc)))
    return out


def main() -> None:
    samples = sorted(CORPUS.glob("*.accdb"))
    print(f"# Corpus: {len(samples)} samples\n")
    for s in samples:
        inv = inventory_sample(s)
        print(f"## {s.name}")
        bas = s.with_suffix(".bas")
        mod_name, proc = module_and_proc(bas) if bas.exists() else ("", None)
        print(f"   module={mod_name!r}  proc={proc!r}")
        for k, rlen, tags in inv:
            tag_str = ", ".join(tags) if tags else "<no-match>"
            print(f"   {k}  len={rlen:<5}  {tag_str}")
        print()


if __name__ == "__main__":
    main()
