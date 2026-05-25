"""Hunt for the string-literal interning table and other companion
rows referenced by the module-active p-code stream.

Strategy: for samples 040 (MsgBox "hello") and 041 (MsgBox "world"),
list every LVAL row in the database and look for ones whose payload
*contains* the literal text (ASCII or UTF-16-LE). Companion rows
must vary in content with the literal but be stable in location.

Run from repo root::

    python scripts/access_re_find_string_intern.py
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from pyopenvba.access import AccessFile  # noqa: E402

CORPUS = ROOT / "tests" / "live_access_test" / "re_corpus" / "samples"


def find_rows_containing(path: Path, needle: bytes) -> list[tuple[int, int, int, int]]:
    db = AccessFile(path)
    hits: list[tuple[int, int, int, int]] = []
    for page, slot, row in db._iter_lval_rows():  # pyright: ignore[reportPrivateUsage]
        idx = bytes(row).find(needle)
        if idx != -1:
            hits.append((page, slot, len(row), idx))
    return hits


def hex_dump(b: bytes, *, per_line: int = 32) -> str:
    out: list[str] = []
    for i in range(0, len(b), per_line):
        chunk = b[i : i + per_line]
        hexs = " ".join(f"{x:02x}" for x in chunk)
        ascii_s = "".join(chr(x) if 32 <= x < 127 else "." for x in chunk)
        out.append(f"  {i:04x}  {hexs:<{per_line*3}}  |{ascii_s}|")
    return "\n".join(out)


def main() -> None:
    cases = [
        ("040__sub_msgbox_hello.accdb", "hello"),
        ("041__sub_msgbox_world.accdb", "world"),
        ("042__sub_msgbox_long.accdb", "a much longer literal that is clearly not stored inline"),
        ("049__sub_comment_only.accdb", "a comment"),
    ]
    for fname, text in cases:
        path = CORPUS / fname
        if not path.exists():
            continue
        print("=" * 80)
        print(f"  {fname}   needle={text!r}")
        for encoding in ("ascii", "utf-16-le"):
            needle = text.encode(encoding)
            hits = find_rows_containing(path, needle)
            print(f"  [{encoding}] hits: {hits}")
            # Dump the first non-(68,*) hit (p-code rows are at page 68)
            db = AccessFile(path)
            for page, slot, _length, _idx in hits:
                if page == 68:
                    continue
                row = bytes(db._lval_row_bytes(page, slot))  # pyright: ignore[reportPrivateUsage]
                print(f"  -- page={page} slot={slot} len={len(row)}")
                print(hex_dump(row))
                print()
                break
        print()


if __name__ == "__main__":
    main()
