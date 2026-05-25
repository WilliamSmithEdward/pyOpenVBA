"""Hex-dump the candidate p-code row (68, 1) across body-varying samples."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from pyopenvba.access import AccessFile  # noqa: E402

CORPUS = ROOT / "tests" / "live_access_test" / "re_corpus" / "samples"


def hex_dump(b: bytes, per_line: int = 32) -> str:
    out: list[str] = []
    for i in range(0, len(b), per_line):
        chunk = b[i : i + per_line]
        hexs = " ".join(f"{x:02x}" for x in chunk)
        ascii_s = "".join(chr(x) if 32 <= x < 127 else "." for x in chunk)
        out.append(f"  {i:04x}  {hexs:<{per_line*3}}  |{ascii_s}|")
    return "\n".join(out)


def main() -> None:
    samples = [
        ("030__sub_A_empty.accdb",      (68, 1)),
        ("040__sub_msgbox_hello.accdb", (68, 1)),
        ("041__sub_msgbox_world.accdb", (68, 1)),
        ("042__sub_msgbox_long.accdb",  (68, 1)),
        ("043__sub_msgbox_two.accdb",   (68, 1)),
        ("044__sub_dim_int.accdb",      (68, 1)),
        ("045__sub_dim_long.accdb",     (68, 1)),
        ("046__sub_dim_string.accdb",   (68, 1)),
        ("047__sub_let_int.accdb",      (68, 1)),
        ("048__sub_let_int_42.accdb",   (68, 1)),
        ("049__sub_comment_only.accdb", (68, 1)),
        ("050__sub_if_true.accdb",      (68, 1)),
        ("051__sub_for_1_to_3.accdb",   (68, 1)),
    ]
    for name, (p, s) in samples:
        path = CORPUS / name
        if not path.exists():
            continue
        bas = path.with_suffix(".bas")
        body = bas.read_text(encoding="utf-8-sig").rstrip() if bas.exists() else ""
        db = AccessFile(path)
        row = db._lval_row_bytes(p, s)  # pyright: ignore[reportPrivateUsage]
        print(f"=== {name}  ({p},{s}) len={len(row)} ===")
        print(f"BODY:\n  {body.replace(chr(10), chr(10)+'  ')}\n")
        print(hex_dump(bytes(row)))
        print()


if __name__ == "__main__":
    main()
