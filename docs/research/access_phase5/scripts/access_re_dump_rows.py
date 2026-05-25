"""Dump content of unknown rows for the body-varying samples.

Focus: find the row that GROWS with code body and contains the
authoritative p-code stream.
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from pyopenvba.access import AccessFile  # noqa: E402

CORPUS = ROOT / "tests" / "live_access_test" / "re_corpus" / "samples"


def proc_name(bas: Path) -> str:
    if not bas.exists():
        return ""
    for ln in bas.read_text(encoding="utf-8", errors="replace").splitlines():
        s = ln.strip()
        for kw in ("Sub ", "Function ", "Public Sub ", "Private Sub "):
            if s.startswith(kw):
                head = s.split("(", 1)[0]
                return head.split()[-1]
    return ""


def collect_rows(path: Path) -> list[tuple[int, int, bytes]]:
    db = AccessFile(path)
    rows: list[tuple[int, int, bytes]] = []
    for page, slot, row in db._iter_lval_rows():  # pyright: ignore[reportPrivateUsage]
        if page == 42:
            continue
        rows.append((page, slot, bytes(row)))
    return rows


def hex_dump(b: bytes, prefix: str = "  ", per_line: int = 32) -> str:
    out: list[str] = []
    for i in range(0, len(b), per_line):
        chunk = b[i : i + per_line]
        hexs = " ".join(f"{x:02x}" for x in chunk)
        ascii_s = "".join(
            chr(x) if 32 <= x < 127 else "." for x in chunk
        )
        out.append(f"{prefix}{i:04x}  {hexs:<{per_line*3}}  |{ascii_s}|")
    return "\n".join(out)


def main() -> None:
    samples = [
        "030__sub_A_empty.accdb",
        "040__sub_msgbox_hello.accdb",
        "041__sub_msgbox_world.accdb",
        "042__sub_msgbox_long.accdb",
        "043__sub_msgbox_two.accdb",
        "044__sub_dim_int.accdb",
        "046__sub_dim_string.accdb",
        "047__sub_let_int.accdb",
        "048__sub_let_int_42.accdb",
        "049__sub_comment_only.accdb",
        "050__sub_if_true.accdb",
        "051__sub_for_1_to_3.accdb",
    ]
    bodies = {}
    for name in samples:
        bas = (CORPUS / name).with_suffix(".bas")
        bodies[name] = bas.read_text(encoding="utf-8") if bas.exists() else ""

    print("# Sample bodies\n")
    for name in samples:
        print(f"## {name}")
        print(f"```\n{bodies[name]}\n```\n")

    print("=" * 90)
    print("All non-page-42 rows by sample")
    print("=" * 90)
    for name in samples:
        path = CORPUS / name
        if not path.exists():
            continue
        rows = collect_rows(path)
        # Get module/proc names
        proj = AccessFile(path).read_project_info()
        modname = proj.modules[0].name if proj.modules else ""
        proc = proc_name(path.with_suffix(".bas"))
        print(f"\n### {name}  module={modname!r}  proc={proc!r}")
        for page, slot, row in rows:
            # Identify by content
            tag = []
            if row[:1] == b"\x01":
                tag.append("ovba-prefix")
            if b"Attribute VB_" in row:
                tag.append("has-Attribute")
            if b'ID="{' in row[:64]:
                tag.append("PROJECT-plaintext")
            if proc and proc.encode() in row:
                tag.append(f"has-proc={proc}")
            if modname and modname.encode() in row:
                tag.append(f"has-mod={modname}")
            if b"hello" in row or b"world" in row:
                tag.append("has-string-literal")
            print(f"  ({page},{slot}) len={len(row)} {tag}")


if __name__ == "__main__":
    main()
