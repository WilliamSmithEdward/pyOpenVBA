"""Side-by-side pairwise diff of module-active p-code rows.

Uses the deterministic 12-byte active-row discriminator
(``AccessFile.read_module_pcode_stream``). Prints aligned byte diffs
for each selected pair; useful for opcode reverse-engineering.

Run from repo root::

    python scripts/access_re_pcode_diff.py
"""
from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from pyopenvba.access import AccessFile  # noqa: E402

CORPUS = ROOT / "tests" / "live_access_test" / "re_corpus" / "samples"


@dataclass(frozen=True)
class Sample:
    sid: str
    name: str

    @property
    def path(self) -> Path:
        return CORPUS / f"{self.sid}__{self.name}.accdb"

    @property
    def bas(self) -> str:
        bas = CORPUS / f"{self.sid}__{self.name}.bas"
        return bas.read_text(encoding="utf-8-sig").rstrip() if bas.exists() else ""


SAMPLES: dict[str, Sample] = {
    s.sid: s
    for s in [
        Sample("010", "empty_StdModule_M"),
        Sample("020", "empty_ClassModule_C"),
        Sample("030", "sub_A_empty"),
        Sample("031", "sub_B_empty"),
        Sample("032", "sub_AB_empty"),
        Sample("040", "sub_msgbox_hello"),
        Sample("041", "sub_msgbox_world"),
        Sample("042", "sub_msgbox_long"),
        Sample("043", "sub_msgbox_two"),
        Sample("044", "sub_dim_int"),
        Sample("045", "sub_dim_long"),
        Sample("046", "sub_dim_string"),
        Sample("047", "sub_let_int"),
        Sample("048", "sub_let_int_42"),
        Sample("049", "sub_comment_only"),
        Sample("050", "sub_if_true"),
        Sample("051", "sub_for_1_to_3"),
    ]
}


PAIRS: list[tuple[str, str, str]] = [
    ("030", "044", "empty Sub vs Dim Integer"),
    ("044", "045", "Dim Integer vs Dim Long  (per-type token)"),
    ("044", "046", "Dim Integer vs Dim String"),
    ("030", "049", "empty Sub vs comment-only"),
    ("047", "048", "x = 1  vs  x = 42  (i16 immediate)"),
    ("040", "041", "MsgBox \"hello\" vs \"world\" (literal intern)"),
    ("040", "042", "MsgBox short vs long literal"),
    ("040", "043", "MsgBox one  vs  two stmts"),
    ("030", "050", "empty Sub vs If/Then"),
    ("030", "051", "empty Sub vs For 1 To 3"),
    ("010", "020", "Std Module M vs Class Module C  (module kind)"),
    ("010", "030", "empty Std Module vs empty Sub A"),
    ("030", "031", "Sub A() vs Sub B()  (proc-name token)"),
    ("030", "032", "Sub A() vs Sub AB() (2-char proc name)"),
]


def fetch(sid: str) -> bytes:
    db = AccessFile(SAMPLES[sid].path)
    return db.read_module_pcode_stream().raw


def hex_dump(b: bytes, *, per_line: int = 16) -> list[str]:
    out: list[str] = []
    for i in range(0, len(b), per_line):
        chunk = b[i : i + per_line]
        hexs = " ".join(f"{x:02x}" for x in chunk)
        ascii_s = "".join(chr(x) if 32 <= x < 127 else "." for x in chunk)
        out.append(f"  {i:04x}  {hexs:<{per_line*3}}  |{ascii_s}|")
    return out


def diff_offsets(a: bytes, b: bytes) -> list[int]:
    n = min(len(a), len(b))
    diffs = [i for i in range(n) if a[i] != b[i]]
    if len(a) != len(b):
        diffs.append(n)
    return diffs


def side_by_side(a: bytes, b: bytes, *, per_line: int = 16) -> str:
    diffs = set(diff_offsets(a, b))
    rows_a = hex_dump(a, per_line=per_line)
    rows_b = hex_dump(b, per_line=per_line)
    n = max(len(rows_a), len(rows_b))
    out: list[str] = []
    for i in range(n):
        la = rows_a[i] if i < len(rows_a) else ""
        lb = rows_b[i] if i < len(rows_b) else ""
        marker = ""
        if any(d // per_line == i for d in diffs):
            offs = sorted({d for d in diffs if d // per_line == i})
            marker = "  <-- diff @ " + ",".join(f"0x{d:02x}" for d in offs)
        out.append(f"{la:<74}  {lb}{marker}")
    return "\n".join(out)


def main() -> None:
    for left, right, desc in PAIRS:
        a = fetch(left)
        b = fetch(right)
        diffs = diff_offsets(a, b)
        print("=" * 110)
        print(f"  {left} <-> {right}    {desc}")
        print(f"  lens: {len(a)} vs {len(b)}   diff_bytes: {len(diffs)}   first_diff_offsets: {[hex(d) for d in diffs[:8]]}")
        print(f"  LEFT  src: {SAMPLES[left].bas!r}")
        print(f"  RIGHT src: {SAMPLES[right].bas!r}")
        print("-" * 110)
        print(side_by_side(a, b))
        print()


if __name__ == "__main__":
    main()
