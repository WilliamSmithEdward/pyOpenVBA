"""Gate: a rewritten module must actually run, not merely parse.

Every other gate here is static. ``verify_compiler`` proves our p-code
equals Microsoft's, ``verify_identity`` proves a module rebuilds byte for
byte -- and both passed for months on databases Access silently refused
to run, because Access executes its ``__SRP_*`` cache and neither gate
forced a recompile. This one closes that hole: it rewrites a procedure,
drops the cache, and asks Access for the answer.

    python docs/research/access_write/verify_execution.py <scratch dir>

Dev-only: needs Windows, desktop Access and ``pyvbaharness``. Never drive
Access with bare ``Application.Eval`` here -- a VBA compile-error dialog
is modal and blocks until the caller times out. The harness reports
``modal-blocked`` instead, which is how a failing case stays debuggable.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from rewrite_module import rewrite

# Each case rewrites the one procedure of a freshly built database and
# states what Access must then return. The line counts differ on purpose:
# a same-count body, a grown one, and a body built from nothing.
CASES: list[tuple[str, list[str], object]] = [
    ("same-count", ["Probe = 777"], 777),
    ("grown", ["n = 6", "Probe = n * 7"], 42),
    ("from-empty", [
        'Set d = CreateObject("Scripting.Dictionary")',
        'd.Add "a", 5',
        'd.Add "b", 7',
        'n = d.Item("a") + d.Item("b")',
        "n = n ^ 2",
        "Select Case n",
        "Case 144",
        "n = n + 1",
        "Case Else",
        "n = 0",
        "End Select",
        "Probe = n",
    ], 145),
    # Declarations: the record, its hash bucket and the arena all have to
    # be right, and the declared type has to actually bind -- `n` being a
    # Long is what makes 3.7 come back as 4.
    ("declares", ["Dim n As Long", "Dim r As Double",
                  "n = 42", "r = 0.5", "Probe = n * r"], 21.0),
    ("typed", ["Dim n As Long", "n = 3.7", "Probe = n"], 4),
]

SOURCE = (
    'Attribute VB_Name = "Runner"\r\n\r\n'
    "Public Function Probe() As Variant\r\n"
    "{body}"
    "End Function\r\n"
)


def build(scratch: Path, name: str, body: str) -> Path:
    """Compile a starting database with Access, so the p-code is real."""
    from build_matrix import build as build_accdb

    bas = scratch / f"{name}.bas"
    bas.write_bytes(SOURCE.format(body=body).encode("latin-1"))
    return build_accdb(scratch / f"{name}.accdb", bas)


def evaluate(accdb: Path, expression: str = "Probe()") -> tuple[str, object]:
    import pyvbaharness

    session = pyvbaharness.AccessSession()
    try:
        session.open_document(str(accdb.resolve()), read_only=False)
        try:
            return ("ok", session.eval(expression, timeout=60.0))
        except Exception as error:
            return (type(error).__name__, str(error).splitlines()[0][:120])
    finally:
        try:
            session.close()
        except Exception:
            pass


def main(argv: list[str]) -> int:
    scratch = Path(argv[0]) if argv else Path.cwd() / "_exec_gate"
    scratch.mkdir(parents=True, exist_ok=True)
    starting = build(scratch, "exec_base", "    Probe = 0\r\n")
    empty = build(scratch, "exec_empty", "")
    failures = 0
    for name, statements, expected in CASES:
        source = empty if name == "from-empty" else starting
        out = scratch / f"exec_{name}.accdb"
        out.unlink(missing_ok=True)
        try:
            rewrite(source, out, statements, None, "Probe")
        except SystemExit as refusal:
            print(f"  {name:<12} REFUSED: {refusal}")
            failures += 1
            continue
        kind, value = evaluate(out)
        ok = kind == "ok" and value == expected
        failures += not ok
        print(f"  {name:<12} {'ok  ' if ok else 'FAIL'} "
              f"expected {expected!r}, got {value!r}"
              f"{'' if kind == 'ok' else f' [{kind}]'}")
    print(f"\n{len(CASES) - failures}/{len(CASES)} cases execute correctly")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
