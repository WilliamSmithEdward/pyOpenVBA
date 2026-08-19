"""Semantic round-trip: p-code -> source -> p-code.

The text round-trip in :mod:`roundtrip` proves the decompiler reproduces
sources it was given. It cannot say anything about a module whose source
is unknown -- which is every real-world module.

This gate closes that hole. It takes a *compiled* module, decompiles it,
recompiles the result with Excel, and compares the two instruction
streams. If they agree, the reconstructed source means the same thing as
the original, whatever it looked like. A mismatch is a real semantic
defect: an expression rebuilt with the wrong precedence, a branch
inverted, an argument dropped.

Raw bytes are deliberately *not* compared. Identifier numbering, offsets
and slot assignment legitimately shift when a project is rebuilt, so the
comparison is over the opcode sequence and the shape of each operand --
the part that determines behaviour.

Line boundaries are not compared either, because they are formatting:
``Case 0: x = f()`` and the same two statements on separate lines compile
to the same opcodes in a different line layout. The stream is therefore
flattened before comparison, and a module whose opcodes match but whose
line structure differs is reported as equivalent-but-reflowed.

    python docs/research/pcode/semantic_roundtrip.py <scratch-dir> [files...]

Dev-only: needs Windows, desktop Excel and `pyvbaharness`.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from pcode_source import decompile

from pyopenvba import ExcelFile, PowerPointFile, WordFile
from pyopenvba.cfb import CFB
from pyopenvba.vba import VBAModuleKind
from pyopenvba.vba_pcode import disassemble_module_stream

HOSTS = {
    ".xlsm": ExcelFile, ".xlsb": ExcelFile, ".xls": ExcelFile,
    ".xlam": ExcelFile, ".docm": WordFile, ".dotm": WordFile,
    ".pptm": PowerPointFile,
}
MODULE = "M"

# Opcodes whose presence is an artifact of how a project was built rather
# than of what the code does, so they are ignored when comparing.
_NOISE = {"BoS", "BoSImplicit", "BoL", "Context", "EndContext", "LineCont"}


def signature(stream: bytes) -> tuple[list[str], int]:
    """The behaviour-carrying shape of a module's p-code.

    Returns the flattened opcode stream -- each mnemonic tagged with its
    operand kinds, but not their values, which renumber on a rebuild --
    plus the number of non-empty source lines, reported separately so a
    pure reflow is distinguishable from a semantic change.
    """
    flat: list[str] = []
    lines = 0
    for line in disassemble_module_stream(stream, is_64bit=True).lines:
        items: list[str] = []
        for ins in line.instructions:
            if ins.mnemonic in _NOISE:
                continue
            kinds = "/".join(kind for kind, _ in ins.operands)
            items.append(f"{ins.mnemonic}:{kinds}" if kinds else ins.mnemonic)
        if items:
            lines += 1
            flat.extend(items)
    return flat, lines


def _recompile(source: str, path: Path, session) -> bytes:
    body = f'Attribute VB_Name = "{MODULE}"\r\n' + source.replace("\n", "\r\n")
    if not body.endswith("\r\n"):
        body += "\r\n"
    if path.exists():
        path.unlink()
    with ExcelFile.create_new(path) as workbook:
        project = workbook.vba_project()
        if MODULE in [m.name for m in project.modules]:
            workbook.set_module(MODULE, body)
        else:
            project.add_module(MODULE, body, kind=VBAModuleKind.standard)
        workbook.save()
    session.open_document(path, read_only=False)
    session.compile_project(watch_seconds=25)
    session.save_as(path)
    with ExcelFile(path) as workbook:
        return bytes(workbook.vba_project().get_module(MODULE).prefix_bytes)


def run(scratch: Path, targets: list[Path]) -> int:
    from pyvbaharness import ExcelSession

    jobs: list[tuple[str, bytes, str]] = []
    for path in targets:
        host = HOSTS.get(path.suffix.lower())
        if host is None:
            continue
        try:
            with host(path) as document:
                project = document.vba_project()
                cfb = CFB.from_bytes(document.vba_project_bytes())
                table = cfb.get_stream_in_storage("VBA", "_VBA_PROJECT")
                for module in project.modules:
                    stream = bytes(module.prefix_bytes)
                    if not stream:
                        continue
                    text = decompile(stream, table)
                    if "[unmapped" in text or "<proc>" in text:
                        continue          # nothing to prove; sweep covers it
                    jobs.append((f"{path.name}:{module.name}", stream, text))
        except Exception as error:
            print(f"  skip {path.name}: {type(error).__name__}: {error}")

    if not jobs:
        print("no decompilable modules found")
        return 0

    failures = 0
    session = ExcelSession()
    try:
        for i, (label, original, text) in enumerate(jobs):
            try:
                rebuilt = _recompile(text, scratch / f"sem_{i}.xlsm", session)
            except Exception as error:
                failures += 1
                print(f"  RECOMPILE-FAILED {label}: "
                      f"{type(error).__name__}: {error}")
                continue
            before, before_lines = signature(original)
            after, after_lines = signature(rebuilt)
            if before == after:
                note = ("" if before_lines == after_lines
                        else f", reflowed {before_lines} -> {after_lines} lines")
                print(f"  equivalent  {label}  "
                      f"({len(before)} opcodes{note})")
                continue
            failures += 1
            print(f"  DIVERGED    {label}")
            for n, (a, b) in enumerate(zip(before, after, strict=False)):
                if a != b:
                    lo = max(0, n - 3)
                    print(f"      first difference at opcode {n}")
                    print(f"        original ...{before[lo:n + 4]}")
                    print(f"        rebuilt  ...{after[lo:n + 4]}")
                    break
            if len(before) != len(after):
                print(f"      opcode count {len(before)} -> {len(after)}")
    finally:
        session.close()
    print(f"\n{len(jobs) - failures}/{len(jobs)} semantically equivalent")
    return failures


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise SystemExit("usage: semantic_roundtrip.py <scratch-dir> [files...]")
    target_dir = Path(sys.argv[1])
    target_dir.mkdir(parents=True, exist_ok=True)
    given = [Path(a) for a in sys.argv[2:]]
    if not given:
        given = [p for p in Path("demo").rglob("*") if p.suffix.lower() in HOSTS]
    raise SystemExit(1 if run(target_dir, given) else 0)
