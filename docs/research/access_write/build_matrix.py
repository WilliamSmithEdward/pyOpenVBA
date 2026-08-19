"""Compile a VBA source file into a fresh .accdb, using Access itself.

    python docs/research/access_write/build_matrix.py out.accdb source.bas

The point of building through Access is to get *its* p-code to diff
against, so the build has to go through real Access COM. ``pyvbaharness``
cannot stand in here: it removes the modules it loaded when its session
closes, which is right for running tests and wrong for making fixtures.

What this adds over calling ``build_matrix.ps1`` directly is that it
cannot hang. Access raises modal dialogs -- Save As, "already exists", a
compile error -- that block COM until the caller gives up, and a stranded
MSACCESS.EXE then holds a lock on the file and poisons the next run. So:
kill strays first, run under a hard timeout, and kill again on the way
out. A build that fails is fine; a build that hangs the terminal is not.

Use the harness in ``verify_execution.py`` for *running* code, which
watches for the same dialogs and reports ``modal-blocked`` rather than
blocking.

Dev-only: needs Windows and desktop Access.
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

SCRIPT = Path(__file__).with_name("build_matrix.ps1")
TIMEOUT_S = 180


def kill_stray_access() -> int:
    """Kill any MSACCESS.EXE left behind by an earlier run."""
    done = subprocess.run(
        ["powershell", "-NoProfile", "-Command",
         "$p = @(Get-Process MSACCESS -ErrorAction SilentlyContinue); "
         "$p | ForEach-Object { $_.Kill() }; $p.Count"],
        capture_output=True, text=True, timeout=60)
    try:
        return int(done.stdout.strip() or 0)
    except ValueError:
        return 0


def build(target: Path, source: Path, timeout: int = TIMEOUT_S) -> Path:
    """Build ``target`` from ``source``, compiled and saved by Access."""
    target = target.resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    target.unlink(missing_ok=True)
    stray = kill_stray_access()
    if stray:
        print(f"  note: killed {stray} stray Access process(es) first")
    try:
        done = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-File", str(SCRIPT),
             "-Target", str(target).replace("/", "\\"),
             "-Source", str(source.resolve())],
            capture_output=True, text=True, timeout=timeout)
    except subprocess.TimeoutExpired:
        kill_stray_access()
        raise SystemExit(
            f"Access did not finish building {source.name} within "
            f"{timeout}s, most likely sitting behind a modal dialog. The "
            "process has been killed; re-run, and if it repeats, build that "
            "source by hand with Access visible to see which dialog it is."
        ) from None
    if not target.exists():
        kill_stray_access()
        raise SystemExit(
            f"no database written to {target}\n{done.stdout}\n{done.stderr}")
    return target


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    out = build(Path(sys.argv[1]), Path(sys.argv[2]))
    print(f"built {out} ({out.stat().st_size} bytes)")
