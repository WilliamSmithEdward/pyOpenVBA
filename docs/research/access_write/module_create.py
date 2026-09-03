"""Create a module: a command line over the shipped writer.

This began as the research that worked out what a new module costs an
`.accdb`, and everything it found now lives in `pyopenvba.access._vba`
and `AccessDatabase.create_module`.  What is left here is a way to run it
from a shell while probing a database by hand; the README keeps the
measurements.

    python module_create.py SOURCE TARGET NAME [CODE] [module|class]

`CODE` uses `|` for a line break, so a shell can pass a whole module
without a heredoc.
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
if str(ROOT / "src") not in sys.path:
    sys.path.insert(0, str(ROOT / "src"))

from pyopenvba.access import AccessDatabase  # noqa: E402


def create(
    source: Path,
    target: Path,
    name: str,
    code: str = "Option Compare Database",
    kind: str = "module",
) -> str:
    db = AccessDatabase(source)
    module = db.create_module(name, code, kind=kind)
    db.save(target)
    return f"{module.kind} {module.name!r} in stream {module.stream_name!r}"


if __name__ == "__main__":
    if len(sys.argv) < 4:
        raise SystemExit("usage: module_create.py SOURCE TARGET NAME [CODE] [module|class]")
    body = sys.argv[4] if len(sys.argv) > 4 else "Option Compare Database"
    print(
        create(
            Path(sys.argv[1]),
            Path(sys.argv[2]),
            sys.argv[3],
            body.replace("|", chr(10)),
            sys.argv[5] if len(sys.argv) > 5 else "module",
        )
    )
