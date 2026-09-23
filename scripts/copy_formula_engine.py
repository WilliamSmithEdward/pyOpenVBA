"""Compare pyopenvba.formula._calc with pyOfficeEditor's formula engine, or copy it over.

The engine is pyOfficeEditor's ``pyofficeeditor.excel._calc``, as that
repository has it committed, with its imports pointed at pyOpenVBA: its
own modules keep their names, and what it took from the rest of
pyOfficeEditor comes along as collate.py and reference.py. cells.py,
host.py and __init__.py are pyOpenVBA's own and are left alone, as is
pyOfficeEditor's workbook driver, engine.py, which pyOpenVBA does not use.

    python scripts/copy_formula_engine.py [path to pyOfficeEditor]

prints how each copied module differs from its source; with --write it
copies the source over. The copies carry pyOpenVBA's own changes too, so
read the differences before writing, and merge where both have moved.
"""

from __future__ import annotations

import argparse
import difflib
import re
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TARGET = ROOT / "src" / "pyopenvba" / "formula" / "_calc"
DEFAULT_SOURCE = ROOT.parent / "pyOfficeEditor"

REWRITES: list[tuple[str, str]] = [
    (r"from pyofficeeditor\.excel import _collate\b", "from pyopenvba.formula._calc import collate as _collate"),
    (r"from pyofficeeditor\.excel\._values import CellError\b", "from pyopenvba.formula._calc.cells import CellError"),
    (r"from pyofficeeditor\.exceptions import UnsupportedFormulaError\b",
     "from pyopenvba.formula._calc.cells import UnsupportedFormulaError"),
    (r"from pyofficeeditor\.excel\._formulas import (quote_sheet_name|translate_formula)\b",
     r"from pyopenvba.formula._calc.host import \1"),
    (r"from pyofficeeditor\.excel\._numfmt import format_value\b", "from pyopenvba.formula._calc.host import format_value"),
    (r"pyofficeeditor\.excel\._reference\b", "pyopenvba.formula._calc.reference"),
    (r"pyofficeeditor\.excel\._collate\b", "pyopenvba.formula._calc.collate"),
    (r"pyofficeeditor\.excel\._calc\b", "pyopenvba.formula._calc"),
    (r"pyofficeeditor\.excel\._numfmt\.format_value\b", "pyopenvba.formula._calc.host.format_value"),
]
#: The engine's own modules that are not copied: pyOfficeEditor's workbook driver, and the package's front.
SKIPPED = frozenset({"engine.py", "__init__.py"})


def _git(source: Path, *args: str) -> str:
    return subprocess.run(["git", "-C", str(source), *args], check=True, capture_output=True, text=True,
                          encoding="utf-8").stdout


def copies(source: Path) -> dict[str, Path]:
    """Each committed module to copy, by its path in pyOfficeEditor, and where it goes."""
    found: dict[str, Path] = {}
    for path in _git(source, "ls-files", "src/pyofficeeditor/excel/_calc").splitlines():
        relative = Path(path).relative_to("src/pyofficeeditor/excel/_calc")
        if path.endswith(".py") and relative.as_posix() not in SKIPPED:
            found[path] = TARGET / relative
    found["src/pyofficeeditor/excel/_collate.py"] = TARGET / "collate.py"
    found["src/pyofficeeditor/excel/_reference.py"] = TARGET / "reference.py"
    return found


def rewritten(text: str) -> str:
    for pattern, replacement in REWRITES:
        text = re.sub(pattern, replacement, text)
    return text


def main() -> None:
    parser = argparse.ArgumentParser(description=(__doc__ or "").split("\n", 1)[0])
    parser.add_argument("source", nargs="?", type=Path, default=DEFAULT_SOURCE, help="a pyOfficeEditor checkout")
    parser.add_argument("--write", action="store_true", help="copy the source over the modules that differ")
    options = parser.parse_args()
    head = _git(options.source, "rev-parse", "--short", "HEAD").strip()
    differing = 0
    for path, target in copies(options.source).items():
        wanted = rewritten(_git(options.source, "show", f"HEAD:{path}"))
        have = target.read_text(encoding="utf-8") if target.exists() else ""
        if have == wanted:
            continue
        differing += 1
        if options.write:
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_text(wanted, encoding="utf-8", newline="\n")
        else:
            print("".join(difflib.unified_diff(have.splitlines(keepends=True), wanted.splitlines(keepends=True),
                                               f"pyOpenVBA/{target.relative_to(ROOT).as_posix()}",
                                               f"pyOfficeEditor@{head}/{path}")))
    print(f"{differing} modules differ from pyOfficeEditor {head}" + ("; copied" if options.write and differing else ""))


if __name__ == "__main__":
    main()
