"""
create_files.py
---------------
Create blank, ready-to-use Office files (Excel, Word, PowerPoint) using
pyOpenVBA's built-in templates.  Each file contains an empty Module1 and
no other user code.

Run from the repo root or the demo/ folder:

    python demo/create_files.py
    python demo/create_files.py --out-dir my_output
"""

import argparse
from pathlib import Path

from pyopenvba import ExcelFile, WordFile, PowerPointFile


def create_blank_files(out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)

    targets = [
        ("blank_workbook.xlsm",    ExcelFile),
        ("blank_document.docm",    WordFile),
        ("blank_presentation.pptm", PowerPointFile),
    ]

    for filename, cls in targets:
        path = out_dir / filename
        with cls.create_new(path) as f:   # type: ignore[attr-defined]
            names = f.module_names()
        print(f"  created  {path}  ({len(names)} module(s): {', '.join(names)})")


def main() -> None:
    parser = argparse.ArgumentParser(description="Create blank macro-enabled Office files.")
    parser.add_argument(
        "--out-dir",
        default=str(Path(__file__).parent / "output"),
        help="Directory to write the files into (default: demo/output/)",
    )
    args = parser.parse_args()

    out_dir = Path(args.out_dir)
    print(f"Writing blank files to: {out_dir.resolve()}")
    create_blank_files(out_dir)
    print("Done.")


if __name__ == "__main__":
    main()
