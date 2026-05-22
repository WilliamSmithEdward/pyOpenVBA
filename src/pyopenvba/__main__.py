"""
Command-line entry point.

Usage::

    python -m pyopenvba pull <workbook> <dest_dir>
    python -m pyopenvba push <src_dir> <workbook> [--out <new_path>] [--strict]
    python -m pyopenvba ls   <workbook>
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from pyopenvba import ExcelFile, pull, push


def _cmd_pull(args: argparse.Namespace) -> int:
    written = pull(args.workbook, args.dest, overwrite=not args.no_overwrite)
    for p in written:
        print(p)
    return 0


def _cmd_push(args: argparse.Namespace) -> int:
    updated = push(args.src, args.workbook, out=args.out, strict=args.strict)
    for name in updated:
        print(name)
    return 0


def _cmd_ls(args: argparse.Namespace) -> int:
    with ExcelFile(args.workbook) as wb:
        for m in wb.vba_project().modules:
            print(f"{m.kind.name:8s}  {m.name}")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="pyopenvba")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_pull = sub.add_parser("pull", help="Export VBA modules to a directory.")
    p_pull.add_argument("workbook", type=Path)
    p_pull.add_argument("dest", type=Path)
    p_pull.add_argument("--no-overwrite", action="store_true")
    p_pull.set_defaults(func=_cmd_pull)

    p_push = sub.add_parser("push", help="Import VBA modules from a directory and save.")
    p_push.add_argument("src", type=Path)
    p_push.add_argument("workbook", type=Path)
    p_push.add_argument("--out", type=Path, default=None)
    p_push.add_argument("--strict", action="store_true",
                        help="Fail if any source file has no matching module.")
    p_push.set_defaults(func=_cmd_push)

    p_ls = sub.add_parser("ls", help="List VBA modules in a workbook.")
    p_ls.add_argument("workbook", type=Path)
    p_ls.set_defaults(func=_cmd_ls)

    args = parser.parse_args(argv)
    return int(args.func(args))


if __name__ == "__main__":   # pragma: no cover
    sys.exit(main())
