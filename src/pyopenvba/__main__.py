"""
Command-line entry point.

Usage::

    python -m pyopenvba pull <workbook> <dest_dir>
    python -m pyopenvba push <src_dir> <workbook> [--out <new_path>] [--strict]
    python -m pyopenvba ls   <workbook>
    python -m pyopenvba access-ls    <accdb>
    python -m pyopenvba access-pull  <accdb> <dest_dir>
    python -m pyopenvba access-disasm <accdb> [--module <name>]
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


def _cmd_access_ls(args: argparse.Namespace) -> int:
    from pyopenvba.access import AccessFile

    db = AccessFile(args.database)
    for m in db.iter_vba_modules():
        print(m.name)
    return 0


def _cmd_access_pull(args: argparse.Namespace) -> int:
    from pyopenvba.access import AccessFile

    db = AccessFile(args.database)
    args.dest.mkdir(parents=True, exist_ok=True)
    for m in db.iter_vba_modules():
        out = args.dest / f"{m.name}.bas"
        out.write_text(m.source, encoding="utf-8")
        print(out)
    return 0


def _cmd_access_disasm(args: argparse.Namespace) -> int:
    from pyopenvba.access import AccessFile

    db = AccessFile(args.database)
    if args.module is not None:
        listing = db.disassemble_module(args.module).to_listing()
        print(f"===== {args.module} =====")
        print(listing)
        return 0
    for name, mod in db.disassemble_all_modules().items():
        print(f"===== {name} =====")
        print(mod.to_listing())
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

    p_als = sub.add_parser(
        "access-ls",
        help="List VBA modules in an Access database (.accdb/.mdb).",
    )
    p_als.add_argument("database", type=Path)
    p_als.set_defaults(func=_cmd_access_ls)

    p_apl = sub.add_parser(
        "access-pull",
        help="Export VBA modules from an Access database to a directory.",
    )
    p_apl.add_argument("database", type=Path)
    p_apl.add_argument("dest", type=Path)
    p_apl.set_defaults(func=_cmd_access_pull)

    p_ad = sub.add_parser(
        "access-disasm",
        help="Disassemble compiled VBA p-code from an Access database.",
    )
    p_ad.add_argument("database", type=Path)
    p_ad.add_argument(
        "--module",
        default=None,
        help="Disassemble just this module (default: every module).",
    )
    p_ad.set_defaults(func=_cmd_access_disasm)

    args = parser.parse_args(argv)
    return int(args.func(args))


if __name__ == "__main__":   # pragma: no cover
    sys.exit(main())
