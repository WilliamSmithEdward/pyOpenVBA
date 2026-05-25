"""
Command-line entry point.

Usage::

    python -m pyopenvba pull <workbook> <dest_dir>
    python -m pyopenvba push <src_dir> <workbook> [--out <new_path>] [--strict]
    python -m pyopenvba ls   <workbook>
    python -m pyopenvba access-ls    <accdb>
    python -m pyopenvba access-pull  <accdb> <dest_dir>
    python -m pyopenvba access-disasm <accdb> [--module <name>] [--with-source]
    python -m pyopenvba disasm <workbook|doc|pptx> [--module <name>] [--with-source]
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from pyopenvba import ExcelFile, PowerPointFile, WordFile, pull, push


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
    from pyopenvba.access_read import AccessReader

    db = AccessReader(args.database)
    for m in db.iter_vba_modules():
        print(m.name)
    return 0


def _cmd_access_pull(args: argparse.Namespace) -> int:
    from pyopenvba.access_read import AccessReader

    db = AccessReader(args.database)
    args.dest.mkdir(parents=True, exist_ok=True)
    for m in db.iter_vba_modules():
        out = args.dest / f"{m.name}.bas"
        out.write_text(m.source, encoding="utf-8")
        print(out)
    return 0


def _cmd_access_disasm(args: argparse.Namespace) -> int:
    from pyopenvba.access_read import AccessReader

    db = AccessReader(args.database)
    sources: dict[str, str] = {}
    if args.with_source:
        sources = {m.name: m.source for m in db.iter_vba_modules()}
    if args.module is not None:
        mod = db.disassemble_module(args.module)
        src = sources.get(args.module, "") if args.with_source else ""
        listing = (
            mod.to_annotated_listing(src)
            if args.with_source
            else mod.to_listing()
        )
        print(f"===== {args.module} =====")
        print(listing)
        return 0
    for name, mod in db.disassemble_all_modules().items():
        print(f"===== {name} =====")
        if args.with_source:
            print(mod.to_annotated_listing(sources.get(name, "")))
        else:
            print(mod.to_listing())
    return 0


_HOST_BY_SUFFIX: dict[str, type] = {
    ".xlsm": ExcelFile,
    ".xlsb": ExcelFile,
    ".xltm": ExcelFile,
    ".xlam": ExcelFile,
    ".docm": WordFile,
    ".dotm": WordFile,
    ".pptm": PowerPointFile,
    ".potm": PowerPointFile,
    ".ppam": PowerPointFile,
}


def _cmd_disasm(args: argparse.Namespace) -> int:
    suffix = args.workbook.suffix.lower()
    host_cls = _HOST_BY_SUFFIX.get(suffix)
    if host_cls is None:
        print(
            f"error: unsupported file type {suffix!r}; expected one of "
            f"{sorted(_HOST_BY_SUFFIX)}",
            file=sys.stderr,
        )
        return 2
    with host_cls(args.workbook) as host:
        project = host.vba_project()
        modules = list(project.modules)
    if args.module is not None:
        modules = [m for m in modules if m.name == args.module]
        if not modules:
            print(
                f"error: module {args.module!r} not found",
                file=sys.stderr,
            )
            return 1
    for mod in modules:
        disasm = mod.disassemble()
        print(f"===== {mod.name} =====")
        if args.with_source:
            print(disasm.to_annotated_listing(mod.source))
        else:
            print(disasm.to_listing())
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
    p_ad.add_argument(
        "--with-source",
        action="store_true",
        help="Interleave original VBA source lines with the p-code.",
    )
    p_ad.set_defaults(func=_cmd_access_disasm)

    p_d = sub.add_parser(
        "disasm",
        help=(
            "Disassemble compiled VBA p-code from an Excel/Word/"
            "PowerPoint host (.xlsm/.docm/.pptm/...)."
        ),
    )
    p_d.add_argument("workbook", type=Path)
    p_d.add_argument(
        "--module",
        default=None,
        help="Disassemble just this module (default: every module).",
    )
    p_d.add_argument(
        "--with-source",
        action="store_true",
        help="Interleave original VBA source lines with the p-code.",
    )
    p_d.set_defaults(func=_cmd_disasm)

    args = parser.parse_args(argv)
    return int(args.func(args))


if __name__ == "__main__":   # pragma: no cover
    sys.exit(main())
