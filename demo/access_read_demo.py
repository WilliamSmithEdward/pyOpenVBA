"""Demo: read everything out of an Access .accdb VBA project.

Showcases the fully-working read story for Access:

  - List modules (catalog-ordered, std vs class)
  - Project metadata (codepage, syskind, lcid, references, module flags)
  - Module source (CRLF body, with-Attribute form)
  - Bulk export to .bas / .cls files
  - Disassemble VBA7 p-code per module (source-annotated listing)
  - Identifier table
  - Interned string literals

This is what's locked in today and byte-for-byte verified against
Access COM across the project's RE corpus.
"""

from __future__ import annotations

import shutil
from pathlib import Path

from pyopenvba.access_read import AccessReader

REPO = Path(__file__).resolve().parent.parent
SRC = REPO / "tests" / "live_access_test" / "New Microsoft Access Database.accdb"
OUT = REPO / "demo" / "output" / "access_read_demo"


def _hr(title: str) -> None:
    print()
    print("=" * 70)
    print(f"  {title}")
    print("=" * 70)


def main() -> None:
    if not SRC.exists():
        raise SystemExit(f"missing fixture: {SRC}")

    OUT.mkdir(parents=True, exist_ok=True)
    work = OUT / SRC.name
    shutil.copy(SRC, work)
    print(f"Working copy: {work}")

    db = AccessReader(work)

    _hr("1. Module catalog")
    names = db.vba_module_names()
    print(f"  vba_module_names() -> {names}")

    proj = db.read_project_info()
    print(f"  project name      : {proj.project_name!r}")
    print(f"  code page         : {proj.code_page}")
    print(f"  syskind           : {proj.sys_kind}")
    print(f"  lcid              : {proj.lcid}")
    print(f"  catalog row       : page={proj.catalog_page} slot={proj.catalog_slot}")
    print(f"  references        : {len(proj.references)}")
    for ref in proj.references[:6]:
        print(f"    - {ref}")
    print(f"  modules           : {len(proj.modules)}")
    for m in proj.modules:
        kind = "Class" if m.is_class_module else "Std  "
        flags: list[str] = []
        if m.is_private:
            flags.append("Private")
        if m.is_read_only:
            flags.append("ReadOnly")
        fstr = f" [{','.join(flags)}]" if flags else ""
        print(f"    [{kind}] {m.name:<16} stream={m.stream_name!r}{fstr}")

    _hr("2. Module source (raw bodies)")
    bodies = db.vba_modules()
    for name, body in bodies.items():
        lines = body.splitlines()
        print(f"  --- {name} ({len(body)} chars, {len(lines)} lines) ---")
        for ln in lines[:14]:
            print(f"    {ln}")
        if len(lines) > 14:
            print(f"    ... ({len(lines) - 14} more lines)")

    _hr("3. Module source with VB_Attribute preamble (one example)")
    first = names[0]
    with_attrs = db.read_vba_module_with_attributes(first)
    print(f"  read_vba_module_with_attributes({first!r}):")
    for ln in with_attrs.splitlines()[:15]:
        print(f"    {ln}")

    _hr("4. Bulk export to .bas / .cls files")
    export_dir = OUT / "exported"
    if export_dir.exists():
        shutil.rmtree(export_dir)
    paths = db.export_modules(export_dir, include_attributes=True)
    for p in paths:
        size = p.stat().st_size
        print(f"  wrote {p.relative_to(REPO)} ({size} bytes)")

    _hr("5. P-code disassembly (source-annotated)")
    for name in names:
        try:
            dasm = db.disassemble_module(name)
        except Exception as exc:  # noqa: BLE001
            print(f"  {name}: no p-code ({exc})")
            continue
        total_ops = sum(len(line.instructions) for line in dasm.lines)
        print(
            f"  --- {name}: cafe_offset=0x{dasm.cafe_offset:X} "
            f"num_lines={dasm.num_lines} opcodes={total_ops} ---"
        )
        body = bodies.get(name, "")
        listing = dasm.to_annotated_listing(body).splitlines()
        for ln in listing[:24]:
            print(f"    {ln}")
        if len(listing) > 24:
            print(f"    ... ({len(listing) - 24} more lines)")

    _hr("6. Identifier table")
    ids = db.identifiers()
    print(f"  total identifiers: {len(ids)}")
    for ident in ids[:30]:
        print(
            f"    [{ident.index:>3}] type=0x{ident.type_byte:02X} "
            f"id=0x{ident.id_low:04X} {ident.name!r}"
        )
    if len(ids) > 30:
        print(f"    ... ({len(ids) - 30} more)")

    _hr("7. Interned string literals")
    strings = db.find_interned_strings()
    print(f"  total interned strings: {len(strings)}")
    for s in strings[:30]:
        print(f"    page={s.page:>4} slot={s.slot:>3} {s.value!r}")
    if len(strings) > 30:
        print(f"    ... ({len(strings) - 30} more)")

    _hr("Done")
    print(f"  Exported source: {export_dir}")
    print(f"  Working copy   : {work}")


if __name__ == "__main__":
    main()
