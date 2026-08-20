"""Rename a module in an .accdb -- research prototype."""
import shutil, sys
from pathlib import Path
ROOT = Path("F:/GitHub/pyOpenVBA")
for q in (ROOT/'src', ROOT/'docs/research/access_write'):
    sys.path.insert(0, str(q))
from accdb_write import (Perf, load_module, write_module, drop_srp_cache,
                         rename_in_dir, rename_in_row, find_moduleoffset_pos,
                         write_row, row_extent, ACE_PAGE_SIZE)
from accdb_write import set_lval_payload
from pyopenvba.access_read import (AccessReader, MSYS_TYPE_MODULE,
                                   PAGE_TYPE_DATA)

# The name is written inline in exactly three rows. Two other rows also
# contain it -- the project identifier table and one on the far side of
# the file -- and Access leaves both holding the *old* name, so rewriting
# every row that mentions it corrupts the project.
INLINE_STREAMS = ("PROJECTwm", "DirData")


def _storage_rows(data, old):
    """(page, slot) of the PROJECTwm and DirData rows naming this module."""
    needle = old.encode("utf-16-le")
    tags = [s.encode("utf-16-le") for s in INLINE_STREAMS]
    out = []
    for page in range(len(data)//ACE_PAGE_SIZE):
        base = page*ACE_PAGE_SIZE
        if data[base] != PAGE_TYPE_DATA:
            continue
        n = int.from_bytes(data[base+12:base+14], "little")
        if not n or 14+2*n > ACE_PAGE_SIZE:
            continue
        for slot in range(n):
            entry = int.from_bytes(data[base+14+2*slot:base+16+2*slot], "little")
            if entry >> 12:
                continue
            s, e = row_extent(data, base, slot)
            row = bytes(data[base+s:base+e])
            if needle in row and any(tag in row for tag in tags):
                out.append((page, slot))
    return out


def rename_module(src, dst, old, new):
    shutil.copy(src, dst)
    info = load_module(dst, old)
    perf = Perf(info["row"], info["modoff"])
    renamed = rename_in_dir(info["dir_dec"], old, new)
    info = dict(info, dir_dec=renamed,
                modoff_pos=find_moduleoffset_pos(renamed, new))
    attrs = [l.replace(f'"{old}"', f'"{new}"') for l in perf.attribute_lines()]
    blob = "\r\n".join(attrs + perf.source_lines()).encode("latin-1")
    row, modoff = perf.build(new_source=blob)
    data = bytearray(dst.read_bytes())
    write_module(data, info, row, modoff)
    obj = next(x for x in AccessReader(dst).msys_objects()
               if x.type_ == MSYS_TYPE_MODULE and x.name == old)
    targets = _storage_rows(data, old) + [(obj.page, obj.slot)]
    # Descending, so resizing one row cannot move another still to come.
    for page, slot in sorted(targets, reverse=True):
        base = page*ACE_PAGE_SIZE
        s, e = row_extent(data, base, slot)
        write_row(data, page, slot,
                  rename_in_row(bytes(data[base+s:base+e]), old, new))
    # The MS-OVBA PROJECT stream lists the module by name in plain text.
    # This is the one Access actually reads for the VBE's module list --
    # renaming everything else leaves the old name showing.
    reader = AccessReader(dst)
    for page, slot, prow in reader._iter_lval_rows():
        raw = bytes(prow)
        if b"Module=" + old.encode("latin-1") not in raw:
            continue
        set_lval_payload(
            data, page, slot,
            raw.replace(b"Module=" + old.encode("latin-1"),
                        b"Module=" + new.encode("latin-1")),
            len(raw))
        targets.append((page, slot))
        break
    drop_srp_cache(data)
    dst.write_bytes(bytes(data))
    return targets


if __name__ == "__main__":
    sp = Path(sys.argv[1])
    t = rename_module(sp/"base.accdb", sp/"ren_full.accdb", "Alpha", "Beta")
    print(f"rewrote inline rows at {t}")
    r = AccessReader(sp/"ren_full.accdb")
    print("  reader streams:", [x.name for x in r.find_module_streams()])
    print("  MSysObjects   :", [(x.name, x.id_) for x in r.msys_objects()
                                if x.type_ == MSYS_TYPE_MODULE])
