"""Decode all 34 rows of MSysObjects (page 17 in sample 040) using
a hypothesized Jet 4 row format:
- u16 col_count
- 4B Id (Long)
- 4B ParentId (Long)
- 2B Type (Integer)
- 8B DateCreate
- 8B DateUpdate
- 4B Flags
- variable section...
- u16 jump-table-of-N pointing back to var-column starts
- u8[ceil(cols/8)] null bitmap at end (or just before jump table)

Then look for the row whose Name == "M" (our VBA module) to learn
the module-row layout."""
from pathlib import Path

ACE_PAGE_SIZE = 4096
p = Path("tests/live_access_test/re_corpus/samples/040__sub_msgbox_hello.accdb")
data = p.read_bytes()


def page_bytes(pn: int) -> bytes:
    base = pn * ACE_PAGE_SIZE
    return data[base : base + ACE_PAGE_SIZE]


def decode_row(row: bytes) -> dict[str, object]:
    cols = int.from_bytes(row[0:2], "little")
    id_ = int.from_bytes(row[2:6], "little")
    parent = int.from_bytes(row[6:10], "little")
    typ = int.from_bytes(row[10:12], "little", signed=True)
    date_c = row[12:20]
    date_u = row[20:28]
    flags = int.from_bytes(row[28:32], "little")
    return {
        "cols": cols,
        "Id": id_,
        "ParentId": parent,
        "Type": typ,
        "DateCreate": date_c.hex(),
        "DateUpdate": date_u.hex(),
        "Flags": flags,
        "var_section": row[32:].hex(),
    }


def scan_var_strings(varblob: bytes) -> list[str]:
    """Find UTF-16-LE substrings in the variable section."""
    found: list[str] = []
    i = 0
    while i + 4 <= len(varblob):
        j = i
        ok = 0
        while (
            j + 1 < len(varblob)
            and 32 <= varblob[j] < 127
            and varblob[j + 1] == 0
        ):
            j += 2
            ok += 1
        if ok >= 1:
            try:
                s = varblob[i:j].decode("utf-16-le")
                found.append(s)
            except UnicodeDecodeError:
                pass
            i = j
        else:
            i += 1
    return found


pg = page_bytes(17)
row_count = int.from_bytes(pg[12:14], "little")
print(f"MSysObjects DATA page 17: row_count={row_count}, owner_tdef={int.from_bytes(pg[4:8], 'little')}")
print()

OFFSET_MASK = 0x1FFF

# Read row offsets, then compute lengths by sorting
entries: list[tuple[int, int]] = []  # (idx, offset)
for i in range(row_count):
    ent = int.from_bytes(pg[14 + 2 * i : 16 + 2 * i], "little")
    off = ent & OFFSET_MASK
    entries.append((i, off))

# Build (idx, off, length): sorted by offset, length = next_offset - this_offset
# (highest offset is at top, page_size at end)
sorted_entries = sorted(entries, key=lambda x: x[1])
prev_offsets = sorted(set(off for _, off in entries))
# row at off ends at next-higher-offset OR ACE_PAGE_SIZE
end_of = {
    off: (
        prev_offsets[idx + 1] if idx + 1 < len(prev_offsets) else ACE_PAGE_SIZE
    )
    for idx, off in enumerate(prev_offsets)
}

print(f"{'idx':>4}  {'Id':>10}  {'ParentId':>10}  {'Type':>8}  Name(s) in variable section")
print("-" * 100)
for i, off in entries:
    length = end_of[off] - off
    row = pg[off : off + length]
    info = decode_row(row)
    strings = scan_var_strings(row[32:])
    str_strs = ", ".join(repr(s) for s in strings if len(s) > 0)
    print(
        f"{i:>4}  0x{info['Id']:08x}  0x{info['ParentId']:08x}  "
        f"{info['Type']:>+8d}  {str_strs}"
    )
