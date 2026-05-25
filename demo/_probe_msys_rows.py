"""Probe: try to decode the row layout on a DATA page (page 6 in
sample 040), assuming it belongs to MSysObjects. Row format on Jet 4:
- u16 row_count at offset 8 (or 12 depending on version)
- Each row entry in the offset table is u16 with the high bits used
  for flags (0xC000 = mask, 0x8000 = deleted, 0x4000 = overflow).
- Row data starts from end-of-page downward."""
from pathlib import Path

ACE_PAGE_SIZE = 4096

p = Path("tests/live_access_test/re_corpus/samples/040__sub_msgbox_hello.accdb")
data = p.read_bytes()


def page_bytes(pn: int) -> bytes:
    base = pn * ACE_PAGE_SIZE
    return data[base : base + ACE_PAGE_SIZE]


def hexdump(buf: bytes, start: int, length: int) -> str:
    out: list[str] = []
    end = start + length
    for off in range(start, end, 16):
        chunk = buf[off : min(off + 16, end)]
        hexs = " ".join(f"{b:02x}" for b in chunk)
        asc = "".join(chr(b) if 32 <= b < 127 else "." for b in chunk)
        out.append(f"  {off:04x}  {hexs:<48}  {asc}")
    return "\n".join(out)


def probe_data_page(pn: int) -> None:
    pg = page_bytes(pn)
    print()
    print("=" * 70)
    print(f"DATA PAGE {pn}")
    print("=" * 70)
    print(f"  type=0x{pg[0]:02x} reserved=0x{pg[1]:02x}")
    print(f"  free_space     = 0x{int.from_bytes(pg[2:4], 'little'):04x}")
    print(f"  owner_tdef_pn  = 0x{int.from_bytes(pg[4:8], 'little'):08x}")
    print(f"  field@8..12    = 0x{int.from_bytes(pg[8:12], 'little'):08x}")
    row_count = int.from_bytes(pg[12:14], "little")
    print(f"  row_count@12   = {row_count}")

    # The row offset table starts at offset 14 (one u16 per row).
    print(f"  row offset table:")
    DELETED = 0x8000
    OVERFLOW = 0x4000
    OFFSET_MASK = 0x1FFF  # safe lower-bits mask for 4 KiB pages
    rows: list[tuple[int, int, int, int]] = []
    last_top = ACE_PAGE_SIZE
    for i in range(row_count):
        ent = int.from_bytes(pg[14 + 2 * i : 16 + 2 * i], "little")
        flags = ent & 0xE000
        off = ent & OFFSET_MASK
        if off > 0:
            length = last_top - off
            last_top = off
        else:
            length = 0
        rows.append((i, ent, off, length))
        flag_str = []
        if flags & DELETED:
            flag_str.append("DELETED")
        if flags & OVERFLOW:
            flag_str.append("OVERFLOW")
        if flags & 0x2000:
            flag_str.append("?0x2000")
        print(
            f"    row {i:3d}  raw=0x{ent:04x}  off=0x{off:04x}  "
            f"len={length:4d}  flags={','.join(flag_str) or '-'}"
        )

    # Dump the first few rows
    print()
    print("  first 5 row payloads:")
    for i, ent, off, length in rows[:5]:
        if off == 0 or length == 0:
            continue
        print(f"  -- row {i} -- off=0x{off:04x} len={length}")
        print(hexdump(pg, off, min(length, 96)))


# Probe the suspected MSysObjects data pages
for pn in [6, 9, 11, 13, 17, 24, 40, 41]:
    probe_data_page(pn)
