"""Probe: dump the structure of the MSysObjects TDEF page (page 2)
in sample 040, and try to locate its DATA pages. Goal is to learn
the row layout used by the system catalog."""
from pathlib import Path

ACE_PAGE_SIZE = 4096

p = Path("tests/live_access_test/re_corpus/samples/040__sub_msgbox_hello.accdb")
data = p.read_bytes()


def page_bytes(pn: int) -> bytes:
    base = pn * ACE_PAGE_SIZE
    return data[base : base + ACE_PAGE_SIZE]


def hexdump(buf: bytes, width: int = 16, max_bytes: int = 512) -> str:
    out: list[str] = []
    n = min(len(buf), max_bytes)
    for off in range(0, n, width):
        chunk = buf[off : off + width]
        hexs = " ".join(f"{b:02x}" for b in chunk)
        asc = "".join(chr(b) if 32 <= b < 127 else "." for b in chunk)
        out.append(f"  {off:04x}  {hexs:<48}  {asc}")
    if len(buf) > max_bytes:
        out.append(f"  ... ({len(buf) - max_bytes} more bytes)")
    return "\n".join(out)


# ---- 1. Print db header for context ----
print("=" * 70)
print("PAGE 0 (database header) -- first 256 bytes")
print("=" * 70)
print(hexdump(page_bytes(0), max_bytes=256))

# ---- 2. Page 2: MSysObjects TDEF ----
print()
print("=" * 70)
print("PAGE 2 (presumed MSysObjects TDEF) -- first 256 bytes")
print("=" * 70)
pg2 = page_bytes(2)
print(hexdump(pg2, max_bytes=256))

# Hypothesized TDEF layout (Jet 4 / ACE):
#  +0  page_type=0x02
#  +1  unused=0x01
#  +2  free_space (u16-LE)
#  +4  prev_tdef (u32-LE) -- 0 for first
#  +8  next_tdef (u32-LE)
#  +12 ??? (4 bytes; check)
#
# At an offset further in we should find the column count.
print()
print("PAGE 2 header fields (hypothesized):")
print(f"  page_type       = 0x{pg2[0]:02x}")
print(f"  reserved        = 0x{pg2[1]:02x}")
print(f"  free_space      = 0x{int.from_bytes(pg2[2:4], 'little'):04x}")
print(f"  field@4..8      = {pg2[4:8].hex()}  (often zero)")
print(f"  field@8..12     = 0x{int.from_bytes(pg2[8:12], 'little'):08x}")
print(f"  field@12..16    = 0x{int.from_bytes(pg2[12:16], 'little'):08x}")
print(f"  field@16..20    = 0x{int.from_bytes(pg2[16:20], 'little'):08x}")
print(f"  field@20..24    = 0x{int.from_bytes(pg2[20:24], 'little'):08x}")
print(f"  field@24..28    = 0x{int.from_bytes(pg2[24:28], 'little'):08x}")
print(f"  field@28..32    = 0x{int.from_bytes(pg2[28:32], 'little'):08x}")

# ---- 3. Walk continuation chain via next_tdef field ----
# In Jet 4 the TDEF chain pointer is at offset 4 (next page) for type 0x02.
# We saw page 2 head bytes [8:12] = 1a 04 00 00 -> 0x041a = 1050 bytes used.
# Let's check what's at the offset claimed by [4:8].
print()
nxt = int.from_bytes(pg2[4:8], "little")
print(f"  next-page-from-[4:8] = 0x{nxt:08x}")
if 0 < nxt < len(data) // ACE_PAGE_SIZE:
    print(f"  -> page {nxt} type=0x{page_bytes(nxt)[0]:02x}")

# ---- 4. ASCII strings inside page 2 (look for table names) ----
print()
print("ASCII strings >= 4 chars in page 2:")
buf = pg2
i = 0
while i < len(buf):
    start = i
    while i < len(buf) and 32 <= buf[i] < 127:
        i += 1
    if i - start >= 4:
        s = buf[start:i].decode("ascii", errors="replace")
        print(f"  @0x{start:04x}  {s!r}")
    i += 1

# ---- 5. UTF-16-LE strings >= 4 chars inside page 2 ----
print()
print("UTF-16-LE strings >= 4 chars in page 2 (lone-byte BMP):")
i = 0
while i + 8 <= len(buf):
    # Heuristic: at least 4 chars where every odd byte is 0 and even byte is ASCII printable
    j = i
    ok = 0
    while j + 1 < len(buf) and 32 <= buf[j] < 127 and buf[j + 1] == 0:
        j += 2
        ok += 1
    if ok >= 4:
        s = buf[i:j].decode("utf-16-le", errors="replace")
        print(f"  @0x{i:04x}  ({ok} chars)  {s!r}")
        i = j
    else:
        i += 1
