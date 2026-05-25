"""Probe: scan sample 040 for occurrences of module name "M" or
"MyModule" in DATA-page-typed pages (page_type != 0x01 LVAL,
!= 0x00 header). This should reveal MSysObjects table rows that
reference the VBA module by name."""
from pathlib import Path

ACE_PAGE_SIZE = 4096

p = Path("tests/live_access_test/re_corpus/samples/040__sub_msgbox_hello.accdb")
data = p.read_bytes()
n_pages = len(data) // ACE_PAGE_SIZE

# Page type byte at offset 0. Known: 0x00 = db header, 0x01 = data
# table (LVAL is page type 0x01 with tag "LVAL" at [4:8]; the LVAL
# tag is specifically a long-value page; other 0x01 pages are
# regular DATA pages). Let me actually print the page-type byte
# distribution first.
counts: dict[int, int] = {}
for i in range(n_pages):
    pt = data[i * ACE_PAGE_SIZE]
    counts[pt] = counts.get(pt, 0) + 1
print(f"page count = {n_pages}")
print("page-type distribution:", dict(sorted(counts.items())))

# Print first 32 bytes of each page so we can categorize.
for i in range(min(n_pages, 90)):
    base = i * ACE_PAGE_SIZE
    head = data[base : base + 32]
    tag = data[base + 4 : base + 8]
    pt = data[base]
    hex_ = " ".join(f"{b:02x}" for b in head[:16])
    ascii_tag = tag.decode("ascii", errors="replace")
    print(f"  page {i:3d} type=0x{pt:02x} tag={ascii_tag!r} head={hex_}")
