"""Inspect the real class module added by the user to Presentation1.pptm."""
import zipfile
from pyopenvba.cfb import CFB
from pyopenvba.vba import decompress, parse_vba_project

z = zipfile.ZipFile("tests/live_powerpoint_testing/Presentation1.pptm")
raw = z.read("ppt/vbaProject.bin")
cfb = CFB.from_bytes(raw)
proj = parse_vba_project(cfb)

print("Modules:", [(m.name, m.kind.name, m.stream_name) for m in proj.modules])
print()
print("=== PROJECT ===")
print(cfb.get_stream("PROJECT").decode("cp1252", errors="replace"))
print()
for m in proj.modules:
    if m.kind.name == "standard" and m.name == "Module1":
        continue
    print(f"=== {m.name} (kind={m.kind.name}) stream={m.stream_name!r} text_offset={m.text_offset} prefix_len={len(m.prefix_bytes)} ===")
    raw_stream = cfb.get_stream_in_storage("VBA", m.stream_name)
    print(f"  raw stream len = {len(raw_stream)}")
    print(f"  prefix (raw bytes) hex[:80] = {m.prefix_bytes[:80].hex()}")
    print(f"  source[:400] = {m.source[:400]!r}")
    print()
