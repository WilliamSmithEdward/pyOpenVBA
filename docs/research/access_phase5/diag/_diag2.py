"""Compare DataModel module stream + dir entries vs Module1."""
import zipfile, struct
from pyopenvba.cfb import CFB
from pyopenvba.vba import decompress, parse_vba_project

z = zipfile.ZipFile("demo/output/C_module1_plus_datamodel.pptm")
raw = z.read("ppt/vbaProject.bin")
cfb = CFB.from_bytes(raw)
proj = parse_vba_project(cfb)

print("=== Modules ===")
for m in proj.modules:
    print(f"  {m.name}: kind={m.kind.name} stream={m.stream_name!r} text_offset={m.text_offset} prefix_len={len(m.prefix_bytes)}")
    print(f"    source[:200]={m.source[:200]!r}")
    print()

# Inspect the raw stream of DataModel
dm_raw = cfb.get_stream_in_storage("VBA", "DataModel")
print(f"DataModel stream raw len = {len(dm_raw)}")
print(f"DataModel decompressed = {decompress(dm_raw)!r}")
print()
mod1_raw = cfb.get_stream_in_storage("VBA", "Module1")
print(f"Module1 stream raw len = {len(mod1_raw)}")
print(f"Module1 first 4 bytes  = {mod1_raw[:4].hex()}")
# Decompressed source portion only (after text_offset)
m1 = next(m for m in proj.modules if m.name == "Module1")
print(f"Module1 text_offset={m1.text_offset}")
print(f"Module1 decompressed source (post-offset) = {decompress(mod1_raw[m1.text_offset:])!r}")
