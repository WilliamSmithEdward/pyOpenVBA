import zipfile
from pyopenvba.cfb import CFB

for label, path in [
    ("ORIGINAL", "tests/live_powerpoint_testing/Presentation1.pptm"),
    ("OUTPUT",   "demo/output/all_types_from_presentation1.pptm"),
]:
    z = zipfile.ZipFile(path)
    raw = z.read("ppt/vbaProject.bin")
    cfb = CFB.from_bytes(raw)
    print(f"=== {label}: {path} ===")
    print("  top storages:", cfb.list_storages())
    print("  VBA streams :", cfb.list_streams_in_storage("VBA"))
    try:
        print("  UF1 streams :", cfb.list_streams_in_storage("VBA/UserForm1"))
    except Exception as e:
        print("  UF1 streams : (none)", e)
    print()
    print("  PROJECT:")
    for line in cfb.get_stream("PROJECT").decode("cp1252", errors="replace").splitlines():
        print("    " + line)
    print()
