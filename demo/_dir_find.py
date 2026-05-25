from pyopenvba.access import AccessFile
import sys

db = AccessFile(sys.argv[1] if len(sys.argv) > 1 else
                "tests/live_access_test/re_corpus/samples/040__sub_msgbox_hello.accdb")
catalog = db._find_catalog_row()
assert catalog
_, _, raw = catalog
labels = {
    0x000F: "PROJECTMODULES", 0x0019: "MODULENAME",
    0x0047: "MODULENAMEUNICODE", 0x002B: "MODULE-TERM",
    0x002C: "MODULECOOKIE", 0x0010: "DIR-TERM",
    0x001A: "MODULESTREAMNAME", 0x0032: "MODULESTREAMNAMEUNICODE",
    0x0031: "MODULEOFFSET",
}
for target, name in labels.items():
    i = 0
    while True:
        j = raw.find(target.to_bytes(2, "little"), i)
        if j < 0:
            break
        i = j + 1
        print(f"{name:25s} at {j:4d}: {raw[j:j+24].hex(' ')}")
