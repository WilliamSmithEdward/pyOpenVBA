"""One-off probe of decompressed dir-stream layout."""
import struct
import sys

from pyopenvba.access import AccessFile

db = AccessFile(sys.argv[1] if len(sys.argv) > 1 else
                "tests/live_access_test/re_corpus/samples/040__sub_msgbox_hello.accdb")
found = db._find_catalog_row()
assert found
page, slot, raw = found
print(f"catalog page={page} slot={slot} raw_len={len(raw)}")
pos = 0
n = 0
labels = {
    0x0001: "PROJECTSYSKIND", 0x0002: "PROJECTLCID", 0x0003: "PROJECTCODEPAGE",
    0x0004: "PROJECTNAME", 0x0005: "PROJECTDOCSTRING", 0x0006: "PROJECTHELPFILE",
    0x0007: "PROJECTHELPCTX", 0x0008: "PROJECTLIBFLAGS", 0x0009: "PROJECTVERSION",
    0x000F: "PROJECTMODULES", 0x0010: "DIR-TERM", 0x0013: "PROJECTCOOKIE",
    0x0014: "PROJECTLCIDINVOKE", 0x0019: "MODULENAME", 0x001A: "MODULESTREAMNAME",
    0x001C: "MODULEDOCSTRING", 0x001E: "MODULEHELPCTX", 0x0021: "MODULETYPE_STD",
    0x0022: "MODULETYPE_OTHER", 0x0025: "MODULEREADONLY", 0x0028: "MODULEPRIVATE",
    0x002B: "MODULE-TERM", 0x002C: "MODULECOOKIE", 0x0031: "MODULEOFFSET",
    0x0032: "MODULESTREAMNAMEUNICODE", 0x003C: "PROJECTCONSTANTSUNICODE",
    0x003D: "PROJECTHELPFILEPATH2", 0x003E: "REFERENCENAMEUNICODE",
    0x0040: "PROJECTDOCSTRINGUNICODE", 0x0047: "MODULENAMEUNICODE",
    0x0048: "MODULEDOCSTRINGUNICODE", 0x004A: "PROJECTCOMPATVERSION",
    0x000C: "PROJECTCONSTANTS", 0x000D: "REFERENCEREGISTERED",
    0x000E: "REFERENCEPROJECT", 0x0016: "REFERENCENAME",
    0x002F: "REFERENCECONTROL_TWIDDLED", 0x0030: "REFERENCECONTROL_EXTENDED",
    0x0033: "REFERENCEORIGINAL",
}
while pos + 2 <= len(raw):
    rid = struct.unpack_from("<H", raw, pos)[0]
    if rid == 0x0009:
        print(f"  pos={pos:5d} id=0x{rid:04x} PROJECTVERSION (fixed 10B)")
        pos += 10
        n += 1
        continue
    if rid == 0x0010:
        print(f"  pos={pos:5d} id=0x{rid:04x} DIR-TERM")
        break
    sz = struct.unpack_from("<I", raw, pos + 2)[0]
    data = raw[pos + 6 : pos + 6 + sz]
    name = labels.get(rid, "?")
    suffix = ""
    if rid == 0x0019:
        suffix = f"  mbcs={data!r}"
    elif rid == 0x0047:
        suffix = f"  u16={data.decode('utf-16-le', errors='replace')!r}"
    elif rid == 0x001A:
        suffix = f"  mbcs={data!r}"
    elif rid == 0x000F:
        suffix = f"  count={struct.unpack('<H', data)[0]}"
    elif rid == 0x0004:
        suffix = f"  {data!r}"
    print(f"  pos={pos:5d} id=0x{rid:04x} sz={sz:3d} {name}{suffix}")
    pos += 6 + sz
    n += 1
print(f"  total records: {n}")
