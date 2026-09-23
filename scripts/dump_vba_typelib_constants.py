"""The constants of VBA's own type library, read from VBE7.DLL.

The pyVBAReference dumps build_object_inventory.py reads leave out the
VBA library's constant modules -- ColorConstants (vbRed), KeyCodeConstants
(vbKeyReturn), SystemColorConstants (vbButtonFace) -- and two enums,
VbQueryClose and FormShowConstants (vbModeless). This reads every
constant of the library straight from the type library Office installs
and writes scripts/vba_typelib_constants.json, which the generator adds
where the dumps have nothing.

    python scripts/dump_vba_typelib_constants.py [path\\to\\VBE7.DLL]
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pythoncom

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "scripts" / "vba_typelib_constants.json"
DEFAULT = Path(r"C:\Program Files\Microsoft Office\root\vfs\ProgramFilesCommonX64\Microsoft Shared\VBA\VBA7.1"
               r"\VBE7.DLL")
#: The VARTYPE a constant is declared with, by VBA's name for it: an enum's members are VT_INT, a Long to VBA,
#: and the Constants module's strings VT_LPSTR.
TYPES = {2: "Integer", 3: "Long", 4: "Single", 5: "Double", 8: "String", 11: "Boolean", 17: "Byte", 22: "Long",
         30: "String"}


def main() -> None:
    path = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT
    library = pythoncom.LoadTypeLib(str(path))
    groups: dict[str, dict[str, object]] = {}
    for index in range(library.GetTypeInfoCount()):
        info = library.GetTypeInfo(index)
        attr = info.GetTypeAttr()
        constants: dict[str, object] = {}
        for var in range(attr.cVars):
            desc = info.GetVarDesc(var)
            declared = desc.elemdescVar[0]
            constants[info.GetNames(desc.memid)[0]] = {"value": desc.value, "type": TYPES.get(declared, declared)}
        if constants:
            groups[library.GetDocumentation(index)[0]] = constants
    OUT.write_text(json.dumps({"source": path.name, "groups": groups}, indent=1, ensure_ascii=False) + "\n",
                   encoding="utf-8")
    print(f"{sum(len(one) for one in groups.values())} constants in {len(groups)} groups -> {OUT.name}")


if __name__ == "__main__":
    main()
