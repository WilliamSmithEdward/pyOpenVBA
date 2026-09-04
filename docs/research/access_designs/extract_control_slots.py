"""Turn what Access wrote into a CONTROL_SLOTS table."""

import sys

sys.path.insert(0, "F:/GitHub/pyOpenVBA/src")

from pyopenvba.access import AccessDatabase  # noqa: E402
from pyopenvba.access._designs import CONTROL_TYPES, PROPERTY_CODES  # noqa: E402

BASE = (
    "C:/Users/William/AppData/Local/Temp/claude/F--GitHub-pyOpenVBA/"
    "01fc1bef-010d-4cbe-a6dc-14dbfb128b42/scratchpad/controls/"
)
NAMES = {code: name for name, code in PROPERTY_CODES.items()}
# Codes seen on the new controls that the property table does not name yet.
EXTRA = {
    0: "Unknown0",
    261: "PictureType",
    700: "TopPadding",
    701: "BottomPadding",
    702: "LeftPadding",
    703: "RightPadding",
    27: "ControlSource",
    35: "FontSize",
    37: "FontWeight",
    204: "ForeColor",
}


def label(code: int) -> str:
    return NAMES.get(code) or EXTRA.get(code) or f"code{code}"


def main() -> None:
    db = AccessDatabase(BASE + "rest.accdb")
    design = db.form("Form1")
    for obj in design.objects:
        if obj.type is None:
            continue
        kind = CONTROL_TYPES.get(obj.type, str(obj.type))
        named = [r for r in obj.records if r.code == 20]
        if not named:
            continue  # a prototype, not one of the controls we made
        name = named[0].value.decode("utf-16-le")
        if not name.startswith("X"):
            continue
        print(f'    "{kind}": {{')
        for r in obj.records:
            print(f"        {label(r.code)!r}: ({r.id}, {r.code}, {r.value_type}, {r.width}),"
                  f"  # {r.value.hex()}")
        print("    },")


if __name__ == "__main__":
    main()
