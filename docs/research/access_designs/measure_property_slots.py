"""Where each named property sits in each object type's schema.

A record's id is its slot, and the slot differs by control type, so
changing a property that is not already on an object means knowing the id
Access would have given it.  Reading that off objects Access itself wrote,
across every type, is what makes the slot table.
"""

import pathlib
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[3] / "src"))

from pyopenvba.access import AccessDatabase  # noqa: E402
from pyopenvba.access._designs import CONTROL_TYPES, PROPERTY_CODES  # noqa: E402

#: Where the databases Access built were left; pass another with argv[1].
SCRATCH = Path(sys.argv[1] if len(sys.argv) > 1 else ".")
#: Every database under it that Access built, however it was built: the
#: more objects the sweep sees, the more slots it can confirm.
SOURCES = sorted(SCRATCH.rglob("*.accdb"))
NAMES = {code: name for name, code in PROPERTY_CODES.items()}
ROOT = "_Design"
#: Value types whose length comes from the value, not from the slot.
VARIABLE = (10, 11, 12)

Slot = tuple[int, int, int, int, int]


def main() -> None:
    seen: dict[str, dict[str, set[Slot]]] = defaultdict(lambda: defaultdict(set))
    for path in SOURCES:
        if not path.exists():
            print("missing:", path.name)
            continue
        try:
            db = AccessDatabase(path)
            forms = [f.name for f in db.forms()]
        except Exception as exc:  # a database with no designs in it
            print(f"skipped {path.name}: {exc}")
            continue
        for name in forms:
            design = db.form(name)
            for obj in design.objects:
                kind = ROOT if obj.type is None else CONTROL_TYPES.get(obj.type, "")
                if not kind:
                    continue
                for r in obj.records:
                    prop = NAMES.get(r.code)
                    if prop is None:
                        continue
                    # A string's length is its text's; only the fixed
                    # types have a length that belongs to the slot.
                    size = 0 if r.value_type in VARIABLE else len(r.value)
                    seen[kind][prop].add((r.id, r.code, r.value_type, r.width, size))

    steady: dict[str, dict[str, Slot]] = {}
    wobbly: list[str] = []
    for kind, props in sorted(seen.items()):
        for prop, slots in sorted(props.items()):
            ids = {(s[0], s[2], s[4]) for s in slots}
            if len(ids) == 1:
                # The width Access wrote can differ run to run; the id,
                # code and value type are what pin the slot.
                one = sorted(slots)[0]
                steady.setdefault(kind, {})[prop] = one
            else:
                wobbly.append(f"{kind}.{prop}: {sorted(ids)}")

    print(f"{len(steady)} object types, "
          f"{sum(len(v) for v in steady.values())} steady slots, "
          f"{len(wobbly)} that moved")
    for line in wobbly[:20]:
        print("  moved:", line)

    out = Path(__file__).parent / "slots.txt"
    with out.open("w", encoding="utf-8") as fh:
        for kind, props in sorted(steady.items()):
            fh.write(f'    "{kind}": {{\n')
            for prop, (ident, code, vtype, width, size) in sorted(
                props.items(), key=lambda kv: kv[1][0]
            ):
                fh.write(f'        "{prop}": ({ident}, {code}, {vtype}, {width}, {size}),\n')
            fh.write("    },\n")
    print("written to", out)
    for kind, props in sorted(steady.items()):
        print(f"  {kind:20} {len(props)} properties")


if __name__ == "__main__":
    main()
