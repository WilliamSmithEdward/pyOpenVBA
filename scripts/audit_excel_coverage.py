"""Inventory Excel API bindings and missing names without requiring Office.

Run from a source checkout::

    python scripts/audit_excel_coverage.py --output excel-coverage.json

This is a discovery report, not a conformance score. A registered method
may be a stub or only implement some arguments. The reference inventory
contains interfaces, aliases, events and UI members, so it is not a
count of distinct headless features. Behavioral gates live in
docs/host_completeness.md.
"""

from __future__ import annotations

import argparse
import inspect
import json
from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))


def report() -> dict[str, object]:
    from pyopenvba.apps.excel import _model, _shapes
    from pyopenvba.formula._functions import known_names
    from pyopenvba.formula._inventory import unimplemented
    from pyopenvba.interpreter._inventory_data import MEMBERS
    from pyopenvba.interpreter._objects import VBAObject
    from pyopenvba.mlang._inventory import known
    from pyopenvba.mlang._library import LIBRARY

    bindings: list[dict[str, object]] = []
    matched_keys: set[str] = set()
    for module in (_model, _shapes):
        for name, cls in sorted(vars(module).items()):
            if not inspect.isclass(cls) or cls.__module__ != module.__name__:
                continue
            if not issubclass(cls, VBAObject) or not cls._vba_members:
                continue
            type_name = cls.vba_type_name.lower()
            # Some coclasses have no members; their interface carries them.
            # Keep the exact keys in the report so this mapping is auditable.
            candidates = [
                f"excel:{type_name}",
                f"excel:i{type_name}",
                f"excel:_{type_name}",
            ]
            keys = [key for key in candidates if key in MEMBERS]
            expected: set[str] = set()
            for key in keys:
                expected.update(MEMBERS.get(key) or ())
            matched_keys.update(keys)
            registered = set(cls._vba_members)
            unsupported_paths: set[str] = set()
            for member_name, spec in cls._vba_members.items():
                for function in (spec.getter, spec.setter, spec.ref_setter):
                    if function is None:
                        continue
                    source = inspect.getsource(function)
                    if "VBAUnsupportedError" in source or "_unsupported(" in source:
                        unsupported_paths.add(member_name)
            bindings.append(
                {
                    "python_class": f"{module.__name__}.{name}",
                    "vba_type": cls.vba_type_name,
                    "reference_keys": keys,
                    "reference_name_count": len(expected),
                    "registered_names": sorted(registered),
                    "registered_reference_names": sorted(registered & expected),
                    "missing_reference_names": sorted(expected - registered),
                    "registered_without_reference": sorted(registered - expected),
                    "explicit_unsupported_paths": sorted(unsupported_paths),
                }
            )

    m_names = known()
    registered_m = {name.lower() for name in LIBRARY}
    return {
        "schema_version": 1,
        "meaning": "Name inventory only; registration does not prove implementation or conformance.",
        "limitations": [
            "Reference aliases, event interfaces and UI types are not distinct headless features.",
            "Missing reference data is unknown coverage, never a pass.",
            "Unsupported-path detection inspects direct source only; delegated gaps may remain.",
            "M reference names are normalized to lowercase by the existing inventory.",
            "Context-provided M bindings such as Excel.CurrentWorkbook are separate from LIBRARY.",
        ],
        "object_bindings": bindings,
        "unmapped_reference_types": {
            key: sorted(names)
            for key, names in sorted(MEMBERS.items())
            if key.startswith("excel:") and key not in matched_keys
        },
        "formula": {
            "registered_names": sorted(known_names()),
            "missing_reference_names": unimplemented(),
            "worksheet_function_bindings": sorted(_model.WorksheetFunction._vba_members),
        },
        "m_library": {
            "registered_names": sorted(LIBRARY),
            "reference_name_count": len(m_names),
            "registered_reference_names": sorted(m_names & registered_m),
            "missing_reference_names": sorted(m_names - registered_m),
            "registered_without_reference": sorted(registered_m - m_names),
        },
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, help="Write UTF-8 JSON to this path; otherwise use stdout.")
    args = parser.parse_args()
    text = json.dumps(report(), indent=2, ensure_ascii=True) + "\n"
    if args.output is None:
        sys.stdout.write(text)
    else:
        args.output.write_text(text, encoding="utf-8")


if __name__ == "__main__":
    main()
