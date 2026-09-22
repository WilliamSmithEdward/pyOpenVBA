"""Build docs/excel_checklist.csv: one row per member of every Excel class the model implements.

Reads the type-library dumps in the pyVBAReference repository, the same
source scripts/build_object_inventory.py uses, and the members the model
registers. The object, member, kind, signature and interface_gaps columns
are generated. The status, evidence and note columns are kept from the
file already there, so rerunning this after a change only adds rows and
reconciles statuses the registry contradicts.

Usage (the default path is the sibling checkout):

    python scripts/build_excel_checklist.py [path\\to\\pyVBAReference]

Statuses are defined in docs/host_completeness.md.
"""

from __future__ import annotations

import csv
import inspect
import json
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

DEFAULT_REFERENCE = ROOT.parent / "pyVBAReference"
CHECKLIST = ROOT / "docs" / "excel_checklist.csv"
COLUMNS = ("object", "member", "kind", "signature", "interface_gaps", "status", "evidence", "note")
IMPLEMENTED = ("VERIFIED", "PARTIAL", "UNASSESSED")
NOT_IMPLEMENTED = ("MISSING", "EXCLUDED")


def bound_classes() -> dict[str, type]:
    """The model class behind each Excel type, one per type name."""
    from pyopenvba.apps.excel import _model, _shapes
    from pyopenvba.interpreter._objects import VBAObject

    found: dict[str, type] = {}
    for module in (_model, _shapes):
        for cls in vars(module).values():
            if not inspect.isclass(cls) or cls.__module__ != module.__name__:
                continue
            if not issubclass(cls, VBAObject) or not cls._vba_members:
                continue
            # Rows and Columns views share Range's type name; the class
            # with the most members is the one that stands for the type.
            current = found.get(cls.vba_type_name)
            if current is None or len(cls._vba_members) > len(current._vba_members):  # type: ignore[attr-defined]
                found[cls.vba_type_name] = cls
    return dict(sorted(found.items()))


def reference_members(reference: Path, type_name: str) -> list[dict[str, object]]:
    path = reference / "reference" / "excel" / "json" / f"{type_name}.json"
    if not path.is_file():
        return []
    data = json.loads(path.read_text(encoding="utf-8"))
    return [*data.get("properties", []), *data.get("methods", []), *data.get("events", [])]


def describe(entry: dict[str, object]) -> tuple[str, str]:
    kind = str(entry.get("kind", ""))
    if kind == "property":
        access = str(entry.get("access", ""))
        return ("property, read-only" if access == "read-only" else "property"), f"As {entry.get('type', 'Variant')}"
    return kind, str(entry.get("signature", ""))


def interface_gaps(entry: dict[str, object] | None, spec: object | None) -> str:
    """What the registration visibly lacks next to the type library."""
    if spec is None or entry is None:
        return ""
    gaps: list[str] = []
    kind = str(entry.get("kind", ""))
    if kind == "property" and entry.get("access") != "read-only" and not getattr(spec, "settable", False):
        gaps.append("no setter")
    wanted = [str(one.get("name", "")) for one in entry.get("parameters", []) or []]  # type: ignore[union-attr]
    if wanted and not getattr(spec, "varargs", False):
        have = {name.lower() for name in getattr(spec, "parameters", ())}
        absent = [name for name in wanted if name.lower() not in have]
        if absent:
            gaps.append("no " + ", ".join(absent))
    return "; ".join(gaps)


def existing_rows() -> dict[tuple[str, str], dict[str, str]]:
    if not CHECKLIST.is_file():
        return {}
    with CHECKLIST.open(encoding="utf-8", newline="") as handle:
        return {(row["object"], row["member"].lower()): row for row in csv.DictReader(handle)}


def build(reference: Path) -> list[dict[str, str]]:
    kept = existing_rows()
    rows: list[dict[str, str]] = []
    for type_name, cls in bound_classes().items():
        registered = cls._vba_members  # type: ignore[attr-defined]
        entries = {str(entry["name"]).lower(): entry for entry in reference_members(reference, type_name)}
        for key in sorted(set(entries) | set(registered)):
            entry = entries.get(key)
            spec = registered.get(key)
            if entry is not None:
                name = str(entry["name"])
                kind, signature = describe(entry)
            else:
                name, kind, signature = spec.name, "not in the type library", ""
            previous = kept.get((type_name, key), {})
            status = previous.get("status", "")
            if spec is not None and status not in IMPLEMENTED:
                if status:
                    print(f"note: {type_name}.{name} is registered; {status} becomes UNASSESSED")
                status = "UNASSESSED"
            elif spec is None and status not in NOT_IMPLEMENTED:
                if status:
                    print(f"note: {type_name}.{name} is not registered; {status} becomes MISSING")
                status = "MISSING"
            rows.append({
                "object": type_name,
                "member": name,
                "kind": kind,
                "signature": signature,
                "interface_gaps": interface_gaps(entry, spec),
                "status": status,
                "evidence": previous.get("evidence", ""),
                "note": previous.get("note", ""),
            })
    return rows


def main() -> None:
    reference = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_REFERENCE
    if not (reference / "reference" / "excel" / "json").is_dir():
        raise SystemExit(f"no Excel type-library dump under {reference}")
    rows = build(reference)
    with CHECKLIST.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=COLUMNS, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)
    counts = Counter((row["object"], row["status"]) for row in rows)
    statuses = ("VERIFIED", "PARTIAL", "UNASSESSED", "MISSING", "EXCLUDED")
    print(f"{len(rows)} rows written to {CHECKLIST.relative_to(ROOT)}")
    print(f"{'object':<20}" + "".join(f"{status:>11}" for status in statuses))
    for type_name in sorted({row["object"] for row in rows}):
        print(f"{type_name:<20}" + "".join(f"{counts[(type_name, status)]:>11}" for status in statuses))


if __name__ == "__main__":
    main()
