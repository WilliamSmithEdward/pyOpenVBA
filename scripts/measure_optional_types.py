"""Optional parameters of every declared kind in native VBA: which compile, and what a missing one holds.

Each case declares Watch with its parameters and calls it from Probe; the
project is compiled, a compile error's message kept, and otherwise Probe
is run and its answer kept: the parameter's TypeName, IsMissing, and its
value or whether it is Nothing.

    python scripts/measure_optional_types.py

writes tests/fixtures/vba_semantics/optional_types.json, which
tests/test_vba_optional_types.py replays.
"""

from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/vba_semantics/optional_types.json"
VALUE = 'TypeName(x) & ":" & IsMissing(x) & ":" & x'
OBJECT = 'TypeName(x) & ":" & IsMissing(x) & ":" & (x Is Nothing)'
SHAPE = 'TypeName(x) & ":" & IsMissing(x)'

#: name -> (lines before Watch, Watch's parameters, what Watch answers, the call Probe makes).
CASES: dict[str, tuple[str, str, str, str]] = {
    "long": ("", "Optional x As Long", VALUE, "Watch()"),
    "integer": ("", "Optional x As Integer", VALUE, "Watch()"),
    "byte": ("", "Optional x As Byte", VALUE, "Watch()"),
    "boolean": ("", "Optional x As Boolean", VALUE, "Watch()"),
    "single": ("", "Optional x As Single", VALUE, "Watch()"),
    "double": ("", "Optional x As Double", VALUE, "Watch()"),
    "currency": ("", "Optional x As Currency", VALUE, "Watch()"),
    "date": ("", "Optional x As Date", VALUE, "Watch()"),
    "string": ("", "Optional x As String", VALUE, "Watch()"),
    "long_long": ("", "Optional x As LongLong", VALUE, "Watch()"),
    "long_ptr": ("", "Optional x As LongPtr", VALUE, "Watch()"),
    "variant": ("", "Optional x As Variant", SHAPE, "Watch()"),
    "untyped": ("", "Optional x", SHAPE, "Watch()"),
    "object": ("", "Optional x As Object", OBJECT, "Watch()"),
    "collection": ("", "Optional x As Collection", OBJECT, "Watch()"),
    "enum": ("Enum Shade\n    Pale = 1\nEnd Enum\n", "Optional x As Shade", VALUE, "Watch()"),
    "integer_character": ("", "Optional x%", VALUE, "Watch()"),
    "string_character": ("", "Optional x$", VALUE, "Watch()"),
    "by_value": ("", "Optional ByVal x As Long", VALUE, "Watch()"),
    "by_reference": ("", "Optional ByRef x As Long", VALUE, "Watch()"),
    "passed": ("", "Optional x As Long", VALUE, "Watch(5)"),
    "skipped": ("", "a As Long, Optional x As Long, Optional c As String", VALUE + ' & ":" & c', 'Watch(1, , "z")'),
    "named": ("", "Optional a As Long, Optional x As Long", VALUE + ' & ":" & a', "Watch(x:=3)"),
    "deftype": ("DefLng X\n", "Optional xval", 'TypeName(xval) & ":" & xval', "Watch()"),
    "deftype_with_as": ("DefLng X\n", "Optional xval As Long", 'TypeName(xval) & ":" & xval', "Watch()"),
    "deftype_with_default": ("DefLng X\n", "Optional xval = 4", 'TypeName(xval) & ":" & xval', "Watch()"),
    "long_array": ("", "Optional x() As Long", SHAPE, "Watch()"),
    "variant_array": ("", "Optional x()", SHAPE, "Watch()"),
    "fixed_string": ("", "Optional x As String * 3", VALUE, "Watch()"),
}


def source(before: str, parameters: str, answer: str, call: str) -> str:
    return (f"{before}Function Watch({parameters}) As String\n    Watch = {answer}\nEnd Function\n"
            f"Function Probe() As String\n    Probe = {call}\nEnd Function\n")


def main() -> None:
    probes = []
    with ExcelSession() as excel:
        excel.new_document()
        version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & '
                                'Application.Build\nEnd Function', "VersionBuild", line_numbers=False)
        assert version.ok, version
        for name, parts in CASES.items():
            excel.new_document()
            text = source(*parts)
            excel.add_module("Harness", text, line_numbers=False)
            compiled = excel.compile_project(watch_seconds=20)
            assert compiled.outcome in ("accepted", "rejected"), (name, compiled)
            case: dict[str, object] = {"name": name, "source": text}
            if not compiled.ok:
                assert compiled.dialog is not None, (name, compiled)
                case["compile_error"] = True
                case["compiler_message"] = compiled.dialog.message
            else:
                result = excel.run_macro("Harness.Probe", timeout=20)
                assert result.ok, (name, result)
                case["result"] = result.value
            probes.append(case)
            print(name, case.get("result", case.get("compiler_message")), flush=True)
    record = {"excel_version_build": version.value, "measured_at": datetime.now(timezone.utc).isoformat(),
              "probes": probes}
    OUT.write_text(json.dumps(record, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
