"""Name a property by changing only it.

Matching on value cannot name a property whose values are small integers
every other property also uses.  Building the same form twice, identical
but for one property, and differencing the records of the same control
can: whatever changed is that property.
"""

import shutil
import pathlib
import sys
from pathlib import Path

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[3] / "src"))

import pyvbaharness  # noqa: E402

from pyopenvba.access import AccessDatabase  # noqa: E402
from pyopenvba.access._designs import PROPERTY_CODES  # noqa: E402

HERE = Path(".").resolve()
TEMPLATE = (
    pathlib.Path(__file__).resolve().parents[3]
    / "src"
    / "pyopenvba"
    / "_templates"
    / "blank_files"
    / "blank_database.accdb"
)
NAMES = {c: n for n, c in PROPERTY_CODES.items()}

#: (property, control type code, the two values to compare).
TRIALS = [
    ("FilterLookup", 109, 1, 0),
    ("AllowAutoCorrect", 109, True, False),
    ("VerticalAnchor", 109, 0, 1),
    ("HorizontalAnchor", 109, 0, 1),
    ("NumeralShapes", 109, 0, 1),
    ("ReadingOrder", 109, 0, 1),
    ("KeyboardLanguage", 109, 0, 1),
    ("ScrollBarAlign", 109, 0, 1),
    ("AsianLineBreak", 109, True, False),
    ("IMEMode", 109, 0, 1),
    ("IMESentenceMode", 109, 3, 0),
    ("TextFormat", 109, 0, 1),
    ("DisplayAsHyperlink", 109, 0, 1),
    ("AllowDatasheetCaption", 109, True, False),
    ("ShowDatePicker", 109, 1, 0),
    ("IsHyperlink", 109, False, True),
    ("SmartTags", 109, "", "x"),
    ("LabelAlign", 109, 0, 1),
    ("AddColon", 109, True, False),
    ("AutoLabel", 109, True, False),
    ("LabelX", 109, 0, 53),
    ("LabelY", 109, 0, 59),
]

VBA = """
Public Function Build(ByVal prop As String, ByVal kind As Long, ByVal raw As String) As String
    Dim f As Object, c As Object
    Set f = CreateForm()
    Set c = CreateControl(f.Name, kind, 0, "", "", 211, 307, 1409, 313)
    c.Name = "Probe"
    On Error Resume Next
    Err.Clear
    If raw = "True" Then
        CallByName c, prop, VbLet, True
    ElseIf raw = "False" Then
        CallByName c, prop, VbLet, False
    ElseIf IsNumeric(raw) Then
        CallByName c, prop, VbLet, CLng(raw)
    Else
        CallByName c, prop, VbLet, raw
    End If
    Build = IIf(Err.Number = 0, "ok", "ERR " & Err.Description)
    On Error GoTo 0
    DoCmd.Save acForm, f.Name
    Build = Build & "|" & f.Name
    DoCmd.Close acForm, f.Name, acSaveYes
End Function
"""


def records(path: Path, form: str) -> dict[int, tuple[int, int, int, bytes]]:
    db = AccessDatabase(path)
    for obj in db.form(form).objects:
        if obj.name == "Probe":
            return {r.id: (r.code, r.value_type, r.width, r.value) for r in obj.records}
    return {}


def main() -> None:
    found: dict[str, str] = {}
    with pyvbaharness.AccessSession() as access:
        for prop, kind, low, high in TRIALS:
            sides: list[dict[int, tuple[int, int, int, bytes]]] = []
            for tag, value in (("a", low), ("b", high)):
                out = HERE / f"d_{prop}_{kind}_{tag}.accdb"
                shutil.copy(TEMPLATE, out)
                access.open_document(out, read_only=False)
                result = access.run_vba(
                    VBA, proc="Build", args=(prop, str(kind), str(value)), timeout=180.0
                )
                if result.outcome != "passed" or "ERR" in str(result.value):
                    sides = []
                    found[f"{prop}/{kind}"] = f"not settable: {result.value}"
                    break
                sides.append(records(out, str(result.value).split("|")[1]))
            if len(sides) != 2:
                continue
            first, second = sides
            # Access mints a new GUID each time it builds the control,
            # so that record always differs and says nothing.
            changed = [
                ident
                for ident in set(first) | set(second)
                if first.get(ident) != second.get(ident)
                and (second.get(ident) or first[ident])[0] != PROPERTY_CODES["GUID"]
            ]
            if len(changed) == 1:
                ident = changed[0]
                slot = second.get(ident) or first[ident]
                code = slot[0]
                known = NAMES.get(code)
                mark = "" if known is None else f"  (we call {code} {known})"
                found[f"{prop}/{kind}"] = (
                    f"id {ident} code {code} type {slot[1]} width {slot[2]}{mark}"
                )
            else:
                detail = ", ".join(
                    f"id {i} code {(second.get(i) or first[i])[0]}" for i in sorted(changed)
                )
                found[f"{prop}/{kind}"] = f"{len(changed)} moved: {detail}"
    for key, value in found.items():
        print(f"  {key:24} {value}")


if __name__ == "__main__":
    main()
