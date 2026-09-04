"""Name codes 700-703 on a command button by differencing: build the same
form twice, identical but for one property, and see which record moved.
The candidates are the themed hover/pressed colours Access 2010 added to
buttons, which are the only Long-typed button properties still unnamed."""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "src"))

import pyvbaharness  # noqa: E402

from pyopenvba.access import AccessDatabase  # noqa: E402
from pyopenvba.access._designs import PROPERTY_CODES  # noqa: E402

HERE = Path(".").resolve()
TEMPLATE = Path(__file__).resolve().parents[3] / "src" / "pyopenvba" / "_templates" / "blank_files" / "blank_database.accdb"
NAMES = {c: n for n, c in PROPERTY_CODES.items()}

COMMAND_BUTTON = 104
TOGGLE_BUTTON = 122

#: (property, control type, low value, high value)
TRIALS = [
    ("HoverColor", COMMAND_BUTTON, 0, 12611584),
    ("PressedColor", COMMAND_BUTTON, 0, 8421504),
    ("HoverForeColor", COMMAND_BUTTON, 0, 255),
    ("PressedForeColor", COMMAND_BUTTON, 0, 65280),
    ("HoverThemeColorIndex", COMMAND_BUTTON, 4, 6),
    ("PressedThemeColorIndex", COMMAND_BUTTON, 4, 6),
    ("HoverForeThemeColorIndex", COMMAND_BUTTON, 0, 2),
    ("PressedForeThemeColorIndex", COMMAND_BUTTON, 0, 2),
    ("HoverShade", COMMAND_BUTTON, 100, 60),
    ("PressedShade", COMMAND_BUTTON, 100, 60),
    ("HoverTint", COMMAND_BUTTON, 100, 60),
    ("PressedTint", COMMAND_BUTTON, 100, 60),
    ("HoverForeShade", COMMAND_BUTTON, 100, 60),
    ("PressedForeShade", COMMAND_BUTTON, 100, 60),
    ("HoverForeTint", COMMAND_BUTTON, 100, 60),
    ("PressedForeTint", COMMAND_BUTTON, 100, 60),
    ("Gradient", COMMAND_BUTTON, 0, 3),
    ("Bevel", COMMAND_BUTTON, 0, 3),
    ("Shape", COMMAND_BUTTON, 0, 3),
    ("Glow", COMMAND_BUTTON, 0, 3),
    ("Shadow", COMMAND_BUTTON, 0, 3),
    ("QuickStyle", COMMAND_BUTTON, 0, 3),
    ("QuickStyleMask", COMMAND_BUTTON, 0, 3),
    ("CursorOnHover", COMMAND_BUTTON, 0, 1),
    ("PictureCaptionArrangement", COMMAND_BUTTON, 0, 2),
    ("UseTheme", COMMAND_BUTTON, True, False),
    ("HoverColor", TOGGLE_BUTTON, 0, 12611584),
    ("PressedColor", TOGGLE_BUTTON, 0, 8421504),
    ("HoverForeColor", TOGGLE_BUTTON, 0, 255),
    ("PressedForeColor", TOGGLE_BUTTON, 0, 65280),
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
                result = access.run_vba(VBA, proc="Build", args=(prop, str(kind), str(value)), timeout=180.0)
                if result.outcome != "passed" or "ERR" in str(result.value):
                    sides = []
                    found[f"{prop}/{kind}"] = f"not settable: {result.value}"
                    break
                sides.append(records(out, str(result.value).split("|")[1]))
            if len(sides) != 2:
                continue
            first, second = sides
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
                found[f"{prop}/{kind}"] = f"id {ident} code {code} type {slot[1]} width {slot[2]} value {slot[3].hex()}{mark}"
            else:
                detail = ", ".join(f"id {i} code {(second.get(i) or first[i])[0]}" for i in sorted(changed))
                found[f"{prop}/{kind}"] = f"{len(changed)} moved: {detail}"
    for key, value in found.items():
        print(f"  {key:34} {value}")


if __name__ == "__main__":
    main()
