"""Measure native VBA Implements: calls through an interface variable, identity,
TypeName and TypeOf, late binding, casting back, and parameters and returns typed
as the interface.

Each probe is a module whose ``Function Probe() As String`` the fixture records as a
value or a run-time error number; the class modules are added to every session
(see vba_probe_fixture.py).
"""
from __future__ import annotations
from pathlib import Path
from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_interfaces.json'
CLASSES = {
    "IShape": '''Public Name As String
Public Function Area() As Double
End Function
Public Sub Grow(ByVal f As Double)
End Sub
''',
    "CSquare": '''Implements IShape
Public Side As Double
Private mName As String
Private Function IShape_Area() As Double
IShape_Area = Side * Side
End Function
Private Sub IShape_Grow(ByVal f As Double)
Side = Side * f
End Sub
Private Property Get IShape_Name() As String
IShape_Name = "sq:" & mName
End Property
Private Property Let IShape_Name(ByVal v As String)
mName = v
End Property
Public Function Area() As String
Area = "own area"
End Function
''',
    "CPlain": '''Public Side As Double
''',
}
HEAD = r'''
Private Function Describe(ByVal sh As IShape) As String
Describe = TypeName(sh) & ":" & sh.Area
End Function
Private Function AsShape(ByVal sq As CSquare) As IShape
Set AsShape = sq
End Function
'''
CASES = {
    "call_through_interface": 'Dim s As New CSquare, sh As IShape\ns.Side = 3\nSet sh = s\nProbe = sh.Area & "|" & s.Area',
    "is_identity": 'Dim s As New CSquare, sh As IShape\nSet sh = s\nProbe = (sh Is s) & "|" & (s Is sh)',
    "typename_typeof": 'Dim s As New CSquare, sh As IShape\nSet sh = s\nProbe = TypeName(sh) & "|" & (TypeOf sh Is CSquare) & "|" & (TypeOf s Is IShape) & "|" & (TypeOf sh Is IShape)',
    "typeof_not_implemented": 'Dim p As New CPlain\nProbe = (TypeOf p Is IShape)',
    "interface_property": 'Dim s As New CSquare, sh As IShape\nSet sh = s\nsh.Name = "a"\nProbe = sh.Name',
    "interface_sub": 'Dim s As New CSquare, sh As IShape\ns.Side = 2\nSet sh = s\nsh.Grow 3\nProbe = s.Side',
    "parameter_interface": 'Dim s As New CSquare\ns.Side = 4\nProbe = Describe(s)',
    "return_interface": 'Dim s As New CSquare\ns.Side = 5\nProbe = AsShape(s).Area',
    "cast_back": 'Dim s As New CSquare, sh As IShape, s2 As CSquare\ns.Side = 6\nSet sh = s\nSet s2 = sh\nProbe = s2.Side & "|" & s2.Area',
    "object_keeps_interface": 'Dim s As New CSquare, sh As IShape, o As Object\ns.Side = 7\nSet sh = s\nSet o = sh\nProbe = o.Area',
    "object_class_member_missing": 'Dim s As New CSquare, sh As IShape, o As Object\ns.Side = 7\nSet sh = s\nSet o = sh\nProbe = o.Side',
    "variant_keeps_interface": 'Dim s As New CSquare, sh As IShape, v As Variant\ns.Side = 7\nSet sh = s\nSet v = sh\nProbe = v.Area',
    "not_implemented_mismatch": 'Dim p As New CPlain, sh As IShape\nOn Error Resume Next\nSet sh = p\nProbe = Err.Number & "|" & (sh Is Nothing)',
    "for_each_interface": 'Dim c As New Collection, sh As IShape, s As CSquare, t As Double\nSet s = New CSquare: s.Side = 1: c.Add s\nSet s = New CSquare: s.Side = 2: c.Add s\nFor Each sh In c\nt = t + sh.Area\nNext\nProbe = t',
    "collection_item_interface": 'Dim c As New Collection, s As New CSquare, sh As IShape\ns.Side = 3\nc.Add s\nSet sh = c(1)\nProbe = sh.Area',
    "interface_class_itself": 'Dim i As New IShape\ni.Name = "x"\nProbe = i.Name & "|" & i.Area',
    "nothing_interface_call": 'Dim sh As IShape\nOn Error Resume Next\nDim a As Double\na = sh.Area\nProbe = Err.Number',
    "callbyname_interface": 'Dim s As New CSquare, sh As IShape\ns.Side = 3\nSet sh = s\nProbe = CallByName(sh, "Area", VbMethod)',
}


def macro(name: str) -> str:
    return HEAD + '\nFunction Probe() As String\n' + CASES[name] + '\nEnd Function\n'


def main() -> None:
    record({name: macro(name) for name in CASES}, OUT, classes=CLASSES)


if __name__ == '__main__':
    main()
