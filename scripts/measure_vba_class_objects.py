"""Measure everyday class-module and object idioms: properties, events, CallByName,
object arrays, arguments a member takes none of, Err across a member call, a Let
into an object variable and Is on values that are not objects.

Each probe is a module whose ``Function Probe() As String`` the fixture records as
a value or a run-time error number (see vba_probe_fixture.py); the class modules
are added to every session.
"""
from __future__ import annotations
from pathlib import Path
from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_class_objects.json'

CLASSES = {
    "CPoint": '''Option Explicit
Private mX As Double
Private mY As Double
Public Tag As String
Public Items As Collection
Public Event Moved(ByVal dx As Double)

Public Property Get X() As Double
    X = mX
End Property
Public Property Let X(ByVal v As Double)
    If v < 0 Then Err.Raise vbObjectError + 1, "CPoint", "negative"
    mX = v
End Property
Public Property Get Y() As Double
    Y = mY
End Property
Public Property Let Y(ByVal v As Double)
    mY = v
End Property
Public Function Length() As Double
    Length = Sqr(mX * mX + mY * mY)
End Function
Public Sub Move(ByVal dx As Double, Optional ByVal dy As Double = 0)
    mX = mX + dx
    mY = mY + dy
    RaiseEvent Moved(dx)
End Sub
Public Function Clone() As CPoint
    Dim p As New CPoint
    p.X = mX: p.Y = mY
    Set Clone = p
End Function
Public Property Get Self() As CPoint
    Set Self = Me
End Property
Private Sub Class_Initialize()
    Tag = "init"
End Sub
''',
    "CBag": '''Option Explicit
Private mItems As Collection
Private Sub Class_Initialize()
    Set mItems = New Collection
End Sub
Public Sub Add(ByVal v As Variant)
    mItems.Add v
End Sub
Public Property Get Count() As Long
    Count = mItems.Count
End Property
Public Property Get Item(ByVal i As Long) As Variant
    Item = mItems(i)
End Property
Public Property Let Item(ByVal i As Long, ByVal v As Variant)
    mItems.Remove i
    If i > mItems.Count Then mItems.Add v Else mItems.Add v, , i
End Property
Public Property Get All() As Collection
    Set All = mItems
End Property
''',
    "CWatcher": '''Option Explicit
Public WithEvents P As CPoint
Public Log As String
Private Sub P_Moved(ByVal dx As Double)
    Log = Log & "moved " & dx & ";"
End Sub
''',
    "CErrProbe": '''Option Explicit
Public Field As Long
Public Property Get Value() As Long
    Value = Err.Number
End Property
Public Function Peek() As String
    Peek = "inside " & Err.Number
End Function
Public Sub Touch()
    Field = Err.Number
End Sub
Public Function Handled() As String
    On Error Resume Next
    Handled = "inside " & Err.Number
End Function
Public Function Raises() As String
    On Error Resume Next
    Err.Raise 1004
    Raises = "inside " & Err.Number
End Function
''',
}

HELPERS = '''
Private Function Plain() As String
    Plain = "inside " & Err.Number
End Function
'''
RAISED = 'On Error Resume Next\nErr.Raise 7\n'

CASES = {
    "class_props": 'Dim p As New CPoint\np.X = 3: p.Y = 4\nProbe = p.Length & "|" & p.Tag & "|" & TypeName(p) & "|" & (TypeOf p Is CPoint)',
    "class_prop_error": 'Dim p As New CPoint\nOn Error Resume Next\np.X = -1\nProbe = Err.Number - vbObjectError & "|" & Err.Source & "|" & Err.Description & "|" & p.X',
    "class_clone_identity": 'Dim p As New CPoint, q As CPoint\np.X = 1\nSet q = p.Clone\nq.X = 5\nProbe = p.X & "|" & q.X & "|" & (p Is q) & "|" & (p Is p.Self)',
    "class_nothing_errors": 'Dim p As CPoint\nOn Error Resume Next\nDim x As Double\nx = p.X\nProbe = Err.Number & "|" & (p Is Nothing) & "|" & TypeName(p)',
    "class_param_property": 'Dim b As New CBag\nb.Add "a": b.Add "b": b.Add "c"\nb.Item(2) = "z"\nProbe = b.Count & "|" & b.Item(1) & b.Item(2) & b.Item(3)',
    "class_property_returns_collection": 'Dim b As New CBag, v, s As String\nb.Add 1: b.Add 2\nFor Each v In b.All: s = s & v: Next\nb.All.Add 3\nProbe = s & "|" & b.Count & "|" & b.All(3)',
    "class_with_block": 'Dim p As New CPoint\nWith p\n.X = 6\n.Y = 8\nProbe = .Length & "|" & .Tag\nEnd With',
    "class_events": 'Dim w As New CWatcher\nSet w.P = New CPoint\nw.P.Move 2\nw.P.Move 3, 1\nProbe = w.Log & "|" & w.P.X & "|" & w.P.Y',
    "class_events_unhooked": 'Dim w As New CWatcher, p As New CPoint\nSet w.P = p\np.Move 1\nSet w.P = Nothing\np.Move 1\nProbe = w.Log & "|" & p.X',
    "class_public_field_object": 'Dim p As New CPoint\nSet p.Items = New Collection\np.Items.Add "q"\nProbe = p.Items.Count & "|" & p.Items(1)',
    "class_callbyname": 'Dim p As New CPoint\nCallByName p, "X", VbLet, 7\nCallByName p, "Move", VbMethod, 1, 2\nProbe = CallByName(p, "X", VbGet) & "|" & CallByName(p, "Y", VbGet) & "|" & CallByName(p, "Length", VbMethod)',
    "class_callbyname_missing": 'Dim p As New CPoint\nOn Error Resume Next\nDim v\nv = CallByName(p, "Nope", VbGet)\nProbe = Err.Number',
    "class_array_of_objects": 'Dim a(1 To 3) As CPoint, i As Long, s As Double\nFor i = 1 To 3\nSet a(i) = New CPoint\na(i).X = i\nNext\nFor i = 1 To 3: s = s + a(i).X: Next\nProbe = s & "|" & TypeName(a) & "|" & TypeName(a(1))',
    "class_in_collection_foreach": 'Dim c As New Collection, p As CPoint, i As Long, s As String\nFor i = 1 To 3\nSet p = New CPoint\np.X = i\nc.Add p\nNext\nFor Each p In c: s = s & p.X: Next\nProbe = s',
    "class_object_compare_eq": 'Dim p As New CPoint\nOn Error Resume Next\nDim b As Boolean\nb = (p = p)\nProbe = Err.Number',
    "class_variant_holding_object": 'Dim v As Variant\nSet v = New CPoint\nv.X = 2\nProbe = TypeName(v) & "|" & IsObject(v) & "|" & v.X & "|" & VarType(v)',
    "class_method_optional": 'Dim p As New CPoint\np.Move 1\np.Move dy:=5, dx:=2\nProbe = p.X & "|" & p.Y',
    "as_new_reinit": 'Dim p As New CPoint\np.X = 5\nSet p = Nothing\nProbe = p.X & "|" & (p Is Nothing)',
    "typeof_checks": 'Dim o As Object, c As New Collection\nSet o = c\nProbe = (TypeOf o Is Collection) & "|" & (TypeOf o Is CPoint) & "|" & (TypeOf o Is Object)',
    "object_variable_late_bound": 'Dim o As Object\nSet o = New CPoint\no.X = 3\nOn Error Resume Next\no.Nope = 1\nProbe = o.X & "|" & Err.Number',
    "nested_with": 'Dim w As New CWatcher\nSet w.P = New CPoint\nWith w\nWith .P\n.X = 9\nEnd With\nProbe = .P.X\nEnd With',
    "with_new_object": 'Dim s As String\nWith New CPoint\n.X = 4\ns = .Tag & .X\nEnd With\nProbe = s',
    "err_early_property_get": 'Dim c As New CErrProbe\n' + RAISED + 'Probe = c.Value & " / after " & Err.Number',
    "err_early_function": 'Dim c As New CErrProbe\n' + RAISED + 'Probe = c.Peek() & " / after " & Err.Number',
    "err_early_sub": 'Dim c As New CErrProbe\n' + RAISED + 'c.Touch\nProbe = c.Field & " / after " & Err.Number',
    "err_early_handled": 'Dim c As New CErrProbe\n' + RAISED + 'Probe = c.Handled() & " / after " & Err.Number',
    "err_early_raises": 'Dim c As New CErrProbe\n' + RAISED + 'Probe = c.Raises() & " / after " & Err.Number',
    "err_late_property_get": 'Dim o As Object\nSet o = New CErrProbe\n' + RAISED + 'Probe = o.Value & " / after " & Err.Number',
    "err_late_function": 'Dim o As Object\nSet o = New CErrProbe\n' + RAISED + 'Probe = o.Peek() & " / after " & Err.Number',
    "err_variant_function": 'Dim v As Variant\nSet v = New CErrProbe\n' + RAISED + 'Probe = v.Peek() & " / after " & Err.Number',
    "err_plain_call": RAISED + 'Probe = Plain() & " / after " & Err.Number',
    "err_as_new_first_use": RAISED + 'Dim c As New CErrProbe\nProbe = c.Peek() & " / after " & Err.Number',
    "err_callbyname_member": 'Dim c As New CErrProbe\n' + RAISED + 'Probe = CallByName(c, "Peek", VbMethod) & " / after " & Err.Number',
    "typename_object_arrays": 'Dim a(1) As CPoint, b() As CPoint, c(1) As Collection, d(1) As Object, e(1) As Worksheet, f(1) As Range\nReDim b(2)\nProbe = TypeName(a) & "|" & TypeName(b) & "|" & TypeName(c) & "|" & TypeName(d) & "|" & TypeName(e) & "|" & TypeName(f)',
    "typename_object_array_in_variant": 'Dim a(1) As CPoint, v As Variant\nv = a\nProbe = TypeName(v) & "|" & VarType(a) & "|" & VarType(v)',
    "let_into_nothing_class": 'Dim p As New CPoint, q As CPoint\nOn Error Resume Next\nq = p\nProbe = Err.Number',
    "let_into_set_class": 'Dim p As New CPoint, q As CPoint\nSet q = New CPoint\nOn Error Resume Next\nq = p\nProbe = Err.Number',
    "let_number_into_object": 'Dim o As Object\nOn Error Resume Next\no = 5\nProbe = Err.Number',
    "is_on_values": 'Dim v As Variant, r As Boolean, s As String\nOn Error Resume Next\nr = v Is Nothing\ns = Err.Number\nErr.Clear\nv = Null\nr = v Is Nothing\ns = s & "|" & Err.Number\nErr.Clear\nv = 5\nr = v Is Nothing\ns = s & "|" & Err.Number\nErr.Clear\nSet v = Nothing\nr = v Is Nothing\ns = s & "|" & Err.Number & r\nProbe = s',
}


def macro(name: str) -> str:
    return HELPERS + '\nFunction Probe() As String\n' + CASES[name] + '\nEnd Function\n'


def main() -> None:
    record({name: macro(name) for name in CASES}, OUT, CLASSES)


if __name__ == '__main__':
    main()
