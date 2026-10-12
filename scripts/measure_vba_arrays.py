"""Measure everyday array handling: unallocated arrays and the (Not a) idiom,
ReDim and ReDim Preserve edges, Erase, ByRef and ByVal array parameters,
arrays inside Variants and arrays, ParamArray, Split, Join and Filter, and
which arrays a typed dynamic array takes.

Each probe is a module whose ``Function Probe() As String`` the fixture
records as a value or a run-time error number (see vba_probe_fixture.py).
"""
from __future__ import annotations
from pathlib import Path
from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_arrays.json'

HELPERS = '''
Private Sub FillByRef(a() As Long)
    ReDim a(1 To 2)
    a(2) = 7
End Sub

Private Sub TouchVariant(ByVal v As Variant)
    v(0) = 99
End Sub

Private Sub TouchVariantRef(v As Variant)
    v(0) = 99
End Sub

Private Function CountDims(v As Variant) As Long
    Dim d As Long, n As Long
    On Error GoTo Done
    For d = 1 To 60
        n = UBound(v, d)
    Next
Done:
    CountDims = d - 1
End Function

Private Function SumAll(ParamArray xs() As Variant) As Double
    Dim x
    For Each x In xs
        SumAll = SumAll + x
    Next
End Function
'''

CASES = {
    'unallocated_checks': 'Dim a() As Long\nOn Error Resume Next\nDim n As Long\nn = UBound(a)\nProbe = Err.Number & "|" & IsArray(a) & "|" & (Not a) & "|" & TypeName(a)',
    'ubound_on_non_array': 'Dim v As Variant\nOn Error Resume Next\nDim n As Long\nn = UBound(v)\nProbe = Err.Number\nErr.Clear\nv = "abc"\nn = UBound(v)\nProbe = Probe & "|" & Err.Number',
    'redim_fixed_runtime': 'Dim v As Variant\nv = Array(1, 2)\nReDim Preserve v(0 To 3)\nProbe = UBound(v) & "|" & v(1) & "|" & IsEmpty(v(3))',
    'redim_preserve_lower_bound_change': 'Dim a() As Long\nReDim a(1 To 3)\nOn Error Resume Next\nReDim Preserve a(0 To 3)\nProbe = Err.Number & "|" & LBound(a)',
    'redim_preserve_2d_last_dim': 'Dim a() As Long\nReDim a(1 To 2, 1 To 2)\na(2, 2) = 5\nReDim Preserve a(1 To 2, 1 To 4)\nProbe = a(2, 2) & "|" & UBound(a, 2)\nOn Error Resume Next\nReDim Preserve a(1 To 3, 1 To 4)\nProbe = Probe & "|" & Err.Number',
    'redim_change_dims': 'Dim a() As Long\nReDim a(2)\nReDim a(1, 1)\nProbe = CountDims(a)\nOn Error Resume Next\nReDim Preserve a(1, 1, 1)\nProbe = Probe & "|" & Err.Number',
    'erase_dynamic_then_ubound': 'Dim a() As String\nReDim a(2)\nErase a\nOn Error Resume Next\nDim n As Long\nn = UBound(a)\nProbe = Err.Number & "|" & LBound(Split(""))',
    'erase_variant_array': 'Dim v As Variant\nv = Array(1, 2)\nOn Error Resume Next\nErase v\nProbe = Err.Number & "|" & TypeName(v) & "|" & IsArray(v)',
    'byref_array_param_redim': 'Dim a() As Long\nFillByRef a\nProbe = LBound(a) & UBound(a) & "|" & a(2)',
    'variant_array_byval_byref': 'Dim v As Variant\nv = Array(1, 2)\nTouchVariant v\nProbe = v(0)\nTouchVariantRef v\nProbe = Probe & "|" & v(0)',
    'count_dims': 'Dim a(1, 2, 3) As Byte, b\nb = Array(1)\nProbe = CountDims(a) & CountDims(b)',
    'array_in_array_assign': 'Dim outer(1) As Variant\nouter(0) = Array(1, 2)\nouter(0)(1) = 9\nProbe = outer(0)(1)',
    'fixed_array_assign_error': 'Dim a(2) As Long, b(2) As Long\nb(1) = 4\nOn Error Resume Next\nDim v As Variant\nv = b\nProbe = v(1) & "|" & TypeName(v)',
    'multidim_iteration_order': 'Dim a(1 To 2, 1 To 2) As Long, x, s As String\na(1, 1) = 1: a(2, 1) = 2: a(1, 2) = 3: a(2, 2) = 4\nFor Each x In a: s = s & x: Next\nProbe = s',
    'lbound_ubound_dims': 'Dim a(-2 To 2, 5 To 7) As Long\nOn Error Resume Next\nProbe = LBound(a) & UBound(a) & LBound(a, 2) & UBound(a, 2)\nDim n As Long\nn = UBound(a, 3)\nProbe = Probe & "|" & Err.Number',
    'paramarray_sum': 'Probe = SumAll(1, 2, 3) & "|" & SumAll() & "|" & SumAll(1.5, "2")',
    'array_function_types': 'Dim v\nv = Array()\nProbe = UBound(v) & "|" & TypeName(v) & "|" & TypeName(Array(1)(0)) & "|" & TypeName(Array("a"))',
    'split_limit_compare': 'Dim v\nv = Split("aXbxc", "x", -1, vbTextCompare)\nProbe = UBound(v) & "|" & Join(v, "|")\nv = Split("a  b", " ")\nProbe = Probe & "|" & UBound(v)',
    'join_numbers_and_errors': 'On Error Resume Next\nProbe = Join(Array(1, 2.5, True), "-")\nDim s As String\ns = Join(Array(1, Null), "-")\nProbe = Probe & "|" & Err.Number',
    'filter_case': 'Dim v\nv = Filter(Array("Apple", "banana", "APPLE pie"), "apple", True, vbTextCompare)\nProbe = UBound(v) & "|" & Join(v, ",")\nv = Filter(Array("Apple", "banana"), "a", False)\nProbe = Probe & "|" & Join(v, ",")',
    'array_compare_error': 'Dim a, b\na = Array(1): b = Array(1)\nOn Error Resume Next\nDim r As Boolean\nr = (a = b)\nProbe = Err.Number',
    'array_assign_to_typed': 'Dim a() As String\nOn Error Resume Next\na = Split("x,y", ",")\nProbe = UBound(a) & "|" & Err.Number\nDim b() As Long\nb = Array(1, 2)\nProbe = Probe & "|" & Err.Number',
    'typed_to_typed_assign': 'Dim a() As Long, b() As Long\nReDim b(1 To 2)\nb(2) = 8\na = b\nb(2) = 1\nProbe = a(2) & "|" & LBound(a)',
    'redim_with_variable_bounds': 'Dim n As Long, a() As Double\nn = 4\nReDim a(n - 1)\nProbe = UBound(a) & "|" & TypeName(a(0))',
    'dim_array_const_bound': 'Const N As Long = 3\nDim a(1 To N) As String\nProbe = UBound(a) & "|" & Len(a(1))',
    'isarray_ismissing': 'Dim v As Variant\nProbe = IsArray(Array()) & IsArray(v) & IsArray("x")',
    'for_each_empty_array': 'Dim v, x, n As Long\nv = Array()\nFor Each x In v: n = n + 1: Next\nProbe = n & "|" & IsEmpty(x)',
    'array_of_arrays_ubound': 'Dim v\nv = Array(Array(1, 2, 3), Array())\nProbe = UBound(v(0)) & "|" & UBound(v(1))',
    'redim_preserve_shrink_strings': 'Dim a() As String\nReDim a(3)\na(0) = "a": a(3) = "d"\nReDim Preserve a(1)\nProbe = UBound(a) & a(0)',
    'array_index_rounding': 'Dim a(3) As Long\na(2) = 5\nProbe = a(1.5) & "|" & a(2.5) & "|" & a("2")',
    'array_returner_types': 'Probe = TypeName(Split("a b")) & "|" & TypeName(Split("")) & "|" & TypeName(Filter(Array("a"), "a")) & "|" & TypeName(Filter(Split("a b"), "a")) & "|" & TypeName(Array()) & "|" & TypeName(Application.Transpose(Array(1)))',
    'split_element_coerce': 'Dim v\nv = Split("a b")\nv(0) = 5\nProbe = TypeName(v(0)) & "|" & VarType(v)',
    'split_into_typed': 'Dim a() As String, b() As Variant\nOn Error Resume Next\na = Split("x,y", ",")\nProbe = UBound(a) & "|" & Err.Number\nb = Split("x,y", ",")\nProbe = Probe & "|" & Err.Number\nDim c() As Long\nc = Split("1,2", ",")\nProbe = Probe & "|" & Err.Number',
    'variant_array_into_variant_dyn': 'Dim a() As Variant\nOn Error Resume Next\na = Array(1, 2)\nProbe = Err.Number & "|" & UBound(a)',
    'filter_into_string_array': 'Dim a() As String\nOn Error Resume Next\na = Filter(Array("ab", "c"), "a")\nProbe = Err.Number',
}


def macro(name: str) -> str:
    return HELPERS + '\nFunction Probe() As String\n' + CASES[name] + '\nEnd Function\n'


def main() -> None:
    record({name: macro(name) for name in CASES}, OUT)


if __name__ == '__main__':
    main()
