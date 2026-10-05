"""Measure native typed-record array parameter matching and temporaries."""
from pathlib import Path

import measure_record_type_checks as probes

probes.OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/record_array_arguments.json'
probes.HELPERS += '\nSub ChangeArray(ByRef items() As FirstData)\nitems(0).Value = items(0).Value + 10\nEnd Sub'
probes.CASES = {
    'fixed_array_parameter': ('Dim items(0 To 1) As FirstData\nitems(0).Value = 5\nHelpers.ChangeArray items\nCheck = items(0).Value', ''),
    'dynamic_array_parameter': ('Dim items() As FirstData\nReDim items(0 To 1)\nitems(0).Value = 5\nHelpers.ChangeArray items\nCheck = items(0).Value', ''),
    'mismatched_record_array': ('Dim items(0 To 1) As SecondData\nHelpers.ChangeArray items\nCheck = items(0).Value', ''),
    'scalar_to_array_parameter': ('Dim item As FirstData\nHelpers.ChangeArray item\nCheck = item.Value', ''),
    'variant_array_parameter': ('Dim items(0 To 1) As FirstData, value As Variant\nitems(0).Value = 5\nvalue = items\nHelpers.ChangeArray value\nCheck = items(0).Value', ''),
    'parenthesized_record_array': ('Dim items(0 To 1) As FirstData\nitems(0).Value = 5\nHelpers.ChangeArray (items)\nCheck = items(0).Value', ''),
    'array_to_scalar_parameter': ('Dim items(0 To 1) As FirstData\nHelpers.Change items\nCheck = items(0).Value', ''),
}

if __name__ == '__main__':
    probes.main()
