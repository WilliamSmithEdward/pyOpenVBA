"""Measure native standard-module record coercion and intrinsic arguments."""
from pathlib import Path

import measure_record_type_checks as probes

probes.OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/record_variant_calls.json'
probes.CASES = {
    'record_variant_assignment': ('Dim item As FirstData, value As Variant\nvalue = item\nCheck = 7', ''),
    'record_variant_byref': ('Dim item As FirstData\nChangeVariant item\nCheck = 7', 'Sub ChangeVariant(ByRef value As Variant)\nEnd Sub'),
    'record_variant_byval': ('Dim item As FirstData\nChangeVariant item\nCheck = 7', 'Sub ChangeVariant(ByVal value As Variant)\nEnd Sub'),
    'record_typename': ('Dim item As FirstData\nCheck = Len(TypeName(item))', ''),
    'record_len': ('Dim item As FirstData\nCheck = Len(item)', ''),
    'record_lenb': ('Dim item As FirstData\nCheck = LenB(item)', ''),
    'record_array_variant_byref': ('Dim items(0 To 1) As FirstData\nChangeVariant items\nCheck = 7', 'Sub ChangeVariant(ByRef value As Variant)\nEnd Sub'),
    'record_array_typename': ('Dim items(0 To 1) As FirstData\nCheck = Len(TypeName(items))', ''),
}

if __name__ == '__main__':
    probes.main()
