"""Measure Application.ErrorCheckingOptions as a new Excel session reads it: each property's type and value.

Only read: the options persist in the registry, so the probe sets none of them.

    python scripts/measure_error_checking_options.py

writes tests/fixtures/error_checking_options.json, which tests/test_excel_error_checking_options.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/error_checking_options.json'
NAMES = ['BackgroundChecking', 'EvaluateToError', 'TextDate', 'NumberAsText', 'InconsistentFormula',
         'OmittedCells', 'UnlockedFormulaCells', 'EmptyCellReferences', 'ListDataValidation',
         'InconsistentTableFormula', 'MisleadingNumberFormats', 'IndicatorColorIndex']
SOURCE = '\n'.join([
    'Public Function Probe() As String', '    Dim o As Object, s As String', '    Set o = Application.ErrorCheckingOptions',
    *[f'    s = s & "|{name}=" & TypeName(o.{name}) & ":" & CStr(o.{name})' for name in NAMES],
    '    s = s & "|Type=" & TypeName(o) & "|Parent=" & TypeName(o.Parent) & "|Creator=" & o.Creator',
    '    Probe = Mid(s, 2)', 'End Function']) + '\n'


def main() -> None:
    with ExcelSession(HarnessConfig(lock_wait_s=600.0)) as excel:
        excel.new_document()
        result = excel.run_vba(SOURCE, 'Probe', (), timeout=60.0)
    assert result.ok, result
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': SOURCE, 'answer': result.value}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    print(result.value)


if __name__ == '__main__':
    main()
