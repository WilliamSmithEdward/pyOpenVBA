"""Record native compiler rejection of parenthesized whole-record arguments."""
from __future__ import annotations
import json
from datetime import datetime, timezone
from pathlib import Path
from pyvbaharness import ExcelSession, HarnessConfig
from measure_record_copy import CASES, macro
OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/record_argument_errors.json'
def main() -> None:
    CASES['parenthesized_record'] = 'first.Value = 5\nChange (first)\nProbe = CStr(first.Value)'
    source = macro('parenthesized_record')
    with ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
        excel.new_document()
        version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & Application.Build\nEnd Function', 'VersionBuild')
        assert version.ok and not version.dialogs, version
        result = excel.run_vba(source, 'Probe', timeout=30.0)
        assert not result.ok and result.dialogs and result.dialogs[0].classification == 'compile-error', result
        message = result.dialogs[0].message
        print(message, flush=True)
    OUT.write_text(json.dumps({'excel_version_build': version.value, 'measured_at': datetime.now(timezone.utc).isoformat(), 'probes': [{'name': 'parenthesized_record', 'macro': source, 'compile_error': message}]}, indent=2) + '\n', encoding='utf-8')
if __name__ == '__main__':
    main()
