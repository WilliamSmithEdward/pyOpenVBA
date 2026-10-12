"""Run VBA probes in native Excel and record each one's value or run-time error number.

Shared by the measure_vba_*.py scripts: each probe is a whole module whose
``Function Probe() As String`` is run through pyvbaharness; the fixture holds
its value, or ``error`` with the number it stops with. A failed probe gets a
fresh Excel session, so one error does not leave the next probe in a broken
project. The tests replay each module through the emulator (see
tests/test_vba_builtins.py).
"""
from __future__ import annotations

import json
import re
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

VERSION_PROBE = ('Function VersionBuild() As String\n'
                 'VersionBuild = Application.Version & "|" & Application.Build\nEnd Function')


def record(modules: dict[str, str], out: Path) -> None:
    """Run every module's Probe in native Excel and write the fixture to ``out``."""
    probes: list[dict[str, object]] = []
    version_value = None
    excel: ExcelSession | None = None
    try:
        for name, module in modules.items():
            if excel is None:
                excel = ExcelSession(HarnessConfig(lock_wait_s=30.0))
                excel.__enter__()
                excel.new_document()
                version = excel.run_vba(VERSION_PROBE, 'VersionBuild')
                assert version.ok and not version.dialogs, version
                version_value = version.value
            result = excel.run_vba(module, 'Probe', timeout=30.0)
            assert not result.dialogs, (name, result)
            if result.ok:
                probe: dict[str, object] = {'name': name, 'macro': module, 'result': result.value}
            else:
                found = re.search(r'VBA error (\d+)', str(result.error))
                assert found, (name, result)
                probe = {'name': name, 'macro': module, 'error': int(found.group(1))}
                excel.__exit__(None, None, None)
                excel = None
            probes.append(probe)
            print(name, probe.get('result', probe.get('error')), flush=True)
    finally:
        if excel is not None:
            excel.__exit__(None, None, None)
    out.write_text(json.dumps({'excel_version_build': version_value, 'measured_at': datetime.now(timezone.utc).isoformat(),
                               'probes': probes}, indent=2) + '\n', encoding='utf-8')
