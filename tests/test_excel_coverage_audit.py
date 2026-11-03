"""scripts/audit_excel_coverage.py counts what the whole model registers."""
from __future__ import annotations

import importlib.util
from pathlib import Path
from typing import Any, cast


def _report() -> dict[str, Any]:
    path = Path(__file__).parents[1] / 'scripts/audit_excel_coverage.py'
    spec = importlib.util.spec_from_file_location('audit_excel_coverage', path)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return cast(dict[str, Any], module.report())


def test_the_audit_reaches_every_module_of_the_model() -> None:
    report = _report()
    types = {binding['vba_type'] for binding in report['object_bindings']}
    assert {'Range', 'Chart', 'PivotTable', 'ListObject', 'Comment', 'Hyperlink'} <= types
    totals = report['totals']
    assert 0 < totals['registered_reference_members'] <= totals['reference_members']
    assert totals['reference_types_with_a_class'] <= totals['reference_types']
