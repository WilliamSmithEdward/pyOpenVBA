"""Workbook calculation settings shared with the application, without discarding unknown calcPr attributes."""
from __future__ import annotations

import re
from typing import TYPE_CHECKING

from pyopenvba._xml import attributes

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Workbook

_CALC = re.compile(r'<calcPr\b[^>]*(?:/>|>.*?</calcPr>)', re.DOTALL)
_MODES = {'manual': -4135, 'auto': -4105, 'autoNoTable': 2}
#: What a calcPr attribute left out means, where it is not 1 (ECMA-376 Part 1, 18.2.2).
_DEFAULTS = {'calcMode': 'auto', 'refMode': 'A1', 'iterate': '0', 'iterateCount': '100', 'iterateDelta': '0.001',
             'forceFullCalc': '0'}


def read(book: Workbook, xml: str) -> None:
    """The first ordinary workbook supplies the application calculation settings; each workbook its calcId and the
    full calculation it may want as it opens (see _calculation_state)."""
    from pyopenvba.apps.excel._calculation_state import opened

    found = _CALC.search(xml)
    values = attributes(found.group(0)) if found else {}
    _application_settings(book, values)
    calc_id = values.get('calcId', '')
    opened(book, int(calc_id) if calc_id.isdigit() else 0, values.get('fullCalcOnLoad', '0').lower() in ('1', 'true'),
           values.get('forceFullCalc', '0').lower() in ('1', 'true'))


def _application_settings(book: Workbook, values: dict[str, str]) -> None:
    application = book.application
    excluded = ('.xlam', '.xla', '.xltx', '.xltm', '.xlt')
    if book.name.lower().endswith(excluded) or any(not other.name.lower().endswith(excluded) for other in application.workbooks_.books):
        return
    application.calculation = _MODES.get(values.get('calcMode', 'auto'), -4105)
    application.calculate_before_save = values.get('calcOnSave', '1').lower() not in ('0', 'false')
    # A workbook saved in the R1C1 style says so; one opened after another keeps the style the application has
    # (tests/fixtures/reference_style.json).
    application.reference_style = -4150 if values.get('refMode') == 'R1C1' else 1
    # So do the iteration's settings, which the first workbook Excel opens brings (tests/fixtures/iteration.json).
    application.iteration = values.get('iterate', '0').lower() in ('1', 'true')
    count = values.get('iterateCount', '100')
    application.max_iterations = int(count) if count.isdigit() and 1 <= int(count) <= 32767 else 100
    try:
        application.max_change = float(values.get('iterateDelta', '0.001'))
    except ValueError:
        application.max_change = 0.001


def write(book: Workbook, xml: str) -> str:
    """Persist current mode and save behavior, leaving unrelated calculation settings intact."""
    found = _CALC.search(xml)
    old = found.group(0) if found else '<calcPr/>'
    application = book.application
    values = attributes(old)
    desired = {'calcMode': {value: key for key, value in _MODES.items()}.get(application.calculation, 'auto'),
               'calcOnSave': '1' if application.calculate_before_save else '0'}
    pending = any(cell.formula and cell.stale for sheet in book.sheets_ for cell in sheet.cells_.values())
    if not book.calculator.automatic and pending:
        desired['calcCompleted'] = '0'
    elif 'calcCompleted' in values:
        desired['calcCompleted'] = '0' if pending else '1'
    changed = old
    if application.reference_style == -4150:
        desired['refMode'] = 'R1C1'
    elif 'refMode' in values:
        # Saved in the A1 style, the workbook says nothing of it.
        changed = re.sub(r'\srefMode="[^"]*"', '', changed)
    # The calcId of the engine that last worked it out in full, what fullCalcOnLoad asked for done once it is, and
    # ForceFullCalculation (see _calculation_state).
    if book.calculation_version:
        desired['calcId'] = str(book.calculation_version)
    if not book.full_pending:
        changed = re.sub(r'\sfullCalcOnLoad="[^"]*"', '', changed)
    if book.force_full_calculation:
        desired['forceFullCalc'] = '1'
    else:
        changed = re.sub(r'\sforceFullCalc="[^"]*"', '', changed)
    # The iteration's settings are written where they are not Excel's own, off, 100 and 0.001, each on its own
    # (tests/fixtures/iteration.json).
    from pyopenvba.apps.excel._styles import excel_number

    for name, value, default in (('iterate', '1' if application.iteration else '0', '0'),
                                 ('iterateCount', str(application.max_iterations), '100'),
                                 ('iterateDelta', excel_number(application.max_change), '0.001')):
        if value != default:
            desired[name] = value
        elif name in values:
            changed = re.sub(r'\s' + name + r'="[^"]*"', '', changed)
    for name, value in desired.items():
        if values.get(name, _DEFAULTS.get(name, '1')) == value:
            continue
        pattern = re.compile(r'\s' + name + r'="[^"]*"')
        if pattern.search(changed):
            changed = pattern.sub(f' {name}="{value}"', changed)
        else:
            end = changed.index('>')
            at = end - 1 if changed[end - 1] == '/' else end
            changed = changed[:at] + f' {name}="{value}"' + changed[at:]
    if changed == old:
        return xml
    if found:
        return xml[:found.start()] + changed + xml[found.end():]
    # Preserve the workbook schema order, including extension records.
    after = re.search(r'<(?:oleSize|customWorkbookViews|pivotCaches|smartTagPr|smartTagTypes|webPublishing|fileRecoveryPr|webPublishObjects|extLst)\b|</workbook>', xml)
    assert after is not None
    return xml[:after.start()] + changed + xml[after.start():]
