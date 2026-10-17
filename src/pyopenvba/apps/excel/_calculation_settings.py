"""Workbook calculation settings shared with the application, without discarding unknown calcPr attributes."""
from __future__ import annotations

import re
from typing import TYPE_CHECKING

from pyopenvba._xml import attributes

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Workbook

_CALC = re.compile(r'<calcPr\b[^>]*(?:/>|>.*?</calcPr>)', re.DOTALL)
_MODES = {'manual': -4135, 'auto': -4105, 'autoNoTable': 2}


def read(book: Workbook, xml: str) -> None:
    """The first ordinary workbook supplies the application calculation settings."""
    application = book.application
    excluded = ('.xlam', '.xla', '.xltx', '.xltm', '.xlt')
    if book.name.lower().endswith(excluded) or any(not other.name.lower().endswith(excluded) for other in application.workbooks_.books):
        return
    found = _CALC.search(xml)
    values = attributes(found.group(0)) if found else {}
    application.calculation = _MODES.get(values.get('calcMode', 'auto'), -4105)
    application.calculate_before_save = values.get('calcOnSave', '1').lower() not in ('0', 'false')


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
    for name, value in desired.items():
        if values.get(name, 'auto' if name == 'calcMode' else '1') == value:
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
