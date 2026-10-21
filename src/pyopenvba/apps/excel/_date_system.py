"""Workbook-local calendar selection and narrow workbookPr persistence."""
from __future__ import annotations

import re
from typing import TYPE_CHECKING

from pyopenvba._opc_xml import opc_part
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import to_bool

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Workbook

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
_ATTRIBUTE = re.compile(rb"([A-Za-z_:][A-Za-z0-9_.:-]*)\s*=\s*(?:\"[^\"]*\"|'[^']*')")


def read(book: Workbook, raw: bytes) -> None:
    """Read only the actual direct workbook properties; reject ambiguous flags."""
    part = opc_part(raw, NS, 'workbook')
    properties = [attrs for local, attrs, _, _ in part.children if local == NS + '}workbookPr']
    if len(properties) > 1:
        raise VBAUnsupportedError('Ambiguous workbook date-system properties')
    flag = properties[0].get('date1904', '0') if properties else '0'
    if flag not in ('0', '1', 'false', 'true'):
        raise VBAUnsupportedError('Invalid workbook date-system flag')
    book.epoch_1904 = flag in ('1', 'true')


def write(book: Workbook, raw: bytes) -> bytes:
    """Change the one owned attribute, retaining opaque attributes and siblings."""
    if not book.date_system_changed:
        return raw
    part = opc_part(raw, NS, 'workbook')
    properties = [entry for entry in part.children if entry[0] == NS + '}workbookPr']
    if len(properties) > 1:
        raise VBAUnsupportedError('Ambiguous workbook date-system properties')
    value = b'1' if book.epoch_1904 else b'0'
    if not properties:
        before = next((start for local, _, start, _ in part.children
                       if local not in (NS + '}fileVersion', NS + '}fileSharing')), part.closing_start)
        return part.insert(before, part.child('workbookPr', 'date1904="' + value.decode() + '"'))
    selected = opc_part(raw, NS, 'workbook', container_path=('workbookPr',))
    assert selected is not None
    start = properties[0][2]
    opening = raw[start:selected.opening_end]
    matched = False

    def replace(found: re.Match[bytes]) -> bytes:
        nonlocal matched
        if found.group(1) != b'date1904':
            return found.group(0)
        matched = True
        return b'date1904="' + value + b'"'

    changed = _ATTRIBUTE.sub(replace, opening)
    if not matched:
        at = len(changed) - (2 if changed.endswith(b'/>') else 1)
        changed = changed[:at] + b' date1904="' + value + b'"' + changed[at:]
    return raw[:start] + changed + raw[selected.opening_end:]


def set_mode(book: Workbook, value: object) -> None:
    """Keep stored serials and invalidate calculations on a calendar change."""
    wanted = to_bool(value)
    if wanted == book.epoch_1904:
        return
    if book.application.udf_callers or book.calculation_in_progress:
        raise VBAUnsupportedError('Changing the workbook date system during formula evaluation is not implemented')
    book.epoch_1904 = wanted
    book.date_system_changed = True
    book.saved = False
    for owner in book.sheets_:
        for cell in owner.cells_.values():
            if cell.formula:
                cell.stale = True
                cell.save_dirty = None
    from pyopenvba.apps.excel import _events

    _events.recalculated(book)
