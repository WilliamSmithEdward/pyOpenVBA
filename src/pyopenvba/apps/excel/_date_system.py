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
    """Read only the actual direct workbook properties; reject ambiguous flags.

    The part is read in whatever encoding it declares, with DTDs and
    declared entities refused, as every other Office XML part is.
    """
    from pyopenvba.apps.excel._charts_file import xml_tree

    root = xml_tree(raw)
    if root.tag != '{' + NS + '}workbook':
        raise VBAUnsupportedError('A workbook part outside the transitional SpreadsheetML namespace is not implemented')
    properties = [child.attrib for child in root if child.tag == '{' + NS + '}workbookPr']
    if len(properties) > 1:
        raise VBAUnsupportedError('Ambiguous workbook date-system properties')
    flag = properties[0].get('date1904', '0') if properties else '0'
    if flag not in ('0', '1', 'false', 'true'):
        raise VBAUnsupportedError('Invalid workbook date-system flag')
    book.epoch_1904 = flag in ('1', 'true')
    book.workbook_pr = dict(properties[0]) if properties else {}


#: workbookPr's attributes in the order the schema, and Excel, writes them.
_ORDER = ('date1904', 'showObjects', 'showBorderUnselectedTables', 'filterPrivacy', 'promptedSolutions',
          'showInkAnnotation', 'backupFile', 'saveExternalLinkValues', 'updateLinks', 'codeName', 'hidePivotFieldList',
          'showPivotChartFilter', 'allowRefreshQuery', 'publishItems', 'checkCompatibility', 'autoCompressPictures',
          'refreshAllConnections', 'defaultThemeVersion')


def write(book: Workbook, raw: bytes) -> bytes:
    """Change the owned attributes, the date system and the flags a macro set (_workbook_settings), retaining
    opaque attributes and siblings."""
    raw = _write_flags(book, raw)
    if not book.date_system_changed:
        return raw
    try:
        raw.decode('utf-8')
    except UnicodeDecodeError:
        raise VBAUnsupportedError('Changing the date system of a workbook part that is not UTF-8 is not '
                                  'implemented') from None
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


def _write_flags(book: Workbook, raw: bytes) -> bytes:
    """The workbookPr flags a macro changed, each where the schema orders it (tests/fixtures/workbook_settings.json:
    hidePivotFieldList="1" before defaultThemeVersion)."""
    changed = {name: book.workbook_pr[name] for name in book.workbook_pr_changed if name in book.workbook_pr}
    if not changed:
        return raw
    try:
        raw.decode('utf-8')
    except UnicodeDecodeError:
        raise VBAUnsupportedError('Changing the properties of a workbook part that is not UTF-8 is not '
                                  'implemented') from None
    part = opc_part(raw, NS, 'workbook')
    properties = [entry for entry in part.children if entry[0] == NS + '}workbookPr']
    if len(properties) > 1:
        raise VBAUnsupportedError('Ambiguous workbook properties')
    if not properties:
        before = next((start for local, _, start, _ in part.children
                       if local not in (NS + '}fileVersion', NS + '}fileSharing')), part.closing_start)
        text = ' '.join(f'{name}="{changed[name]}"' for name in _ORDER if name in changed)
        return part.insert(before, part.child('workbookPr', text))
    selected = opc_part(raw, NS, 'workbook', container_path=('workbookPr',))
    assert selected is not None
    start = properties[0][2]
    opening = raw[start:selected.opening_end].decode('utf-8')
    present = [found.group(1).decode() for found in _ATTRIBUTE.finditer(opening.encode())]
    for name, value in changed.items():
        if name in present:
            opening = re.sub(rf'\b{name}\s*=\s*(?:"[^"]*"|\'[^\']*\')', f'{name}="{value}"', opening)
            continue
        later = next((other for other in present if other in _ORDER and _ORDER.index(other) > _ORDER.index(name)),
                     None)
        at = opening.index(f' {later}') if later is not None else \
            len(opening) - (2 if opening.endswith('/>') else 1)
        opening = opening[:at] + f' {name}="{value}"' + opening[at:]
        present.append(name)
    return raw[:start] + opening.encode('utf-8') + raw[selected.opening_end:]


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
