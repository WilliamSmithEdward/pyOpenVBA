"""Members read on a new workbook, replayed from live Excel (scripts/measure_missing_reads.py,
tests/fixtures/missing_reads.json).

The fixture reads more than the model implements; the replay keeps the reads in IMPLEMENTED and drops the others'
lines from the probe, since a read the model does not implement stops the whole procedure."""
from __future__ import annotations

import json
from pathlib import Path
import re
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/missing_reads.json').read_text(encoding='utf-8'))
IMPLEMENTED = {
    'wb.Password', 'wb.WritePassword', 'wb.PasswordEncryptionProvider', 'wb.PasswordEncryptionAlgorithm',
    'wb.PasswordEncryptionKeyLength', 'wb.PasswordEncryptionFileProperties', 'wb.Container',
    'wb.ContentTypeProperties', 'wb.ServerPolicy', 'wb.ExternalCodeServiceTimeout', 'wb.DocumentLibraryVersions',
    'wb.SharedWorkspace', 'wb.Permission', 'wb.ActiveSlicer',
    'ch.Index', 'ch.CodeName', 'ch.Visible', 'ch.ProtectContents', 'ch.ProtectData', 'ch.ProtectDrawingObjects',
    'ch.ProtectFormatting', 'ch.ProtectionMode', 'ch.ProtectSelection', 'ch.ShowAllFieldButtons',
    'ch.ShowAxisFieldButtons', 'ch.ShowLegendFieldButtons', 'ch.ShowReportFilterFieldButtons',
    'ch.ShowValueFieldButtons', 'ch.ShowExpandCollapseEntireFieldButtons', 'ch.PivotLayout', 'ch.Tab', 'ch.Next',
    'ch.Previous',
    'ax.Application', 'ax.CategorySortOrder', 'ax.DisplayUnitLabel', 'vx.DisplayUnitLabel', 'vx.CategorySortOrder',
    'at.Application', 'le.Application', 'sr.BubbleSizes', 'co.ZOrder', 'co.ProtectChartObject',
    'sh.FormControlType', 'sh.HasSmartArt', 'sh.Child', 'sh.Connector', 'sh.Hyperlink', 'sh.BlackWhiteMode',
    'sr2.Application', 'sr2.Parent', 'sr2.ID', 'sr2.HasChart', 'sr2.Child', 'sr2.Connector', 'sr2.ZOrderPosition',
    'sr2.GroupItems', 'sr2.ParentGroup', 'tf.Application', 'tf.Parent',
    'g.Child', 'g.GroupItems(1).Child', 'g.GroupItems(1).Connector', 'g.FormControlType', 'bt.FormControlType',
    'cb.FormControlType', 'bt.HasSmartArt', 'bt.BlackWhiteMode', 'ws.Shapes.Range(Array(g.Name)).Child',
    'ws.Shapes.Range(Array(g.Name)).GroupItems', 'ws.Shapes.Range(Array(sh.Name, g.Name)).ID',
    'ws.Shapes.Range(Array(sh.Name, g.Name)).HasChart', 'ws.Shapes.Range(Array(sh.Name, g.Name)).ZOrderPosition',
    'ws.Shapes.Range(Array(sh.Name, g.Name)).Parent', 'ws.Shapes.Range(Array(co.Name)).HasChart',
}
EXPECTED = {read: answer for read, answer in RECORD['cases'] if read in IMPLEMENTED}
READ_LINE = re.compile(r'    s = s & "@@@" & Fetch\((.*), "(\w+)"\)$')


def _kept(line: str) -> bool:
    found = READ_LINE.match(line)
    return found is None or f'{found[1]}.{found[2]}' in IMPLEMENTED


@pytest.fixture(scope='module')
def answers() -> dict[str, str]:
    source = '\n'.join(line for line in RECORD['source'].split('\n') if _kept(line))
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source, name='ReadsProbe')
    return dict(zip(EXPECTED, str(app.run('ReadsProbe.Probe')).split('@@@'), strict=True))


def test_every_implemented_read_is_in_the_fixture() -> None:
    assert set(EXPECTED) == IMPLEMENTED


@pytest.mark.parametrize('read', list(EXPECTED))
def test_each_read_as_excel_answered(answers: dict[str, str], read: str) -> None:
    assert answers[read] == EXPECTED[read]
