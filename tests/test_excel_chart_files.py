"""Chart parts, drawing relationships and runtime readback share one state."""
from __future__ import annotations
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET
import importlib.util
import os
import pytest
from pyopenvba.apps.excel import ExcelApplication

SOURCE = '''Sub MakeChart()
Dim item As ChartObject, series As Series
Set item = ActiveSheet.ChartObjects.Add(10, 20, 300, 180)
item.Name = "RevenuePlot"
item.Chart.HasTitle = True
item.Chart.ChartTitle.Text = "Revenue & Expenses"
Set series = item.Chart.SeriesCollection.NewSeries
series.Name = "Revenue"
series.Values = Array(2, 4)
series.XValues = Array("a", "b")
End Sub
Function ReadChart() As String
Dim item As ChartObject
Set item = ActiveSheet.ChartObjects(1)
ReadChart = item.Name & "|" & item.Chart.ChartType & "|" & item.Chart.ChartTitle.Text & "|" & item.Chart.SeriesCollection(1).Formula
End Function'''

def test_new_chart_parts_and_readback(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='Harness')
    app.run('MakeChart')
    path = tmp_path / 'chart.xlsx'
    app.save(path)
    with ZipFile(path) as package:
        xml = package.read('xl/charts/chart1.xml')
        root = ET.fromstring(xml)
        ns = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart', 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
        assert root.findtext('.//c:tx/c:v', namespaces=ns) == 'Revenue'
        assert [node.text for node in root.findall('.//c:cat/c:strLit/c:pt/c:v', ns)] == ['a', 'b']
        assert [node.text for node in root.findall('.//c:val/c:numLit/c:pt/c:v', ns)] == ['2', '4']
        assert b'Revenue &amp; Expenses' in xml
        assert b'chart1.xml' in package.read('xl/drawings/_rels/drawing1.xml.rels')
    reopened = ExcelApplication.open(path)
    reopened.add_module(SOURCE, name='Harness')
    assert reopened.run('ReadChart') == 'RevenuePlot|51|Revenue & Expenses|=SERIES("Revenue",{"a","b"},{2,4},1)'
    second = tmp_path / 'same.xlsx'
    reopened.save(second)
    with ZipFile(path) as original, ZipFile(second) as saved:
        assert original.read('xl/charts/chart1.xml') == saved.read('xl/charts/chart1.xml')

def test_chart_delete_removes_part_and_relationship(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE + '\nSub RemoveChart()\nActiveSheet.ChartObjects(1).Delete\nEnd Sub', name='Harness')
    app.run('MakeChart')
    app.save(tmp_path / 'before.xlsx')
    app.run('RemoveChart')
    after = tmp_path / 'after.xlsx'
    app.save(after)
    with ZipFile(after) as package:
        assert 'xl/charts/chart1.xml' not in package.namelist()
        assert b'/chart"' not in package.read('xl/drawings/_rels/drawing1.xml.rels')


def test_linked_cache_updates_without_replacing_chart_formatting(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Sub MakeLinked()\nRange("A1").Value = "Category"\nRange("B1").Value = "Revenue"\nRange("A2").Value = "a"\nRange("B2").Value = 2\nDim item As ChartObject\nSet item = ActiveSheet.ChartObjects.Add(10, 20, 300, 180)\nitem.Chart.SetSourceData Range("A1:B2"), xlColumns\nEnd Sub\nSub UpdateValue()\nRange("B2").Value = 20\nEnd Sub', name='Harness')
    app.run('MakeLinked')
    before = tmp_path / 'linked.xlsx'
    app.save(before)
    reopened = ExcelApplication.open(before)
    reopened.add_module('Sub UpdateValue()\nRange("B2").Value = 20\nEnd Sub', name='Harness')
    reopened.run('UpdateValue')
    after = tmp_path / 'linked_updated.xlsx'
    reopened.save(after)
    with ZipFile(before) as original, ZipFile(after) as saved:
        old, new = original.read('xl/charts/chart1.xml'), saved.read('xl/charts/chart1.xml')
        assert b'<c:v>20</c:v>' in new
        assert old.replace(b'<c:v>2</c:v>', b'<c:v>20</c:v>') == new


def test_native_smart_art_is_not_an_embedded_chart() -> None:
    app = ExcelApplication.open(Path(__file__).parent / 'fixtures/charts/smart_art.xlsx')
    app.add_module('Function Probe() As String\nProbe = ActiveSheet.ChartObjects.Count & ":" & ActiveSheet.Shapes(1).HasChart & ":" & ActiveSheet.Shapes(1).Type\nEnd Function', name='Harness')
    assert app.run('Harness.Probe') == '0:0:24'


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='requires explicitly enabled native Excel')
@pytest.mark.parametrize('chart_type', [51, 57, 4, 65, 5, -4120, -4169, 1])
def test_native_excel_opens_the_model_chart(tmp_path: Path, chart_type: int) -> None:
    harness = pytest.importorskip('pyvbaharness')

    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE + '\nSub SetKind(ByVal kind As Long)\nActiveSheet.ChartObjects(1).Chart.ChartType = kind\nEnd Sub', name='Harness')
    app.run('MakeChart')
    app.run('SetKind', chart_type)
    path = tmp_path / 'native_chart.xlsx'
    app.save(path)
    with harness.ExcelSession() as excel:
        excel.open_document(path)
        result = excel.run_vba(SOURCE, 'ReadChart', timeout=60.0)
        assert result.ok and not result.dialogs, result
        assert result.value == f'RevenuePlot|{chart_type}|Revenue & Expenses|=SERIES("Revenue",{{"a","b"}},{{2,4}},1)'
