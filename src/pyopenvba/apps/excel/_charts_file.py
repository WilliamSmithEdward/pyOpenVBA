"""Chart parts linked to the worksheet's existing drawing layer.

New-chart markup is adapted from pyOfficeEditor at
4723383e5d3039763e64ecd2a3ee092786c91bb5. Imported chart XML is preserved;
editing imported chart contents is deliberately refused until narrow writers exist.
"""
from __future__ import annotations

import posixpath
import re
import xml.etree.ElementTree as ET
from typing import TYPE_CHECKING

from pyopenvba._xml import attributes, unescape
from pyopenvba._a1 import parse_reference
from pyopenvba.apps.excel._chart_xml import CHART_KINDS, CT_CHART, SeriesData, cache_markup, chart_frame, chart_part
from pyopenvba.apps.excel._charts import CHART_TYPES, Chart, ChartState, Series, SeriesState, referenced_values
from pyopenvba.exceptions import VBAUnsupportedError, VBARuntimeError
from pyopenvba.interpreter._values import EMPTY, VBAErrorValue, to_number, to_text
from pyopenvba.powerquery._sheets import add_content_type, add_relationship, drop_content_type

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet
    from pyopenvba.powerquery._opc import OpcFile

NS = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
REL_CHART = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'

def rels_part(part: str) -> str:
    folder, name = posixpath.split(part)
    return f'{folder}/_rels/{name}.rels'

def _targets(package: OpcFile, part: str) -> dict[str, str]:
    rels = rels_part(part)
    if not package.has(rels):
        return {}
    root = ET.fromstring(package.read(rels))
    return {entry.attrib.get('Id', ''): posixpath.normpath(posixpath.join(posixpath.dirname(part), entry.attrib['Target']))
            if not entry.attrib['Target'].startswith('/') else entry.attrib['Target'][1:]
            for entry in root if entry.attrib.get('Type') == REL_CHART and 'Target' in entry.attrib}

def read(sheet: Worksheet, package: OpcFile) -> None:
    targets = _targets(package, sheet.drawing_part)
    for shape in sheet.shapes_:
        if shape.kind != 'chart':
            continue
        found = re.search(r'<(?:\w+:)?chart\b([^>]*)', shape.source)
        if found is None:
            continue
        fields = attributes(found.group(1))
        part = targets.get(fields.get('r:id', ''), '')
        if part and package.has(part):
            xml = package.read(part).decode('utf-8')
            sheet.chart_states[shape.shape_id] = read_chart(xml, part)

def read_chart(xml: str, part: str) -> ChartState:
    root = ET.fromstring(xml)
    result = ChartState(part=part, original=xml, dirty=False, chart_type=0)
    title = root.find('c:chart/c:title', NS)
    result.title = ''.join(title.itertext()) if title is not None else None
    if title is not None:
        result.title = ''.join(node.text or '' for node in title.findall('.//a:t', NS)) or 'Chart Title'
    result.legend = root.find('c:chart/c:legend', NS) is not None
    plot = root.find('c:chart/c:plotArea', NS)
    if plot is None:
        return result
    plots = [entry for entry in plot if entry.tag.rpartition('}')[2].endswith('Chart')]
    if len(plots) != 1:
        return result
    entry = plots[0]
    kind = entry.tag.rpartition('}')[2]
    grouping = entry.find('c:grouping', NS)
    group = grouping.attrib.get('val', '') if grouping is not None else ''
    if kind == 'barChart' and group == 'clustered':
        direction = entry.find('c:barDir', NS)
        result.chart_type = 57 if direction is not None and direction.attrib.get('val') == 'bar' else 51
    elif kind == 'lineChart' and group == 'standard':
        marker = entry.find('c:marker', NS)
        result.chart_type = 65 if marker is not None and marker.attrib.get('val') == '1' else 4
    elif kind == 'scatterChart':
        style = entry.find('c:scatterStyle', NS)
        result.chart_type = -4169 if style is not None and style.attrib.get('val') == 'marker' else 74
    elif kind in ('pieChart', 'doughnutChart') or kind == 'areaChart' and group in ('standard', ''):
        result.chart_type = {'pieChart': 5, 'doughnutChart': -4120, 'areaChart': 1}[kind]
    for series in entry.findall('c:ser', NS):
        item = SeriesState()
        name_ref = series.find('c:tx/c:strRef/c:f', NS)
        item.name_reference = name_ref.text or '' if name_ref is not None else ''
        name = series.find('c:tx/c:v', NS)
        item.name = name.text if name is not None else None
        cat = series.find('c:cat', NS)
        if cat is None:
            cat = series.find('c:xVal', NS)
        val = series.find('c:val', NS)
        if val is None:
            val = series.find('c:yVal', NS)
        item.category_reference, item.categories = _data(cat)
        item.value_reference, item.values = _data(val)
        result.series.append(item)
    if result.series and result.series[0].value_reference:
        try:
            area = parse_reference(result.series[0].value_reference)[0]
            if area.top == area.bottom and area.left != area.right:
                result.plot_by = 1
        except ValueError:
            pass
    return result

def _data(node: ET.Element | None) -> tuple[str, list[object]]:
    if node is None:
        return '', []
    formula = node.find('.//c:f', NS)
    if formula is not None:
        return formula.text or '', []
    numeric = node.find('c:numLit', NS) is not None
    values: list[object] = []
    for point in node.findall('.//c:pt', NS):
        index = int(point.attrib.get('idx', '0'))
        while len(values) <= index:
            values.append(EMPTY)
        text = point.findtext('c:v', default='', namespaces=NS)
        try:
            values[index] = float(text) if numeric else text
        except ValueError:
            values[index] = EMPTY
    return '', values

def _number(value: object) -> float | None:
    if value is EMPTY or isinstance(value, VBAErrorValue):
        return None
    return float(to_number(value))

def refresh(sheet: Worksheet, package: OpcFile) -> None:
    """Refresh only changed reference caches; leave every other chart byte alone."""
    if not any(owner.dirty for owner in sheet.book.sheets_):
        return
    for chart in sheet.chart_states.values():
        if not chart.part or not package.has(chart.part):
            continue
        original = package.read(chart.part).decode('utf-8')
        def update(match: re.Match[str]) -> str:
            raw = match.group(0)
            formula = re.search(r'<c:f>(.*?)</c:f>', raw)
            cached = re.search(r'<c:(numCache|strCache)>.*?</c:\1>', raw)
            if formula is None or cached is None:
                return raw
            numeric = cached.group(1) == 'numCache'
            try:
                source = referenced_values(sheet, unescape(formula.group(1)))
                values: list[str | float | None] = [_number(value) if numeric else to_text(value) for value in source]
            except (ValueError, KeyError, VBAUnsupportedError, VBARuntimeError):
                # Unknown/external chart reference expressions retain their original caches.
                return raw
            root = ET.fromstring(f'<root xmlns:c="{NS["c"]}">{cached.group(0)}</root>')
            count = root.find('.//c:ptCount', NS)
            points = {int(point.attrib['idx']): point.findtext('c:v', default='', namespaces=NS)
                      for point in root.findall('.//c:pt', NS)}
            expected = {index: value for index, value in enumerate(values) if value is not None and value != ''}
            try:
                same = (count is not None and int(count.attrib['val']) == len(values) and
                        set(points) == set(expected) and all(
                            (float(points[index]) == value if numeric else points[index] == value)
                            for index, value in expected.items()))
            except ValueError:
                same = False
            if same:
                return raw
            code = root.findtext('.//c:formatCode', default='General', namespaces=NS)
            replacement = cache_markup(values, numeric=numeric, code=code)
            return raw[:cached.start()] + replacement + raw[cached.end():]
        updated = re.sub(r'<c:(numRef|strRef)>.*?</c:\1>', update, original)
        if updated != original:
            package.write(chart.part, updated.encode('utf-8'))

def markup(sheet: Worksheet, shape_id: int) -> str:
    shape = next(item for item in sheet.shapes_ if item.shape_id == shape_id)
    chart = Chart(sheet, shape)
    data: list[SeriesData] = []
    for item in chart.state.series:
        series = Series(chart, item)
        categories = series.categories()
        numeric = bool(categories) and all(isinstance(value, (int, float)) and not isinstance(value, bool) for value in categories)
        data.append(SeriesData(values=item.value_reference, value_cache=[_number(value) for value in series.values()],
                               name=item.name_reference or None, name_cache=to_text(series.Name()) if item.name_reference else item.name,
                               categories=item.category_reference or ('' if item.categories else None),
                               category_cache=[_number(value) if numeric else to_text(value) for value in categories],
                               numeric_categories=numeric))
    kind = CHART_TYPES.get(chart.state.chart_type)
    if kind not in CHART_KINDS:
        raise ValueError(f'Chart type {chart.state.chart_type} has no chart writer')
    xml = chart_part(kind, data, title=chart.state.title)
    if chart.state.title is None:
        xml = re.sub(r'<c:title>.*?</c:title><c:autoTitleDeleted val="0"/>', '<c:autoTitleDeleted val="1"/>', xml, count=1)
    if not chart.state.legend:
        xml = re.sub(r'<c:legend>.*?</c:legend>', '', xml, count=1)
    if chart.state.chart_type == -4169:
        xml = xml.replace('<c:scatterStyle val="lineMarker"/>', '<c:scatterStyle val="marker"/>')
    # Chart builder references become literal caches when the series was assigned an array.
    xml = re.sub(r'<c:numRef><c:f></c:f><c:numCache>(.*?)</c:numCache></c:numRef>', r'<c:numLit>\1</c:numLit>', xml)
    xml = re.sub(r'<c:strRef><c:f></c:f><c:strCache>(.*?)</c:strCache></c:strRef>', r'<c:strLit>\1</c:strLit>', xml)
    return xml

def write(sheet: Worksheet, package: OpcFile, drawing: str) -> None:
    retained = {chart.part for owner in sheet.book.sheets_ for identifier, chart in owner.chart_states.items()
                if any(shape.shape_id == identifier and shape.kind == 'chart' for shape in owner.shapes_)}
    retired = set(sheet.deleted_chart_parts)
    for part in retired - retained:
        package.remove(part)
        package.remove(rels_part(part))
        drop_content_type(package, part)
    rels = rels_part(drawing)
    if package.has(rels) and retired:
        xml = package.read(rels).decode('utf-8')
        def keep(match: re.Match[str]) -> str:
            fields = attributes(match.group(0))
            target = fields.get('Target', '')
            part = posixpath.normpath(posixpath.join(posixpath.dirname(drawing), target)) if not target.startswith('/') else target[1:]
            return '' if fields.get('Type') == REL_CHART and part in retired else match.group(0)
        package.write(rels, re.sub(r'<Relationship\b[^>]*/>', keep, xml).encode('utf-8'))
    for shape in sheet.shapes_:
        if shape.kind != 'chart' or shape.shape_id not in sheet.chart_states:
            continue
        chart = sheet.chart_states[shape.shape_id]
        if not chart.part:
            number = 1
            while package.has(f'xl/charts/chart{number}.xml'):
                number += 1
            chart.part = f'xl/charts/chart{number}.xml'
            identifier = add_relationship(package, rels, REL_CHART, posixpath.relpath(chart.part, posixpath.dirname(drawing)))
            shape.source = chart_frame(shape_id=shape.shape_id, name=shape.name, relationship=identifier,
                                       left=shape.left, top=shape.top, width=shape.width, height=shape.height,
                                       grid=sheet.drawing_grid())
        if chart.dirty and not chart.original:
            package.write(chart.part, markup(sheet, shape.shape_id).encode('utf-8'))
            add_content_type(package, chart.part, CT_CHART)
            chart.dirty = False
    sheet.deleted_chart_parts.clear()
