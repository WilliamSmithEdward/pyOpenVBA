"""Chart parts linked to the worksheet's existing drawing layer.

New-chart markup is adapted from pyOfficeEditor at
4723383e5d3039763e64ecd2a3ee092786c91bb5. Imported chart XML is preserved;
editing imported chart contents is deliberately refused until narrow writers exist.
"""
from __future__ import annotations

import posixpath
import re
import xml.etree.ElementTree as ET
from xml.parsers import expat
from typing import TYPE_CHECKING, NoReturn

from pyopenvba._xml import attributes, escape as escape_attribute, escape_text, unescape
from pyopenvba._relationships import in_office_order
from pyopenvba._a1 import parse_reference
from pyopenvba.apps.excel._chart_xml import BLANKS, CHART_KINDS, DATA_TABLE, legend_markup, CT_CHART, SeriesData, cache_markup, chart_frame, chart_part
from pyopenvba.apps.excel._charts import (MAX_CHART_POINTS, CHART_TYPES, STACKED, Chart, ChartState, Series, SeriesState,
                                          MARKER_AUTOMATIC, MARKER_NONE, MARKER_SYMBOLS, MARKER_UNREAD, marker_markup,
                                          category_levels, linked_title, referenced_values, shown_cells, shown_type)
from pyopenvba.apps.excel._chart_title import read_look, title_element
from pyopenvba.apps.excel._chart_error_bars import ErrorBarState, markup as error_bars_markup
from pyopenvba.apps.excel._chart_trendlines import TYPES as TRENDLINE_TYPES, TrendlineState, markup as trendlines_markup
from pyopenvba.exceptions import VBAUnsupportedError, VBARuntimeError
from pyopenvba.interpreter._values import EMPTY, VBAErrorValue, to_number, to_text
from pyopenvba.powerquery._sheets import add_content_type, add_relationship, drop_content_type

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet
    from pyopenvba.powerquery._opc import OpcFile

NS = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
REL_CHART = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
C16R3 = 'http://schemas.microsoft.com/office/drawing/2017/03/chart'

# Shared with the pivot reader: constructing a node does not parse XML.
XmlElement = ET.Element
XmlParseError = expat.ExpatError


def xml_bytes(root: XmlElement) -> bytes:
    """Serialize with local namespace scopes, including aliases in QName values.

    ElementTree's automatic prefix replacement cannot update arbitrary attribute
    values such as mc:Ignorable or Choice Requires. Retain parsed declarations
    and choose lexical element/attribute names against those same local scopes.
    New namespaces receive unused prefixes without changing global ET state.
    """
    reserved = {key[6:] for entry in root.iter() for key in entry.attrib if key.startswith('xmlns:')}
    counter = 0

    def copy(entry: XmlElement, inherited: dict[str, str]) -> XmlElement:
        nonlocal counter
        scope = dict(inherited)
        declarations = {key: value for key, value in entry.attrib.items()
                        if key == 'xmlns' or key.startswith('xmlns:')}
        for key, value in declarations.items():
            scope['' if key == 'xmlns' else key[6:]] = value
        if not entry.tag.startswith('{') and scope.get(''):
            declarations['xmlns'] = ''
            scope[''] = ''

        def name(expanded: str, *, attribute: bool = False) -> str:
            nonlocal counter
            if not expanded.startswith('{'):
                return expanded
            uri, local = expanded[1:].split('}', 1)
            prefix = next((prefix for prefix, value in scope.items()
                           if value == uri and (prefix or not attribute)), None)
            if prefix is None:
                while 'ns' + str(counter) in reserved:
                    counter += 1
                prefix = 'ns' + str(counter)
                counter += 1
                reserved.add(prefix)
                scope[prefix] = uri
                declarations['xmlns:' + prefix] = uri
            return prefix + ':' + local if prefix else local

        tag = name(entry.tag)
        # An element's own declarations keep their place among its attributes, as Excel writes
        # <ext uri="..." xmlns:x15="...">; one made here for a namespace goes first.
        attrs = {key if key == 'xmlns' or key.startswith('xmlns:') else name(key, attribute=True): value
                 for key, value in entry.attrib.items()}
        attrs.update((key, value) for key, value in declarations.items() if key in attrs)
        cloned = XmlElement(tag, {**{key: value for key, value in declarations.items() if key not in attrs}, **attrs})
        cloned.text, cloned.tail = entry.text, entry.tail
        cloned.extend(copy(child, scope) for child in entry)
        return cloned

    written = ET.tostring(copy(root, {'xml': 'http://www.w3.org/XML/1998/namespace'}),
                          encoding='utf-8', xml_declaration=True)
    # ElementTree closes an empty element with " />" and Excel with "/>". The space can stand nowhere else: ">" in
    # text and attribute values is written as &gt;.
    return written.replace(b' />', b'/>')


def xml_parser() -> expat.XMLParserType:
    """Create the shared namespace-aware parser with DTD/entity handling refused."""
    parser = expat.ParserCreate(namespace_separator='}')
    def refuse(*_args: object) -> NoReturn:
        raise ValueError('Office XML must not contain DTDs or declared entities')
    parser.StartDoctypeDeclHandler = refuse
    parser.EntityDeclHandler = refuse
    parser.ExternalEntityRefHandler = refuse
    parser.SetParamEntityParsing(expat.XML_PARAM_ENTITY_PARSING_NEVER)
    return parser


def xml_tree(data: str | bytes) -> ET.Element:
    """Build a read-only tree with DTDs and all declared entities refused."""
    parser = xml_parser()
    builder = ET.TreeBuilder()
    declarations: dict[str, str] = {}
    def namespace(prefix: str | None, uri: str | None) -> None:
        declarations['xmlns:' + prefix if prefix else 'xmlns'] = uri or ''
    def expanded(name: str) -> str:
        return '{' + name if '}' in name else name
    def start(name: str, attrs: dict[str, str]) -> None:
        builder.start(expanded(name), {**declarations, **{expanded(key): value for key, value in attrs.items()}})
        declarations.clear()
    def end(name: str) -> None:
        builder.end(expanded(name))
    parser.StartElementHandler = start
    parser.StartNamespaceDeclHandler = namespace
    parser.EndElementHandler = end
    parser.CharacterDataHandler = builder.data
    parser.Parse(data, True)
    return builder.close()

def rels_part(part: str) -> str:
    folder, name = posixpath.split(part)
    return f'{folder}/_rels/{name}.rels'

def _targets(package: OpcFile, part: str) -> dict[str, str]:
    rels = rels_part(part)
    if not package.has(rels):
        return {}
    root = xml_tree(package.read(rels))
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

def read_display(result: ChartState, root: ET.Element) -> None:
    """The chart's display options as Excel reads them: an element left out, or written with no val, reads as
    DisplayBlanksAs 2, PlotVisibleOnly and ShowDataLabelsOverMaximum True, DisplayValueNotAvailableAsBlank False."""
    def flag(path: str, fallback: bool, namespaces: dict[str, str]) -> bool | None:
        node = root.find(path, namespaces)
        text = node.get('val') if node is not None else None
        return fallback if text is None else {'1': True, 'true': True, '0': False, 'false': False}.get(text)

    blanks = root.find('c:chart/c:dispBlanksAs', NS)
    text = blanks.get('val') if blanks is not None else None
    result.blanks = 2 if text is None else {spelled: value for value, spelled in BLANKS.items()}.get(text, 0)
    flags = (flag('c:chart/c:plotVisOnly', True, NS), flag('c:chart/c:showDLblsOverMax', True, NS),
             flag('c:chart/c:extLst/c:ext/c16r3:dataDisplayOptions16/c16r3:dispNaAsBlank', False, NS | {'c16r3': C16R3}))
    if text not in (None, *BLANKS.values()) or None in flags:
        result.display_readable = False
        return
    result.visible_only, result.over_maximum, result.not_available_as_blank = (bool(value) for value in flags)

def read_chart(xml: str, part: str) -> ChartState:
    root = xml_tree(xml)
    result = ChartState(part=part, original=xml, dirty=False, chart_type=0)
    title = root.find('c:chart/c:title', NS)
    result.title = ''.join(title.itertext()) if title is not None else None
    if title is not None:
        typed = ''.join(node.text or '' for node in title.findall('.//a:t', NS))
        rich = title.find('c:tx/c:rich', NS) is not None
        result.title, result.title_typed = typed if rich else 'Chart Title', rich
        result.title_look = read_look(title, NS)
    deleted = root.find('c:chart/c:autoTitleDeleted', NS)
    result.title_deleted = title is None and deleted is not None and deleted.get('val', '1') in ('1', 'true')
    legend = root.find('c:chart/c:legend', NS)
    result.legend = legend is not None
    if legend is not None:
        position = legend.find('c:legendPos', NS)
        if position is not None and position.get('val', 'r') in ('b', 't', 'l', 'r', 'tr'):
            result.legend_position = position.get('val', 'r')
        overlay = legend.find('c:overlay', NS)
        result.legend_overlay = overlay is not None and overlay.get('val', '1') in ('1', 'true')
        effects = legend.find('c:spPr/a:effectLst', NS)
        if effects is not None:
            result.legend_shadow = effects.find('a:outerShdw', NS) is not None
        for entry in legend.findall('c:legendEntry', NS):
            index, deleted = entry.find('c:idx', NS), entry.find('c:delete', NS)
            if (index is not None and re.fullmatch(r'[0-9]{1,6}', index.get('val', '')) and deleted is not None
                    and deleted.get('val', '1') in ('1', 'true')):
                result.legend_deleted.append(int(index.get('val', '')))
    read_display(result, root)
    plot = root.find('c:chart/c:plotArea', NS)
    if plot is None:
        return result
    result.data_table = plot.find('c:dTable', NS) is not None
    plots = [entry for entry in plot if entry.tag.rpartition('}')[2].endswith('Chart')]
    if len(plots) != 1:
        return result
    entry = plots[0]
    kind = entry.tag.rpartition('}')[2]
    _read_group(result, entry)
    grouping = entry.find('c:grouping', NS)
    group = grouping.attrib.get('val', '') if grouping is not None else ''
    # Each kind's types by how it groups its series: clustered or standard, stacked, 100% stacked
    # (tests/fixtures/chart_stacked.json).
    stacking = {'clustered': 0, 'standard': 0, 'stacked': 1, 'percentStacked': 2}.get(group)
    if kind == 'barChart' and stacking is not None and group != 'standard':
        direction = entry.find('c:barDir', NS)
        across = direction is not None and direction.attrib.get('val') == 'bar'
        result.chart_type = ((57, 58, 59) if across else (51, 52, 53))[stacking]
    elif kind == 'lineChart' and stacking is not None and group != 'clustered':
        marker = entry.find('c:marker', NS)
        marked = marker is not None and marker.attrib.get('val') == '1'
        result.chart_type = ((65, 66, 67) if marked else (4, 63, 64))[stacking]
    elif kind == 'areaChart' and group in ('stacked', 'percentStacked'):
        result.chart_type = 76 if group == 'stacked' else 77
    elif kind == 'scatterChart':
        style = entry.find('c:scatterStyle', NS)
        scatter = style.attrib.get('val', '') if style is not None else ''
        series_nodes = entry.findall('c:ser', NS)
        no_lines = bool(series_nodes) and all(node.find('c:spPr/a:ln/a:noFill', NS) is not None
                                            for node in series_nodes)
        result.chart_type = -4169 if scatter == 'marker' or (scatter == 'lineMarker' and no_lines) else 74
    elif kind in ('pieChart', 'doughnutChart') or kind == 'areaChart' and group in ('standard', ''):
        result.chart_type = {'pieChart': 5, 'doughnutChart': -4120, 'areaChart': 1}[kind]
    from pyopenvba.apps.excel._chart_axes_file import read as read_axes

    read_axes(result, root, plot, entry)
    def integer(series: ET.Element, name: str, fallback: int) -> int:
        node = series.find('c:' + name, NS)
        if node is None:
            return fallback
        text = node.get('val', '')
        if not re.fullmatch(r'[0-9]{1,10}', text):
            raise VBAUnsupportedError('Invalid chart series identity/order is not implemented')
        value = int(text)
        if value > 4_294_967_295:
            raise VBAUnsupportedError('Chart series identity/order exceeds the supported integer range')
        return value

    stored_series = entry.findall('c:ser', NS)
    ordered = [(integer(series, 'order', index), integer(series, 'idx', index), series)
               for index, series in enumerate(stored_series)]
    if len({item[0] for item in ordered}) != len(ordered) or len({item[1] for item in ordered}) != len(ordered):
        raise VBAUnsupportedError('Ambiguous chart series identity/order is not implemented')
    for _order, identity, series in sorted(ordered, key=lambda item: item[0]):
        item = SeriesState(series_id=identity)
        item.error_bars = read_error_bars(series, NS)
        item.trendlines = read_trendlines(series, NS)
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
        smooth = series.find('c:smooth', NS)
        item.smooth = smooth is not None and smooth.get('val', '1') in ('1', 'true')
        symbol = series.find('c:marker/c:symbol', NS)
        if symbol is None:
            item.marker = MARKER_AUTOMATIC if kind in ('lineChart', 'scatterChart') else MARKER_NONE
        else:
            spelled = {name: number for number, name in MARKER_SYMBOLS.items() if name}
            item.marker = spelled.get(symbol.get('val', ''), MARKER_UNREAD)
        size = series.find('c:marker/c:size', NS)
        if size is not None and re.fullmatch(r'[0-9]{1,2}', size.get('val', '')) and 2 <= int(size.get('val', '')) <= 72:
            item.marker_size = int(size.get('val', ''))
        look = series.find('c:marker/c:spPr', NS)
        if look is not None:
            item.marker_back = _marker_colour(look)
            line = look.find('a:ln', NS)
            item.marker_fore = _marker_colour(line) if line is not None else None
        labels = series.find('c:dLbls', NS)
        if labels is not None:
            shown = labels.find('c:showVal', NS)
            item.data_labels = shown is not None and shown.get('val', '1') in ('1', 'true')
            leaders = labels.find('c:showLeaderLines', NS)
            item.leader_lines = leaders is not None and leaders.get('val', '1') in ('1', 'true')
        else:
            # A pie's own leader lines, which its series follow until they have labels of their own.
            leaders = entry.find('c:dLbls/c:showLeaderLines', NS)
            if leaders is not None:
                item.leader_lines = leaders.get('val', '1') in ('1', 'true')
        invert = series.find('c:invertIfNegative', NS)
        item.invert_if_negative = invert is not None and invert.get('val', '1') in ('1', 'true')
        explosion = series.find('c:explosion', NS)
        if explosion is not None and re.fullmatch(r'[0-9]{1,9}', explosion.get('val', '')):
            item.explosion = int(explosion.get('val', ''))
        result.series.append(item)
    if result.series and result.series[0].value_reference:
        try:
            area = parse_reference(result.series[0].value_reference)[0]
            if area.top == area.bottom and area.left != area.right:
                result.plot_by = 1
        except ValueError:
            pass
    return result

def _marker_colour(node: ET.Element) -> str | None:
    """A marker's fill or line colour as SeriesState keeps it: RRGGBB, "none", None when the element names none, or
    "?" for a fill the model does not read."""
    fills = [child for child in node if child.tag.rpartition('}')[2].endswith('Fill')]
    if not fills:
        return None
    kind = fills[0].tag.rpartition('}')[2]
    if kind == 'noFill':
        return 'none'
    colour = fills[0].find('a:srgbClr', NS) if kind == 'solidFill' else None
    value = colour.get('val', '') if colour is not None else ''
    return value.upper() if colour is not None and len(colour) == 0 and re.fullmatch(r'[0-9A-Fa-f]{6}', value) else '?'


def _data(node: ET.Element | None) -> tuple[str, list[object]]:
    if node is None:
        return '', []
    formula = node.find('.//c:f', NS)
    if formula is not None:
        return formula.text or '', []
    numeric = node.find('c:numLit', NS) is not None
    values: list[object] = []
    for point in node.findall('.//c:pt', NS):
        try:
            index = int(point.attrib.get('idx', '0'))
        except ValueError:
            raise VBAUnsupportedError('Invalid chart point indexes are not implemented') from None
        if not 0 <= index < MAX_CHART_POINTS:
            raise VBAUnsupportedError('Chart point index exceeds the supported allocation budget')
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
    if isinstance(value, str):
        # A cell of text plots, and is cached, as 0 (tests/fixtures/chart_multilevel_categories.json).
        return 0.0
    return float(to_number(value))

def _read_group(result: ChartState, group: ET.Element) -> None:
    """The chart group's settings as a loaded part holds them (tests/fixtures/chart_groups.json)."""
    def number(name: str) -> int | None:
        node = group.find('c:' + name, NS)
        text = node.get('val', '') if node is not None else ''
        return int(text) if re.fullmatch(r'-?[0-9]{1,3}', text) else None

    vary = group.find('c:varyColors', NS)
    if vary is not None:
        result.vary = vary.get('val', '1') in ('1', 'true')
    gap, overlap, first, hole = number('gapWidth'), number('overlap'), number('firstSliceAng'), number('holeSize')
    result.gap_width = gap if gap is not None else result.gap_width
    result.overlap = overlap if overlap is not None else result.overlap
    result.first_slice = first if first is not None else result.first_slice
    result.hole = hole


def _grouped(xml: str, state: ChartState, kind: str) -> str:
    """The chart group's settings in the part (tests/fixtures/chart_groups.json): varyColors once set, a bar or column
    chart's gapWidth with any overlap after it, a pie's or doughnut's firstSliceAng, and a doughnut's holeSize. A kind
    that draws none of these keeps what was set and writes nothing."""
    if state.vary is not None:
        xml = re.sub(r'<c:varyColors val="[01]"/>', f'<c:varyColors val="{int(state.vary)}"/>', xml, count=1)
    overlap = f'<c:overlap val="{state.overlap}"/>' if state.overlap and kind in ('column', 'bar') else ''
    xml = xml.replace('<c:gapWidth val="150"/>', f'<c:gapWidth val="{state.gap_width}"/>{overlap}', 1)
    xml = xml.replace('<c:firstSliceAng val="0"/>', f'<c:firstSliceAng val="{state.first_slice}"/>', 1)
    if state.hole is not None:
        xml = xml.replace('<c:holeSize val="50"/>', f'<c:holeSize val="{state.hole}"/>', 1)
    return xml


def dated_categories(sheet: Worksheet, state: ChartState) -> bool:
    """Whether the chart's categories are dates, which gives an automatic category axis a time scale: the first
    series' categories are cells formatted as dates (tests/fixtures/chart_category_type.json)."""
    from pyopenvba.apps.excel._typing import is_date_format

    first = next((entry for entry in state.series if entry.category_reference), None)
    # Categories of several levels keep a text axis, dates among them or not (tests/fixtures/
    # chart_multilevel_categories.json).
    return (first is not None and category_levels(sheet, first) is None
            and is_date_format(source_format(sheet, first.category_reference)))


def source_format(sheet: Worksheet, formula: str) -> str:
    from pyopenvba.apps.excel._model import Range

    if not formula:
        return 'General'
    area = parse_reference(formula)[0]
    owner = sheet.book.sheet_named(area.sheet) if area.sheet else sheet
    from pyopenvba._a1 import Area

    return to_text(Range(owner, [Area(area.top, area.left, area.top, area.left)]).NumberFormat())

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
                wanted_code = source_format(sheet, unescape(formula.group(1))) if numeric else 'General'
            except (ValueError, KeyError, VBAUnsupportedError, VBARuntimeError):
                # Unknown/external chart reference expressions retain their original caches.
                return raw
            root = xml_tree(f'<root xmlns:c="{NS["c"]}">{cached.group(0)}</root>')
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
            code = root.findtext('.//c:formatCode', default='General', namespaces=NS)
            if same and (not numeric or code == wanted_code):
                return raw
            replacement = cache_markup(values, numeric=numeric, code=wanted_code)
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
        levels = category_levels(sheet, item) or []
        # Blanks leave categories numbers, all of them blank too (tests/fixtures/chart_multilevel_categories.json).
        numeric = not levels and bool(categories) and all(
            value is EMPTY or isinstance(value, (int, float)) and not isinstance(value, bool) for value in categories)
        name_points = [text for row in shown_cells(sheet, item.name_reference) for text in row] if (
            item.name_reference and len(referenced_values(sheet, item.name_reference)) > 1) else []
        data.append(SeriesData(values=item.value_reference, value_cache=[_number(value) for value in series.values()],
                               value_format=source_format(sheet, item.value_reference),
                               name=item.name_reference or None, name_cache=to_text(series.Name()) if item.name_reference else item.name,
                               name_points=name_points,
                               categories=item.category_reference or ('' if item.categories else None),
                               category_cache=[_number(value) if numeric else to_text(value) for value in categories],
                               category_levels=levels,
                               category_format=source_format(sheet, item.category_reference),
                               numeric_categories=numeric, series_id=item.series_id if item.series_id >= 0 else None,
                               unique_id=item.unique if item.unique >= 0 else None, marker=marker_markup(item),
                               smooth=item.smooth, invert=item.invert_if_negative, explosion=item.explosion,
                               labels=item.data_labels, leader_lines=item.leader_lines,
                               error_bars=error_bars_markup(item.error_bars, CHART_TYPES.get(shown_type(chart.state), ''),
                                                            lambda element, given: _bar_values(sheet, element, given)),
                               trendlines=trendlines_markup(item.trendlines)))
    kind = CHART_TYPES.get(shown_type(chart.state))
    if kind not in CHART_KINDS:
        raise ValueError(f'Chart type {chart.state.chart_type} has no chart writer')
    state = chart.state
    title = ''
    if state.title is not None:
        look = state.title_look
        title = title_element(look, text=state.title if state.title_typed else None,
                              linked_text=linked_title(sheet, look.formula) if look.formula else '')
    xml = chart_part(kind, data, title=title, title_deleted=state.title_deleted,
                     legend=legend_markup(state.legend_position, overlay=state.legend_overlay, shadow=state.legend_shadow,
                                          deleted=state.legend_deleted)
                     if state.legend else "", axis_ids=state.axis_ids, blanks=state.blanks,
                     visible_only=state.visible_only, over_maximum=state.over_maximum,
                     not_available_as_blank=state.not_available_as_blank,
                     grouping=STACKED.get(shown_type(state), ''))
    from pyopenvba.apps.excel._chart_axes_file import generated

    xml = generated(xml, state, lambda formula: linked_title(sheet, formula), dated_categories(sheet, state))
    xml = _grouped(xml, state, kind)
    if data and data[0].numeric_categories and data[0].category_format != 'General':
        # A category axis over numbers shows them in their cells' format, dates as dates
        # (tests/fixtures/chart_category_type.json).
        code = escape_attribute(data[0].category_format)
        xml = re.sub(r'<c:(catAx|dateAx)>.*?</c:\1>', lambda found: found.group(0).replace(
            '<c:numFmt formatCode="General" sourceLinked="1"/>', f'<c:numFmt formatCode="{code}" sourceLinked="1"/>', 1),
            xml, count=1)
    if state.data_table:
        # The data table closes the plot area, after the axes, on every kind (tests/fixtures/chart_series_more.json).
        xml = xml.replace('</c:plotArea>', DATA_TABLE + '</c:plotArea>', 1)
    # Chart builder references become literal caches when the series was assigned an array.
    xml = re.sub(r'<c:numRef><c:f></c:f><c:numCache>(.*?)</c:numCache></c:numRef>', r'<c:numLit>\1</c:numLit>', xml)
    xml = re.sub(r'<c:strRef><c:f></c:f><c:strCache>(.*?)</c:strCache></c:strRef>', r'<c:strLit>\1</c:strLit>', xml)
    # A series with no values yet, as NewSeries makes one, has no values element (tests/fixtures/chart_stacked.json).
    xml = re.sub(r'<c:(val|yVal)><c:numLit><c:formatCode>[^<]*</c:formatCode><c:ptCount val="0"/></c:numLit>'
                 r'</c:\1>', '', xml)
    return xml

def read_error_bars(series: ET.Element, ns: dict[str, str]) -> list[ErrorBarState]:
    """A loaded series' error bars, as far as what they read goes: an imported chart is not changed, so a bar's
    look and a custom bar's values stay in the part as they are (tests/fixtures/chart_error_bars.json)."""
    bars: list[ErrorBarState] = []
    for node in series.findall('c:errBars', ns):
        def value(name: str) -> str:
            found = node.find('c:' + name, ns)
            return found.get('val', '') if found is not None else ''

        amount_text = value('val')
        try:
            amount = float(amount_text) if amount_text else None
        except ValueError:
            amount = None
        bars.append(ErrorBarState(direction=value('errDir') or 'y', include=value('errBarType') or 'both',
                                  kind=value('errValType') or 'fixedVal', amount=amount,
                                  no_cap=value('noEndCap') in ('1', 'true')))
    return bars


def read_trendlines(series: ET.Element, ns: dict[str, str]) -> list[TrendlineState]:
    """A loaded series' trendlines, as far as what they read goes (tests/fixtures/chart_trendlines.json)."""
    lines: list[TrendlineState] = []
    for node in series.findall('c:trendline', ns):
        def value(name: str) -> str:
            found = node.find('c:' + name, ns)
            return found.get('val', '') if found is not None else ''

        def number(name: str) -> float | None:
            try:
                return float(value(name)) if value(name) else None
            except ValueError:
                return None

        kind = value('trendlineType')
        name = node.find('c:name', ns)
        lines.append(TrendlineState(
            kind=kind if kind in TRENDLINE_TYPES.values() else 'linear', order=int(number('order') or 2),
            period=int(number('period') or 2), forward=number('forward') or 0.0, intercept=number('intercept'),
            name=name.text or '' if name is not None else None, equation=value('dispEq') in ('1', 'true'),
            r_squared=value('dispRSqr') in ('1', 'true')))
    return lines


def _bar_values(sheet: Worksheet, element: str, given: tuple[str, list[object]]) -> str:
    """A custom error bar's plus or minus element: a reference with its values cached, or the values given as a
    literal, which markup's last step spells (tests/fixtures/chart_error_bars.json)."""
    formula, values = given
    cached = [_number(value) for value in (referenced_values(sheet, formula) if formula else values)]
    return (f'<c:{element}><c:numRef><c:f>{escape_text(formula)}</c:f>'
            f'{cache_markup(cached, numeric=True, code=source_format(sheet, formula))}</c:numRef></c:{element}>')


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
        generated = markup(sheet, shape.shape_id) if chart.dirty and not chart.original else None
        if not chart.part:
            number = 1
            while package.has(f'xl/charts/chart{number}.xml'):
                number += 1
            chart.part = f'xl/charts/chart{number}.xml'
            identifier = add_relationship(package, rels, REL_CHART, posixpath.relpath(chart.part, posixpath.dirname(drawing)))
            package.write(rels, in_office_order(package.read(rels).decode('utf-8')).encode('utf-8'))
            shape.source = chart_frame(shape_id=shape.shape_id, name=shape.name, relationship=identifier,
                                       left=shape.left, top=shape.top, width=shape.width, height=shape.height,
                                       grid=sheet.drawing_grid())
        if generated is not None:
            package.write(chart.part, generated.encode('utf-8'))
            add_content_type(package, chart.part, CT_CHART)
            chart.dirty = False
    sheet.deleted_chart_parts.clear()
