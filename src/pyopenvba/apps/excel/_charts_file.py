"""Chart parts linked to the worksheet's existing drawing layer.

New-chart markup is adapted from pyOfficeEditor at
4723383e5d3039763e64ecd2a3ee092786c91bb5. Imported chart XML is preserved;
editing imported chart contents is deliberately refused until narrow writers exist.
"""
from __future__ import annotations

import posixpath
import re
import xml.etree.ElementTree as ET
from dataclasses import replace
from pyopenvba.apps.excel._chart_references import NO_ID
from xml.parsers import expat
from typing import TYPE_CHECKING, NoReturn, cast

from pyopenvba._xml import attributes, escape as escape_attribute, escape_text, unescape
from pyopenvba._relationships import in_office_order
from pyopenvba._a1 import parse_reference
from pyopenvba.apps.excel._chart_xml import (BLANKS, CHART_KINDS, DATA_TABLE, DATA_TABLE_WITHOUT_KEYS, LABEL_POSITIONS,
                                             CT_CHART, SeriesData, cache_markup,
                                             chart_frame, chart_part, legend_markup, text_properties)
from pyopenvba.apps.excel._charts import (DATA_TABLE_LOOK, MAX_CHART_POINTS, CHART_TYPES, STACKED, Chart, ChartState, PointState, Series, page_markup,
                                          SeriesState,
                                          MARKER_AUTOMATIC, MARKER_NONE, MARKER_SYMBOLS, MARKER_UNREAD, marker_markup,
                                          categories_gone, category_levels, linked_title, lost, referenced_values,
                                          settle_hidden, shown_cells, shown_type, visible_places)
from pyopenvba.apps.excel._chart_combos import is_combo, regrouped, series_type
from pyopenvba.apps.excel._chart_fonts import read_look as read_font
from pyopenvba.apps.excel._chart_xml import group_markup, plain_label_text
from pyopenvba.apps.excel._chart_title import read_look, title_element
from pyopenvba.apps.excel._chart_error_bars import ErrorBarState, markup as error_bars_markup
from pyopenvba.apps.excel._chart_line_format import group_line_markup
from pyopenvba.apps.excel._chart_trendlines import TYPES as TRENDLINE_TYPES, TrendlineState, markup as trendlines_markup
from pyopenvba.exceptions import VBAUnsupportedError, VBARuntimeError
from pyopenvba.interpreter._values import EMPTY, VBAErrorValue, to_number, to_text
from pyopenvba.powerquery._sheets import add_content_type, add_relationship, drop_content_type

if TYPE_CHECKING:
    from pyopenvba.apps.excel._chart_xml import ChartKind
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

def _axis_elements(plot: ET.Element) -> list[ET.Element]:
    return [child for child in plot if child.tag.rpartition('}')[2] in ('catAx', 'valAx', 'dateAx')]


def _ids(element: ET.Element) -> set[str]:
    return {node.get('val', '') for node in element.findall('c:axId', NS)}


def _group_axis(group: ET.Element, plot: ET.Element) -> int:
    """Which axes a group of a loaded combo chart is on: the part lists the primary axes first."""
    return 1 if _ids(group) == {own for one in _axis_elements(plot)[:2] for own in _ids(one)} else 2


def _combo_groups(plots: list[ET.Element], plot: ET.Element) -> list[ET.Element]:
    """The primary groups first, where every group is of the kinds a combo chart holds, on axes the part defines; else none."""
    axes = _axis_elements(plot)
    kinds = [group.tag.rpartition('}')[2] for group in plots]
    if len(axes) not in (2, 4) or not set(kinds) <= {'barChart', 'lineChart', 'areaChart'}:
        return []
    if any(group.tag.endswith('barChart') and (group.find('c:barDir', NS) is None or group.find('c:barDir', NS).get('val') != 'col') for group in plots):  # type: ignore[union-attr]
        return []
    pairs = [{own for one in axes[at:at + 2] for own in _ids(one)} for at in range(0, len(axes), 2)]
    if any(_ids(group) not in pairs for group in plots):
        return []
    found = sorted(plots, key=lambda group: _group_axis(group, plot))
    return found if _group_axis(found[0], plot) == 1 else []


def _type_of(group: ET.Element, marked: bool) -> int:
    from pyopenvba.apps.excel._chart_combos import own_type_of

    grouping = group.find('c:grouping', NS)
    return own_type_of(group.tag.rpartition('}')[2], grouping.get('val', '') if grouping is not None else '', marked)


def own_group(item: SeriesState, group: ET.Element, plot: ET.Element) -> None:
    """A loaded series' type and axes, from the group it is in."""
    item.own_type = _type_of(group, item.marker != MARKER_NONE)
    item.axis_group = _group_axis(group, plot)


def read_group_settings(result: ChartState, plots: list[ET.Element], plot: ET.Element) -> None:
    from pyopenvba.apps.excel._chart_combos import family_of_type, settle

    for group in plots:
        held = ChartState()
        _read_group(held, group)
        family = family_of_type(_type_of(group, False))
        result.group_settings[_group_axis(group, plot), family] = {
            'gap_width': held.gap_width, 'overlap': held.overlap, 'drop_lines': held.drop_lines, 'hi_lo_lines': held.hi_lo_lines,
            'up_down_bars': held.up_down_bars, 'series_lines': held.series_lines, 'vary': held.vary}
    settle(result)


def read_secondary_axes(result: ChartState, root: ET.Element, plot: ET.Element, plots: list[ET.Element]) -> None:
    """The secondary axes of a loaded combo chart, read as the primary are."""
    from pyopenvba.apps.excel._chart_axes_file import read as read_axes

    second = next((group for group in plots if _group_axis(group, plot) == 2), None)
    if second is None:
        return
    held = ChartState(chart_type=result.chart_type)
    read_axes(held, root, plot, second)
    if not held.axes_readable or not result.axes_readable:
        result.axes_readable = False
        return
    for entry in held.axes.values():
        entry.group = 2
    category = next(one for one in _axis_elements(plot)[2:] if one.tag.endswith('catAx') or one.tag.endswith('dateAx'))
    # A value axis left to cross at the far side says so on a category axis that shows, or nothing where it does not.
    crossing = category.find('c:crosses', NS)
    held.axes[2].far_side = crossing is None or (crossing.get('val') == 'max' and held.axes[1].present)
    result.secondary_axes = held.axes


def _read_fonts(result: ChartState, root: ET.Element) -> None:
    """What the part says of each element's text, as Font reads it (tests/fixtures/chart_fonts_saved.json)."""
    found = {'area': root, 'title': root.find('c:chart/c:title', NS), 'legend': root.find('c:chart/c:legend', NS)}
    plot = root.find('c:chart/c:plotArea', NS)
    if plot is not None:
        table = plot.find('c:dTable', NS)
        found['dataTable'] = table
        axes = _axis_elements(plot)
        for element in axes[:2]:
            kind = 2 if element.tag.endswith('valAx') else 1
            found[f'tick{kind}'] = element
            found[f'title{kind}'] = element.find('c:title', NS)
        for element in axes[2:]:
            kind = 2 if element.tag.endswith('valAx') else 1
            found[f'tick{kind}s'] = element
            found[f'title{kind}s'] = element.find('c:title', NS)
    for key, owner in found.items():
        look = None if owner is None else read_font(owner, NS)
        if look is not None:
            result.element_fonts[key] = look


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
    # A chart without the flag has had its title deleted, as Excel takes it (tests/fixtures/chart_loaded_title.json).
    result.title_deleted = title is None and (deleted is None or deleted.get('val', '1') in ('1', 'true'))
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
    _read_fonts(result, root)
    # A part that leaves roundedCorners out has rounded corners, as the file format's default says.
    rounded = root.find('c:roundedCorners', NS)
    result.rounded = rounded is None or rounded.get('val', '1') in ('1', 'true')
    effects = root.find('c:spPr/a:effectLst', NS)
    result.shadow = None if effects is None else effects.find('a:outerShdw', NS) is not None
    plot = root.find('c:chart/c:plotArea', NS)
    if plot is None:
        return result
    result.data_table = plot.find('c:dTable', NS) is not None
    plots = [entry for entry in plot if entry.tag.rpartition('}')[2].endswith('Chart')]
    if not plots or (len(plots) != 1 and not _combo_groups(plots, plot)):
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

    first_primary = _combo_groups(plots, plot)[0] if len(plots) > 1 else entry
    read_axes(result, root, plot, first_primary)
    if len(plots) > 1:
        read_secondary_axes(result, root, plot, plots)
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

    owners = {id(series): group for group in plots for series in group.findall('c:ser', NS)}
    stored_series = [series for group in plots for series in group.findall('c:ser', NS)]
    ordered = [(integer(series, 'order', index), integer(series, 'idx', index), series)
               for index, series in enumerate(stored_series)]
    if len({item[0] for item in ordered}) != len(ordered) or len({item[1] for item in ordered}) != len(ordered):
        raise VBAUnsupportedError('Ambiguous chart series identity/order is not implemented')
    for _order, identity, series in sorted(ordered, key=lambda item: item[0]):
        item = SeriesState(series_id=identity)
        item.error_bars = read_error_bars(series, NS)
        item.trendlines = read_trendlines(series, NS)
        name_ref = series.find('c:tx/c:strRef/c:f', NS)
        item.name_reference = _as_excel_reads(name_ref.text or '') if name_ref is not None else ''
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
        # A line left without it means its default, on (tests/fixtures/tool_drawings.json).
        item.smooth = (smooth.get('val', '1') in ('1', 'true') if smooth is not None
                       else owners[id(series)].tag.endswith('}lineChart'))
        symbol = series.find('c:marker/c:symbol', NS)
        owner = owners[id(series)]
        owner_kind = owner.tag.rpartition('}')[2]
        if symbol is None:
            item.marker = MARKER_AUTOMATIC if owner_kind in ('lineChart', 'scatterChart') else MARKER_NONE
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
            item.label_look.font = read_font(labels, NS)
            shown = labels.find('c:showVal', NS)
            item.data_labels = shown is not None and shown.get('val', '1') in ('1', 'true')
            leaders = labels.find('c:showLeaderLines', NS)
            item.leader_lines = leaders is not None and leaders.get('val', '1') in ('1', 'true')
        else:
            # A pie's own leader lines, which its series follow until they have labels of their own.
            leaders = owner.find('c:dLbls/c:showLeaderLines', NS)
            if leaders is not None:
                item.leader_lines = leaders.get('val', '1') in ('1', 'true')
        invert = series.find('c:invertIfNegative', NS)
        # Excel always writes it on bars; another writer leaving it out means its default, on (tests/fixtures/tool_drawings.json).
        item.invert_if_negative = (invert.get('val', '1') in ('1', 'true') if invert is not None
                                   else owner_kind in ('barChart', 'bar3DChart'))
        explosion = series.find('c:explosion', NS)
        if explosion is not None and re.fullmatch(r'[0-9]{1,9}', explosion.get('val', '')):
            item.explosion = int(explosion.get('val', ''))
        if len(plots) > 1:
            own_group(item, owner, plot)
        result.series.append(item)
    if len(plots) > 1:
        read_group_settings(result, plots, plot)
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


def _as_excel_reads(reference: str) -> str:
    """One block of a sheet's cells as Excel reads it from a file spelled otherwise: absolute, and its sheet bare where a plain name
    needs no apostrophes (tests/fixtures/tool_drawings.json). Anything else stays as written."""
    found = re.fullmatch(r"(?:'([A-Za-z_][\w.]*)'|([^'!]+)|('(?:[^']|'')+'))!\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?", reference)
    if found is None:
        return reference
    plain, bare, needing, column, row, far_column, far_row = found.groups()
    sheet = plain if plain is not None and not re.fullmatch(r'[A-Za-z]{1,3}\d+', plain) else needing or bare or f"'{plain}'"
    return f'{sheet}!${column}${row}' + (f':${far_column}${far_row}' if far_column else '')


def _data(node: ET.Element | None) -> tuple[str, list[object]]:
    if node is None:
        return '', []
    formula = node.find('.//c:f', NS)
    if formula is not None:
        return _as_excel_reads(formula.text or ''), []
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

    # Left out it means its default, on (tests/fixtures/tool_drawings.json).
    vary = group.find('c:varyColors', NS)
    result.vary = vary is None or vary.get('val', '1') in ('1', 'true')
    gap, overlap, first, hole = number('gapWidth'), number('overlap'), number('firstSliceAng'), number('holeSize')
    result.gap_width = gap if gap is not None else result.gap_width
    result.overlap = overlap if overlap is not None else result.overlap
    result.first_slice = first if first is not None else result.first_slice
    result.hole = hole
    result.drop_lines = group.find('c:dropLines', NS) is not None
    result.hi_lo_lines = group.find('c:hiLowLines', NS) is not None
    result.up_down_bars = group.find('c:upDownBars', NS) is not None
    result.series_lines = group.find('c:serLines', NS) is not None


_NO_LINE = '<c:spPr><a:ln w="38100"><a:noFill/></a:ln></c:spPr>'


def shadowed(look: str, shadow: bool | None) -> str:
    """A series' spPr with its Shadow, once that was set: the same effect as a chart's, or none spelled out
    (tests/fixtures/chart_members_more.json)."""
    if shadow is None:
        return look
    effect = (CHART_SHADOW.removeprefix('<c:spPr>').removesuffix('</c:spPr>') if shadow else '<a:effectLst/>')
    if not look:
        return f'<c:spPr>{effect}</c:spPr>'
    if '<a:effectLst' in look:
        return re.sub(r'<a:effectLst(?:/>|>.*?</a:effectLst>)', lambda _: effect, look, count=1)
    place = look.find('<a:extLst>')
    return look[:place] + effect + look[place:] if place >= 0 else look.replace('</c:spPr>', effect + '</c:spPr>', 1)


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
    # Drop lines in line and area charts, high-low lines and up-down bars in line charts, series lines in stacked
    # column and bar charts, each just before the group's own marker and smooth flags, or its axes
    # (tests/fixtures/chart_group_lines.json).
    lined = kind in ('line', 'lineMarkers')
    def element(tag: str) -> str:
        look = group_line_markup(tag, state.element_looks.get(tag))
        return f'<c:{tag}>{look}</c:{tag}>' if look else f'<c:{tag}/>'

    extra = element('dropLines') if state.drop_lines and (lined or kind == 'area') else ''
    if lined:
        extra += element('hiLowLines') if state.hi_lo_lines else ''
        extra += (f'<c:upDownBars><c:gapWidth val="150"/>{element("upBars")}{element("downBars")}</c:upDownBars>'
                  if state.up_down_bars else '')
    if kind in ('column', 'bar') and shown_type(state) in STACKED and state.series_lines:
        extra += element('serLines')
    if extra:
        if lined:
            xml = re.sub(r'(?:<c:marker val="1"/>)?<c:smooth val="0"/>(?=<c:axId)', lambda found: extra + found.group(0),
                         xml, count=1)
        else:
            xml = xml.replace('<c:axId', extra + '<c:axId', 1)
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

    if not formula or lost(formula):
        return 'General'
    area = parse_reference(formula)[0]
    owner = sheet.book.sheet_named(area.sheet) if area.sheet else sheet
    from pyopenvba._a1 import Area

    return to_text(Range(owner, [Area(area.top, area.left, area.top, area.left)]).NumberFormat())

def refresh(sheet: Worksheet, package: OpcFile) -> None:
    """Refresh only changed reference caches; leave every other chart byte alone, but for the references a chart read
    from a file has had rewritten as its cells moved."""
    for chart in sheet.chart_states.values():
        if chart.dirty and chart.original and chart.part and package.has(chart.part):
            package.write(chart.part, chart.original.encode('utf-8'))
            chart.dirty = False
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
            if lost(unescape(formula.group(1))):
                return raw
            try:
                source = referenced_values(sheet, unescape(formula.group(1)), visible=chart.visible_only)
                if not source:
                    return raw
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

def numeric_categories(categories: list[object], levels: list[list[str]]) -> bool:
    """Whether a series' categories are numbers, which blanks leave them, all of them blank too
    (tests/fixtures/chart_multilevel_categories.json)."""
    return not levels and bool(categories) and all(
        value is EMPTY or isinstance(value, (int, float)) and not isinstance(value, bool) for value in categories)


def linked_formats(chart: Chart) -> tuple[str, str]:
    """The number formats the category and value axes show while linked to their source: a category axis over numbers
    takes the first category cell's, otherwise General, and the value axis the first series' first value cell's
    (tests/fixtures/chart_category_type.json, tests/fixtures/chart_axis_parts.json)."""
    if not chart.state.series:
        return 'General', 'General'
    item = chart.state.series[0]
    categories = Series(chart, item).categories()
    numeric = numeric_categories(categories, category_levels(chart.sheet, item) or [])
    return (source_format(chart.sheet, item.category_reference) if numeric else 'General',
            source_format(chart.sheet, item.value_reference))


#: The extension each kind of chart keeps its filtered series under (tests/fixtures/series_filter.json).
_FILTERED = {'column': 'filteredBarSeries', 'bar': 'filteredBarSeries', 'line': 'filteredLineSeries',
             'lineMarkers': 'filteredLineSeries', 'scatter': 'filteredScatterSeries', 'area': 'filteredAreaSeries',
             'pie': 'filteredPieSeries', 'doughnut': 'filteredPieSeries'}
_C15 = '{02D57815-91ED-43cb-92C2-25804820EDAC}'


def _with_filtered_series(xml: str, kind: object, places: list[int]) -> str:
    """The chart with these series moved from its group to an extension at the group's end, their references spelled as
    such an extension spells them."""
    found = list(re.finditer(r'<c:ser>.*?</c:ser>', xml, re.DOTALL))
    moved = ''
    name = _FILTERED[str(kind)]
    space = 'xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"'
    for place in sorted(places, reverse=True):
        match = found[place]
        xml = xml[:match.start()] + xml[match.end():]
    for turn, place in enumerate(places):
        # Each has a wrapper of its own, and past the first its references name the namespace again.
        again = f' {space}' if turn else ''
        inner = found[place].group(0)[len('<c:ser>'):-len('</c:ser>')]
        inner = inner.replace(f' {space}', again)
        inner = re.sub(r'<c:f>(.+?)</c:f>', lambda hit: (
            f'<c:extLst><c:ext uri="{_C15}"{again}><c15:formulaRef><c15:sqref>{hit.group(1)}</c15:sqref>'
            '</c15:formulaRef></c:ext></c:extLst>'), inner)
        moved += f'<c15:{name}><c15:ser>{inner}</c15:ser></c15:{name}>'
    extension = f'<c:extLst><c:ext uri="{_C15}" {space}>{moved}</c:ext></c:extLst>'
    close = re.search(r'</c:(?:bar|line|scatter|area|pie|doughnut)Chart>', xml)
    assert close is not None
    xml = xml[:close.start()] + extension + xml[close.start():]
    # With nothing drawn the axes have no number format to follow.
    return xml if '<c:ser>' in xml else xml.replace('<c:numFmt formatCode="General" sourceLinked="1"/>', '')


_LOST_NAME = '<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>#REF!</c:v></c:pt></c:strCache>'


def _filtered_title(kind: str, element: str, reference: str, cache: str) -> str:
    """The extension that keeps a reference taken away from all a chart's series, as Excel writes it."""
    return (f'<c:ext uri="{_C15}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"><c15:{kind}>'
            f'<c15:{element}><c:strRef><c:extLst><c:ext uri="{_C15}"><c15:formulaRef><c15:sqref>'
            f'{escape_text(reference)}</c15:sqref></c15:formulaRef></c:ext></c:extLst>{cache}</c:strRef>'
            f'</c15:{element}></c15:{kind}></c:ext>')


def _all_numbers(sheet: Worksheet, reference: str) -> bool:
    return not lost(reference) and all(isinstance(value, (int, float)) and not isinstance(value, bool)
                                       for value in referenced_values(sheet, reference))


def _shown_points(sheet: Worksheet, state: ChartState, item: SeriesState) -> SeriesState:
    """The series with its points numbered among the cells shown, as the file numbers them; those hidden wait."""
    if not state.visible_only or not item.points or not item.value_reference or lost(item.value_reference):
        return item
    places = {place: index for index, place in enumerate(visible_places(sheet, item.value_reference, True))}
    def numbered(place: int, point: PointState) -> PointState:
        # Moved to another place it loses its unique ids, as moved cells do.
        if places[place] in (place, point.drawn_at):
            return point
        return replace(point, unique=-1, label_unique=NO_ID if point.label_unique >= 0 else point.label_unique)

    return replace(item, points={places[place]: numbered(place, point) for place, point in item.points.items()
                                 if place in places})


def _with_series_out_of_sight(xml: str, state: ChartState) -> str:
    """A series with its cells all hidden is in the file with references that hold nothing, and without its own look,
    which it has again once shown."""
    found = list(re.finditer(r'<c:ser>.*?</c:ser>', xml, re.DOTALL))
    for match, entry in reversed(list(zip(found, state.full_series or state.series, strict=False))):
        if any(entry is hidden for hidden in state.out_of_sight):
            series = re.sub(r'<c:numCache>(?:(?!<c:numCache>).)*?</c:numCache></c:numRef></c:(val|yVal)>',
                            r'</c:numRef></c:\1>', match.group(0), flags=re.DOTALL)
            if entry.format_look is not None:
                series = re.sub(r'(</c:tx>|<c:order val="[0-9]+"/>)<c:spPr>.*?</c:spPr>', r'\1', series, count=1,
                                flags=re.DOTALL)
            xml = xml[:match.start()] + series + xml[match.end():]
    # With nothing drawn the axes have no number format to follow.
    return xml if state.series else xml.replace('<c:numFmt formatCode="General" sourceLinked="1"/>', '')


def _with_lost_references(xml: str, sheet: Worksheet, state: ChartState) -> str:
    """The chart with its series' lost names and categories as Excel spells them, which depends on the others: a name
    lost from every series leaves the series for an extension, and categories too except on a scatter chart, which has
    no such extension; lost from only some, they stay as references to nothing
    (tests/fixtures/chart_lost_cells.json)."""
    every = state.full_series or state.series
    names_gone = state.names_lost()
    # With no series drawn their categories stay too.
    all_gone = bool(state.series) and all(categories_gone(sheet, state, entry) for entry in every)
    found = list(re.finditer(r'<c:ser>.*?</c:ser>', xml, re.DOTALL))
    for match, entry in reversed(list(zip(found, every, strict=False))):
        series = match.group(0)
        added = ''
        if names_gone:
            series = re.sub(r'<c:tx>.*?</c:tx>', '', series, count=1, flags=re.DOTALL)
            added += _filtered_title('filteredSeriesTitle', 'tx', entry.name_reference, _LOST_NAME)
        if categories_gone(sheet, state, entry):
            kind = 'numRef' if entry.lost_numbers or _all_numbers(sheet, entry.category_reference) else 'strRef'
            spelled = f'<c:{kind}><c:f>{escape_text(entry.category_reference)}</c:f></c:{kind}>'
            scatter = '<c:xVal>' in series
            if all_gone and not scatter:
                series = re.sub(r'<c:cat>.*?</c:cat>', '', series, count=1, flags=re.DOTALL)
                added += _filtered_title('filteredCategoryTitle', 'cat', entry.category_reference, '')
            else:
                series = re.sub(r'<c:(cat|xVal)>.*?</c:\1>',
                                lambda hit: f'<c:{hit.group(1)}>{spelled}</c:{hit.group(1)}>', series, count=1,
                                flags=re.DOTALL)
        if added:
            at = series.rindex('<c:extLst>') + len('<c:extLst>') if '<c:extLst>' in series else -1
            series = (series[:at] + added + series[at:] if at >= 0
                      else series.replace('</c:ser>', f'<c:extLst>{added}</c:extLst></c:ser>'))
        xml = xml[:match.start()] + series + xml[match.end():]
    if every and categories_gone(sheet, state, every[0]) and state.series and '<c:xVal>' not in xml:
        # The axis follows the first series', and without them has no number format to follow.
        xml = re.sub(r'<c:catAx>.*?</c:catAx>', lambda hit: hit.group(0).replace(
            '<c:numFmt formatCode="General" sourceLinked="1"/>', ''), xml, count=1, flags=re.DOTALL)
    return xml


def _with_lost_values(xml: str, state: ChartState) -> str:
    """An area chart gives the area of a series with its values gone no line (tests/fixtures/chart_lost_cells.json)."""
    found = list(re.finditer(r'<c:ser>.*?</c:ser>', xml, re.DOTALL))
    for match, entry in reversed(list(zip(found, state.full_series or state.series, strict=False))):
        if lost(entry.value_reference):
            series = re.sub(r'(<c:order val="[0-9]+"/>(?:<c:tx>.*?</c:tx>)?)',
                            r'\1<c:spPr><a:ln w="25400"><a:noFill/></a:ln></c:spPr>', match.group(0), count=1,
                            flags=re.DOTALL)
            xml = xml[:match.start()] + series + xml[match.end():]
    return xml


def _with_element_labels(xml: str, shown: bool, place: int | None) -> str:
    """The chart's own c:dLbls, after its last series, as SetElement leaves it
    (tests/fixtures/chart_set_element.json)."""
    at = xml.rindex('</c:ser>')
    tail = xml[at:]
    if shown:
        position = '' if place is None else f'<c:dLblPos val="{LABEL_POSITIONS[place]}"/>'
        tail = tail.replace('<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/>',
                            f'<c:dLbls>{position}<c:showLegendKey val="0"/><c:showVal val="1"/>', 1)
    tail = tail.replace('<c:showLeaderLines val="0"/></c:dLbls>', '<c:showLeaderLines val="1"/></c:dLbls>', 1)
    return xml[:at] + tail


def markup(sheet: Worksheet, shape_id: int) -> str:
    shape = next(item for item in sheet.shapes_ if item.shape_id == shape_id)
    from pyopenvba.apps.excel._chart_format import element_markup, markup as series_look
    from pyopenvba.apps.excel._chart_points import point_labels_markup, points_markup
    from pyopenvba.apps.excel._chart_styled import dressed, is_styled

    chart = sheet.chart_sheet or Chart(sheet, shape)
    settle_hidden(sheet, chart.state)
    data: list[SeriesData] = []
    for item in chart.state.full_series or chart.state.series:
        series = Series(chart, item)
        categories = series.categories()
        levels = category_levels(sheet, item) or []
        numeric = numeric_categories(categories, levels)
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
                               kind=cast('ChartKind', CHART_TYPES[series_type(chart.state, item)]) if is_combo(chart.state) else None,
                               numeric_categories=numeric, series_id=item.series_id if item.series_id >= 0 else None,
                               unique_id=item.unique if item.unique >= 0 else None, marker=marker_markup(item),
                               smooth=item.smooth, invert=item.invert_if_negative, explosion=item.explosion,
                               scatter_line=item.scatter_line,
                               look=shadowed(series_look(item, CHART_TYPES.get(series_type(chart.state, item), ''), series.position(),
                                                         sheet.book.stylesheet.colors.theme) or
                                             (_NO_LINE if chart.state.series and item.shadow is not None and not item.scatter_line
                                              and CHART_TYPES.get(shown_type(chart.state)) == 'scatter' else ''),
                                             item.shadow),
                               own_fill=item.format_look is not None and item.format_look.fill_rgb is not None,
                               points=points_markup(_shown_points(sheet, chart.state, item),
                                                    CHART_TYPES.get(series_type(chart.state, item), ''),
                                                    series.position(), sheet.book.stylesheet.colors.theme),
                               point_labels=point_labels_markup(_shown_points(sheet, chart.state, item)),
                               labels=item.data_labels, leader_lines=item.leader_lines, label_look=item.label_look,
                               error_bars=error_bars_markup(item.error_bars, CHART_TYPES.get(series_type(chart.state, item), ''),
                                                            lambda element, given: _bar_values(sheet, element, given)),
                               trendlines=trendlines_markup(item.trendlines)))
    kind = CHART_TYPES.get(shown_type(chart.state))
    if kind not in CHART_KINDS:
        raise ValueError(f'Chart type {chart.state.chart_type} has no chart writer')
    state = chart.state
    if (is_combo(state) or state.plot_sequence) and (state.full_series or state.out_of_sight):
        raise VBAUnsupportedError('Saving a combo chart with series filtered or out of sight is not implemented')
    title = ''
    if state.title is not None:
        look = state.title_look
        title = title_element(look, text=state.title if state.title_typed else None,
                              linked_text=linked_title(sheet, look.formula) if look.formula else '',
                              frame=element_markup('title', state.element_looks.get('title')),
                              font=state.element_fonts.get('title'))
    frame = element_markup('legend', state.element_looks.get('legend'))
    if frame and state.legend_shadow is not None:
        raise VBAUnsupportedError('A legend with both a shadow and a Format look is not implemented')
    xml = chart_part(kind, data, title=title, title_deleted=state.title_deleted,
                     legend=legend_markup(state.legend_position, overlay=state.legend_overlay, shadow=state.legend_shadow,
                                          deleted=state.legend_deleted, frame=frame,
                                          entry_text={**{int(key[11:]): '' for key in state.element_looks
                                                         if key.startswith('legendEntry')},
                                                      **{int(key[6:]): text_properties(font)
                                                         for key, font in state.element_fonts.items()
                                                         if key.startswith('legend') and key[6:].isdigit()}},
                                          text=text_properties(state.element_fonts.get('legend')))
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
    if data and data[0].value_format != 'General':
        # The value axis shows its first value cell's format while linked (tests/fixtures/chart_axis_parts.json).
        code = escape_attribute(data[0].value_format)
        identifier = f'<c:axId val="{state.axis_ids[1]}"/>'
        xml = re.sub(r'<c:valAx>.*?</c:valAx>', lambda found: found.group(0).replace(
            '<c:numFmt formatCode="General" sourceLinked="1"/>', f'<c:numFmt formatCode="{code}" sourceLinked="1"/>', 1)
            if found.group(0).startswith('<c:valAx>' + identifier) else found.group(0), xml)
    if state.series and not any(entry.value_reference or entry.values for entry in state.series):
        # Series with nothing to draw leave the value axis no number format to follow
        # (tests/fixtures/chart_auto_titles.json).
        xml = re.sub(r'<c:valAx>.*?</c:valAx>', lambda hit: hit.group(0).replace(
            '<c:numFmt formatCode="General" sourceLinked="1"/>', ''), xml, flags=re.DOTALL)
    if state.out_of_sight:
        xml = _with_series_out_of_sight(xml, state)
    if '<c:areaChart>' in xml:
        xml = _with_lost_values(xml, state)
    if any(lost(entry.name_reference) or categories_gone(sheet, state, entry)
           for entry in state.full_series or state.series):
        xml = _with_lost_references(xml, sheet, state)
    filtered = [index for index, item in enumerate(state.full_series)
                if not any(item is other for other in [*state.series, *state.out_of_sight])]
    if filtered:
        xml = _with_filtered_series(xml, kind, filtered)
    if state.layout_chart_labels and '</c:ser>' in xml:
        at = xml.rindex('</c:ser>')
        xml = xml[:at] + re.sub(r'<c:dLbls>.*?</c:dLbls>', lambda _: state.layout_chart_labels, xml[at:], count=1,
                                flags=re.S)
    if state.layout_category_look and '<c:catAx>' in xml:
        xml = xml.replace('<c:crossAx', state.layout_category_look + '<c:crossAx', 1)
    if state.layout_value_look and '<c:valAx>' in xml:
        at = xml.rindex('<c:valAx>')
        xml = xml[:at] + xml[at:].replace('<c:crossAx', state.layout_value_look + '<c:crossAx', 1)
    if state.chart_style != 2 and not is_styled(state.chart_style):
        xml = xml.replace('<c14:style val="102"/>', f'<c14:style val="{100 + state.chart_style}"/>', 1).replace(
            '<c:style val="2"/>', f'<c:style val="{state.chart_style}"/>', 1)
    if state.element_labels is not None and '</c:ser>' in xml:
        xml = _with_element_labels(xml, *state.element_labels)
    if state.data_table:
        # The data table closes the plot area, after the axes, on every kind (tests/fixtures/chart_series_more.json).
        table = DATA_TABLE if state.data_table_keys else DATA_TABLE_WITHOUT_KEYS
        for tag, shown in zip(('Horz', 'Vert', 'Outline'), state.data_table_borders, strict=True):
            name = f'show{tag}' if tag == 'Outline' else f'show{tag}Border'
            table = table.replace(f'<c:{name} val="1"/>', f'<c:{name} val="{int(shown)}"/>')
        look = element_markup(DATA_TABLE_LOOK, state.element_looks.get(DATA_TABLE_LOOK))
        table = table.replace('</c:dTable>', look + text_properties(state.element_fonts.get('dataTable')) + '</c:dTable>')
        xml = xml.replace('</c:plotArea>', table + '</c:plotArea>', 1)
    if 'area' in state.element_fonts:
        xml = plain_label_text(xml)
    # The plot area's look follows all of that, and the chart area's follows the chart
    # (tests/fixtures/chart_element_format.json).
    xml = xml.replace('</c:plotArea>', element_markup('plot', state.element_looks.get('plot')) + '</c:plotArea>', 1)
    area = element_markup('area', state.element_looks.get('area')) + text_properties(state.element_fonts.get('area'))
    if area and state.shadow is not None:
        raise VBAUnsupportedError('A chart area with both a shadow and a Format look is not implemented')
    xml = xml.replace('</c:chart><c:printSettings>', f'</c:chart>{area}<c:printSettings>', 1)
    if state.rounded:
        xml = xml.replace('<c:roundedCorners val="0"/>', '<c:roundedCorners val="1"/>', 1)
    if any(state.protection.values()):
        # (tests/fixtures/chart_protection.json).
        guards = ''.join(f'<c:{name} val="{int(state.protection.get(name, False))}"/>'
                         for name in ('chartObject', 'data', 'formatting', 'selection', 'userInterface'))
        xml = xml.replace('<c:chart>', f'<c:protection>{guards}</c:protection><c:chart>', 1)
    if state.page:
        xml = re.sub(r'<c:printSettings>.*?</c:printSettings>', lambda _: page_markup(state), xml, count=1, flags=re.DOTALL)
    if state.shadow is not None:
        # The chart area's shadow, as ChartObject.Shadow gives it, and its empty effects once turned off
        # (tests/fixtures/chart_object_reads.json).
        look = CHART_SHADOW if state.shadow else '<c:spPr><a:effectLst/></c:spPr>'
        xml = xml.replace('</c:chart><c:printSettings>', '</c:chart>' + look + '<c:printSettings>', 1)
    if is_combo(state) or (state.plot_sequence and state.series):
        first_second = next((entry for entry, item in zip(data, state.series, strict=True) if item.axis_group == 2), None)
        xml = regrouped(xml, state, group_markup, _grouped, first_second.value_format if first_second is not None else 'General')
    if is_styled(state.chart_style):
        xml = dressed(xml, state, sheet.book.stylesheet.colors.theme)
        if 'area' in state.element_fonts:
            xml = re.sub(r'(<c:dLbls>.*?</c:dLbls>)', lambda hit: hit.group(1).replace(' lIns="38100" tIns="19050" rIns="38100" bIns="19050"', '').replace('><a:spAutoFit/></a:bodyPr>', '/>'), xml, flags=re.S)
    # Chart builder references become literal caches when the series was assigned an array.
    xml = re.sub(r'<c:numRef><c:f></c:f><c:numCache>(.*?)</c:numCache></c:numRef>', r'<c:numLit>\1</c:numLit>', xml)
    xml = re.sub(r'<c:strRef><c:f></c:f><c:strCache>(.*?)</c:strCache></c:strRef>', r'<c:strLit>\1</c:strLit>', xml)
    # A series with no values yet, as NewSeries makes one, has no values element (tests/fixtures/chart_stacked.json).
    xml = re.sub(r'<c:(val|yVal)><c:numLit><c:formatCode>[^<]*</c:formatCode><c:ptCount val="0"/></c:numLit>'
                 r'</c:\1>', '', xml)
    if sheet.chart_sheet is not None:
        # A chart sheet's chart leaves its page to the sheet (tests/fixtures/chart_location.json).
        xml = re.sub(r'<c:printSettings>.*?</c:printSettings>', '', xml, count=1, flags=re.S)
    return xml

CT_CHART_STYLE = 'application/vnd.ms-office.chartstyle+xml'
CT_CHART_COLORS = 'application/vnd.ms-office.chartcolorstyle+xml'
REL_CHART_STYLE = 'http://schemas.microsoft.com/office/2011/relationships/chartStyle'
REL_CHART_COLORS = 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle'

#: The chart area's look once ChartObject.Shadow is on (tests/fixtures/chart_object_reads.json).
CHART_SHADOW = ('<c:spPr><a:effectLst><a:outerShdw dist="35921" dir="2700000" algn="br"><a:srgbClr val="000000"/>'
                '</a:outerShdw></a:effectLst></c:spPr>')


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


def _style_part_names(part: str) -> tuple[str, str]:
    number = part[len('xl/charts/chart'):-len('.xml')]
    return f'xl/charts/style{number}.xml', f'xl/charts/colors{number}.xml'


def _write_style_parts(package: OpcFile, chart: ChartState) -> None:
    """The parts a chart of one of AddChart2's styles brings with it, or none where it has gone back to an old style."""
    from pyopenvba.apps.excel._chart_style_table import COLORS_PART, style_part
    from pyopenvba.apps.excel._chart_styled import is_styled

    if not is_styled(chart.chart_style):
        remove_style_parts(package, chart.part)
        return
    style, colours = _style_part_names(chart.part)
    package.write(style, style_part(chart.chart_style).encode('utf-8'))
    package.write(colours, COLORS_PART.encode('utf-8'))
    add_content_type(package, style, CT_CHART_STYLE)
    add_content_type(package, colours, CT_CHART_COLORS)
    package.write(rels_part(chart.part), (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId2" Type="{REL_CHART_COLORS}" Target="{posixpath.basename(colours)}"/>'
        f'<Relationship Id="rId1" Type="{REL_CHART_STYLE}" Target="{posixpath.basename(style)}"/>'
        '</Relationships>').encode('utf-8'))


def remove_style_parts(package: OpcFile, part: str) -> None:
    for name in _style_part_names(part):
        if package.has(name):
            package.remove(name)
            drop_content_type(package, name)
            package.remove(rels_part(part))


def write(sheet: Worksheet, package: OpcFile, drawing: str) -> None:
    retained = {chart.part for owner in sheet.book.sheets_ for identifier, chart in owner.chart_states.items()
                if any(shape.shape_id == identifier and shape.kind == 'chart' for shape in owner.shapes_)}
    retired = set(sheet.deleted_chart_parts)
    for part in retired - retained:
        remove_style_parts(package, part)
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
            _write_style_parts(package, chart)
            chart.dirty = False
    sheet.deleted_chart_parts.clear()
