"""Chart templates: Chart.SaveChartTemplate and Chart.ApplyChartTemplate, as Excel writes and reads a .crtx
(tests/fixtures/chart_templates_more.json, scripts/measure_chart_templates_more.py).

A template is a small package: the chart's part, its style and colours parts where it has a style, the workbook theme's elements
as a theme override, and a picture of the chart. The chart part is the chart's own markup with every cell reference emptied, a
colour map after its style, a typed chart title's text as "Title", and each spPr, txPr and rich declaring c: and each of their
direct a: children declaring a:. The picture comes from the Export hook (ExcelApplication.set_hook), as Chart.Export's does; it is
the function's own and is not claimed to match Excel's."""
from __future__ import annotations

import re
import zipfile
from io import BytesIO
from pathlib import Path
from typing import TYPE_CHECKING

from pyopenvba.exceptions import VBAUnsupportedError

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts import Chart
    from pyopenvba.apps.excel._model import Workbook

_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
#: The colour map a template's chart, and a chart a template is applied to, holds after its style.
COLOUR_MAP = ('bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" '
              'accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>')
_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
_STYLE_TYPES = ('<Override PartName="/chart/charts/style1.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/>'
                '<Override PartName="/chart/charts/colors1.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>')
_THEME_OVERRIDE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride'


def _content_types(styled: bool) -> str:
    return (_DECLARATION + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="bmp" ContentType="image/bmp"/>'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/chart/chart.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
            + (_STYLE_TYPES if styled else '')
            + '<Override PartName="/chart/theme/themeOverride1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.themeOverride+xml"/></Types>')


_PACKAGE_RELS = (_DECLARATION + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                 'Target="chart/media/image1.bmp"/>'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
                 'Target="chart/chart.xml"/></Relationships>')


def _chart_rels(styled: bool) -> str:
    if not styled:
        return (_DECLARATION + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="rId1" Type="{_THEME_OVERRIDE}" Target="theme/themeOverride1.xml"/></Relationships>')
    return (_DECLARATION + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId3" Type="{_THEME_OVERRIDE}" Target="theme/themeOverride1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" '
            'Target="charts/colors1.xml"/>'
            '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" '
            'Target="charts/style1.xml"/></Relationships>')


def _islands(xml: str) -> str:
    """Each spPr, txPr and rich declaring c:, and each of its direct a: children declaring a:."""
    def one(found: re.Match[str]) -> str:
        tag, inner = found.group(1), found.group(2)
        depth, pieces, at = 0, list[str](), 0
        for token in re.finditer(r'<(/?)(a:\w+)([^>]*?)(/?)>', inner):
            closing, name, rest, empty = token.groups()
            if not closing and depth == 0:
                pieces.append(inner[at:token.start()] + f'<{name} xmlns:a="{_A}"{rest}{empty}>')
                at = token.end()
            if not closing and not empty:
                depth += 1
            elif closing:
                depth -= 1
        pieces.append(inner[at:])
        return f'<c:{tag} xmlns:c="{_C}">' + ''.join(pieces) + f'</c:{tag}>'
    return re.sub(r'<c:(spPr|txPr|rich)>(.*?)</c:\1>', one, xml, flags=re.DOTALL)


def template_chart(xml: str) -> str:
    """A chart's markup as its template holds it."""
    xml = re.sub(r'<c:(strRef|numRef)>.*?</c:\1>', r'<c:\1/>', xml, flags=re.DOTALL)
    title = re.search(r'<c:chart><c:title><c:tx><c:rich>.*?</c:rich></c:tx>', xml, re.DOTALL)
    if title is not None:
        if len(re.findall(r'<a:t>', title.group(0))) != 1:
            raise VBAUnsupportedError('Saving a chart template of a chart whose title is not one run of text is not implemented')
        typed = re.sub(r'<a:t>[^<]*</a:t>', '<a:t>Title</a:t>', title.group(0), count=1)
        xml = xml[:title.start()] + typed + xml[title.end():]
    if '</mc:AlternateContent><c:chart>' not in xml:
        raise VBAUnsupportedError('Saving a chart template of this chart is not implemented')
    xml = xml.replace('</mc:AlternateContent><c:chart>', f'</mc:AlternateContent><c:clrMapOvr xmlns:c="{_C}" {COLOUR_MAP}<c:chart>', 1)
    return _islands(xml)


def book_theme(book: Workbook) -> str:
    """The workbook's theme part."""
    from pyopenvba.apps.excel._io import _template_style_parts  # pyright: ignore[reportPrivateUsage]

    package = book.package
    if package is not None and package.has('xl/theme/theme1.xml'):
        return package.read('xl/theme/theme1.xml').decode('utf-8')
    return _template_style_parts()[1]


def theme_override(theme: str) -> str:
    """The theme's elements as a template's theme override holds them."""
    found = re.search(r'<a:themeElements>(.*)</a:themeElements>', theme, re.DOTALL)
    if found is None:
        raise VBAUnsupportedError("Saving a chart template with this workbook's theme is not implemented")
    return _DECLARATION + f'<a:themeOverride xmlns:a="{_A}">' + found.group(1) + '</a:themeOverride>'


def template_parts(chart: Chart) -> dict[str, str]:
    """The text parts of the chart's template, in the order Excel writes them, the picture aside."""
    from pyopenvba.apps.excel._chart_style_table import COLORS_PART, style_part
    from pyopenvba.apps.excel._chart_styled import is_styled
    from pyopenvba.apps.excel._charts_file import markup

    state = chart.state
    if state.original or chart.sheet.chart_sheet is not None:
        raise VBAUnsupportedError('Saving a chart template of this chart is not implemented')
    styled = is_styled(state.chart_style)
    parts = {'[Content_Types].xml': _content_types(styled), '_rels/.rels': _PACKAGE_RELS,
             'chart/chart.xml': template_chart(markup(chart.sheet, chart.shape.shape_id)), 'chart/_rels/chart.xml.rels': _chart_rels(styled)}
    if styled:
        parts['chart/charts/style1.xml'] = style_part(state.chart_style)
        parts['chart/charts/colors1.xml'] = COLORS_PART
    parts['chart/theme/themeOverride1.xml'] = theme_override(book_theme(chart.sheet.book))
    return parts


#: What a series holds of its data, which a template applied to it leaves as it is; the rest is the template's, its unique id
#: included (tests/fixtures/chart_templates_more.json).
_DATA = ('name', 'name_reference', 'categories', 'values', 'category_reference', 'value_reference', 'lost_numbers', 'series_id')
_UNIQUE_END = re.compile(r'(uniqueId val="\{[0-9A-F]{8})-[0-9A-F-]+')
_AXIS_ID = re.compile(r'(<c:(?:axId|crossAx) val=")\d+')


def _masked(xml: str) -> str:
    return _AXIS_ID.sub(r'\1', _UNIQUE_END.sub(r'\1', xml))


def apply_template(chart: Chart, filename: str) -> None:
    """ApplyChartTemplate: the chart becomes the template's, over its own series' data, with the template's colour map, style and
    colours parts and theme override, and an empty extLst at its chart's end (tests/fixtures/chart_templates_more.json). A title
    the chart showed keeps its text, typed, in the template's look; where it showed none, the template's title is as it was.
    Admitted only where the template is a chart the model writes again exactly: one whose own template is this one."""
    import copy

    from pyopenvba.apps.excel._chart_style_table import COLORS_PART, style_part
    from pyopenvba.apps.excel._charts import ChartTitle
    from pyopenvba.apps.excel._charts_file import markup, read_chart
    from pyopenvba.interpreter._files import native_path
    from pyopenvba.interpreter._values import error, to_text

    book = chart.sheet.book
    book.application.outside('Chart.ApplyChartTemplate')
    state = chart.state
    if state.original or chart.sheet.chart_sheet is not None:
        raise VBAUnsupportedError('Applying a chart template to this chart is not implemented')
    path = Path(native_path(filename))
    try:
        with zipfile.ZipFile(path) as package:
            names = set(package.namelist())
            text = {name: package.read(name).decode('utf-8') for name in names if name.endswith(('.xml', '.rels'))}
    except (OSError, zipfile.BadZipFile) as failed:
        raise error(-2147467259, 'Automation error\nUnspecified error') from failed
    xml, override = text.get('chart/chart.xml'), text.get('chart/theme/themeOverride1.xml')
    if xml is None or override is None:
        raise VBAUnsupportedError('Applying a chart template that is not one Excel saved is not implemented')
    # Read as a chart part spells it: the declarations each spPr, txPr, rich and colour map repeats say nothing more.
    plain = xml.replace(f' xmlns:c="{_C}"', '').replace(f' xmlns:a="{_A}"', '')
    plain = plain.replace('<c:chartSpace', f'<c:chartSpace xmlns:c="{_C}"', 1) if f'xmlns:c="{_C}"' not in plain else plain
    plain = plain.replace('<c:chartSpace', f'<c:chartSpace xmlns:a="{_A}"', 1) if f'xmlns:a="{_A}"' not in plain else plain
    template = read_chart(plain, '')
    style = text.get('chart/charts/style1.xml')
    if style is not None:
        found = re.search(r'<cs:chartStyle\b[^>]*\bid="(\d+)"', style)
        if found is None or style != style_part(int(found.group(1))) or text.get('chart/charts/colors1.xml') != COLORS_PART:
            raise VBAUnsupportedError("Applying a chart template whose style the model does not write is not implemented")
        template.chart_style = int(found.group(1))
    # The looks a chart part holds of its elements and series, read as a loaded chart's are, then held as a new chart's.
    from pyopenvba.apps.excel._chart_loaded_looks import read_looks
    from pyopenvba.powerquery._opc import OpcFile

    template.original, template.style_part = plain, style is not None
    read_looks(template, OpcFile.parse(path.read_bytes()), 'chart/chart.xml')
    template.original, template.style_part, template.looks_read = '', False, set()
    if len(template.series) != len(state.series):
        raise VBAUnsupportedError('Applying a chart template of another number of series is not implemented')
    shown = to_text(ChartTitle(chart).Text()) if state.title is not None else None
    candidate = copy.deepcopy(template)
    for mine, theirs in zip(state.series, candidate.series, strict=True):
        for name in _DATA:
            setattr(theirs, name, getattr(mine, name))
    for (_, axis), identifier in zip(sorted(candidate.axes.items()), state.axis_ids):
        axis.identifier = identifier
    candidate.axis_ids = state.axis_ids
    candidate.part, candidate.dirty, candidate.original = state.part, True, ''
    before = copy.copy(state)
    state.__dict__.update(candidate.__dict__)
    try:
        if _masked(template_chart(markup(chart.sheet, chart.shape.shape_id))) != _masked(xml):
            raise VBAUnsupportedError("Applying a chart template the model does not write again as it is is not implemented")
    except VBAUnsupportedError:
        state.__dict__.update(before.__dict__)
        raise
    if state.title is not None and shown is not None:
        state.title, state.title_typed = shown, True
    state.theme_override = override
    chart.changed()


def save_template(chart: Chart, filename: str) -> None:
    """SaveChartTemplate: the picture from the Export hook, which the model cannot draw, then the package over it."""
    from pyopenvba.apps.excel._hooks import call, gate, hook_of
    from pyopenvba.interpreter._files import native_path

    book = chart.sheet.book
    application = book.application
    hook = hook_of(application, 'Export')
    if hook is None:
        raise VBAUnsupportedError('Chart.SaveChartTemplate needs a function that draws the chart, which ExcelApplication.set_hook '
                                  "registers for Export; the model draws no picture of its own")
    path = Path(native_path(filename))
    gate(application, hook, 'Export', path)
    parts = template_parts(chart)
    call(application, hook, 'Export', 'Chart', chart, book, {'Filename': filename, 'FilterName': 'BMP'}, path)
    picture = path.read_bytes() if path.is_file() else b''
    written = BytesIO()
    with zipfile.ZipFile(written, 'w', zipfile.ZIP_DEFLATED) as package:
        for name in ('[Content_Types].xml', '_rels/.rels', 'chart/chart.xml', 'chart/_rels/chart.xml.rels'):
            package.writestr(name, parts[name])
        package.writestr('chart/media/image1.bmp', picture)
        for name in ('chart/charts/style1.xml', 'chart/charts/colors1.xml', 'chart/theme/themeOverride1.xml'):
            if name in parts:
                package.writestr(name, parts[name])
    path.write_bytes(written.getvalue())
