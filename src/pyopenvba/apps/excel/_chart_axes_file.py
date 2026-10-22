"""Read ordinary primary axis configuration; write only generated chart XML."""

from __future__ import annotations

import math
import re
from typing import TYPE_CHECKING

from pyopenvba._xml import escape_text
from pyopenvba.apps.excel._chart_axes import NO_AXES, SUPPORTED, AxisState, TitleState
from pyopenvba.apps.excel._chart_xml import format_number

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts import ChartState
    from pyopenvba.apps.excel._charts_file import XmlElement as Element

NS = {
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


def read(result: ChartState, root: Element, plot: Element, series_plot: Element) -> None:
    """Unknown axis groups/configuration keep the chart's read admission closed."""
    result.axes_readable = False
    # Prove the actual namespace and direct ownership independently of the
    # older chart-kind reader's local-name compatibility behavior.
    if root.tag != "{" + NS["c"] + "}chartSpace":
        return
    charts = root.findall("c:chart", NS)
    if len(charts) != 1:
        return
    owned_plots = charts[0].findall("c:plotArea", NS)
    if len(owned_plots) != 1 or owned_plots[0] is not plot:
        return
    known_plots = {"{" + NS["c"] + "}" + name for name in ("barChart", "scatterChart", "pieChart", "doughnutChart")}
    if series_plot.tag not in known_plots or not any(child is series_plot for child in plot):
        return
    if result.chart_type not in SUPPORTED:
        return
    elements = [child for child in plot if child.tag in {"{" + NS["c"] + "}catAx", "{" + NS["c"] + "}valAx"}]
    if result.chart_type in NO_AXES:
        result.axes_readable = not elements
        return

    def identifier(value: str) -> int | None:
        if re.fullmatch(r"[0-9]{1,10}", value) is None:
            return None
        number = int(value)
        return number if number <= 0xFFFFFFFF else None

    identifiers = [identifier(child.attrib.get("val", "")) for child in series_plot.findall("c:axId", NS)]
    if len(identifiers) != 2 or None in identifiers or len(set(identifiers)) != 2 or len(elements) != 2:
        return
    definitions: dict[int, Element] = {}
    for candidate in elements:
        own_ids = candidate.findall("c:axId", NS)
        if len(own_ids) != 1:
            return
        own = identifier(own_ids[0].attrib.get("val", ""))
        if own is None or own in definitions:
            return
        definitions[own] = candidate
    if set(definitions) != set(identifiers):
        return
    axes: dict[int, AxisState] = {}
    for kind, own in enumerate(identifiers, 1):
        assert own is not None
        element = definitions[own]
        cross = element.findall("c:crossAx", NS)
        expected_cross = identifiers[1 if kind == 1 else 0]
        if len(cross) != 1 or identifier(cross[0].attrib.get("val", "")) != expected_cross:
            return
        scalings = element.findall("c:scaling", NS)
        numeric = kind == 2 or result.chart_type == -4169
        wanted_tag = "{" + NS["c"] + ("}valAx" if numeric else "}catAx")
        if element.tag != wanted_tag or len(scalings) != 1:
            return
        scaling = scalings[0]
        if scaling.find("c:logBase", NS) is not None:
            return
        orientations = scaling.findall("c:orientation", NS)
        if len(orientations) > 1:
            return
        reverse = orientations[0].attrib.get("val", "minMax") if orientations else "minMax"
        if reverse not in ("minMax", "maxMin"):
            return
        deletes = element.findall("c:delete", NS)
        if len(deletes) > 1:
            return
        deleted = deletes[0] if deletes else None
        if deleted is not None and deleted.attrib.get("val") not in ("0", "1", "true", "false"):
            return
        entry = AxisState(
            kind,
            own,
            present=deleted is None or deleted.attrib.get("val") in ("0", "false"),
            numeric=numeric,
            reverse=reverse == "maxMin",
        )
        try:
            for field, tag, container in [
                ("minimum", "min", scaling),
                ("maximum", "max", scaling),
                ("major", "majorUnit", element),
            ]:
                values = container.findall("c:" + tag, NS)
                if len(values) > 1:
                    return
                if values:
                    value = float(values[0].attrib["val"])
                    if not math.isfinite(value) or (field == "major" and value <= 0):
                        return
                    setattr(entry, field, value)
        except (ValueError, KeyError, OverflowError):
            return
        if entry.minimum is not None and entry.maximum is not None and entry.minimum >= entry.maximum:
            return
        titles = element.findall("c:title", NS)
        if len(titles) > 1:
            return
        if titles:
            # Read only the owned literal text path; extension payloads are opaque.
            transactions = titles[0].findall("c:tx", NS)
            if len(transactions) > 1:
                return
            runs: list[Element] = []
            if transactions:
                rich = transactions[0].findall("c:rich", NS)
                if len(rich) != 1 or len(transactions[0]) != 1:
                    return
                paragraphs = rich[0].findall("a:p", NS)
                if len(paragraphs) != 1 or paragraphs[0].find("a:fld", NS) is not None:
                    return
                for run in paragraphs[0].findall("a:r", NS):
                    texts = run.findall("a:t", NS)
                    if len(texts) != 1:
                        return
                    runs.extend(texts)
            if sum(len(node.text or "") for node in runs) > 32768:
                return
            text = "".join(node.text or "" for node in runs)
            entry.title = TitleState(text or "Axis Title")
        axes[kind] = entry
    result.axes = axes
    result.axis_ids = (axes[1].identifier, axes[2].identifier)
    result.axes_readable = True


def generated(xml: str, result: ChartState) -> str:
    """Patch only our own bounded generated templates, never imported originals."""
    if not result.axes:
        return xml
    by_identifier = {axis.identifier: axis for axis in result.axes.values()}

    def update(match: re.Match[str]) -> str:
        raw = match.group()
        identifier = re.search(r'<c:axId val="([0-9]+)"/>', raw)
        if identifier is None:
            return raw
        axis = by_identifier[int(identifier.group(1))]
        raw = raw.replace('<c:delete val="0"/>', '<c:delete val="' + ("0" if axis.present else "1") + '"/>', 1)
        orientation = "maxMin" if axis.reverse else "minMax"
        scaling = '<c:scaling><c:orientation val="' + orientation + '"/>'
        for field, tag in [("maximum", "max"), ("minimum", "min")]:
            value = getattr(axis, field)
            if value is not None:
                scaling += "<c:" + tag + ' val="' + format_number(value) + '"/>'
        scaling += "</c:scaling>"
        raw = re.sub(r"<c:scaling>.*?</c:scaling>", lambda _match: scaling, raw, count=1)
        if axis.title is not None:
            title = (
                '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/>'
                "<a:t>"
                + escape_text(axis.title.text)
                + '</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
            )
            # Title precedes numFmt/tick metadata in CT_CatAx and CT_ValAx.
            marker = re.search(r"<c:(?:numFmt|majorTickMark)\b", raw)
            if marker is not None:
                raw = raw[: marker.start()] + title + raw[marker.start() :]
        if axis.major is not None:
            end = raw.rfind("</c:")
            raw = raw[:end] + '<c:majorUnit val="' + format_number(axis.major) + '"/>' + raw[end:]
        return raw

    return re.sub(r"<c:(?:catAx|valAx)>.*?</c:(?:catAx|valAx)>", update, xml)
