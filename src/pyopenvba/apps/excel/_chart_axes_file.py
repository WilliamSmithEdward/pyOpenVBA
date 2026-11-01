"""Read ordinary primary axis configuration; write only generated chart XML."""

from __future__ import annotations

import math
import re
from collections.abc import Callable
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._chart_axes import (
    CROSSINGS,
    DISPLAY_UNITS,
    LABEL_POSITIONS,
    NO_AXES,
    SUPPORTED,
    TICK_MARKS,
    AxisState,
    TitleState,
)
from pyopenvba.apps.excel._chart_title import read_look, title_element
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
    known_plots = {"{" + NS["c"] + "}" + name for name in (
        "barChart", "lineChart", "areaChart", "scatterChart", "pieChart", "doughnutChart")}
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
            linked = transactions and transactions[0].find("c:strRef/c:f", NS) is not None
            if transactions and not linked:
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
            entry.title = TitleState(text or "Axis Title", typed=bool(text), look=read_look(titles[0], NS))
        axes[kind] = entry
    for kind, own in enumerate(identifiers, 1):
        assert own is not None
        read_lines(axes[kind], axes[3 - kind], definitions[own])
    result.axes = axes
    result.axis_ids = (axes[1].identifier, axes[2].identifier)
    result.axes_readable = True


def read_lines(axis: AxisState, other: AxisState, element: Element) -> None:
    """Tick marks, label position and gridlines from the axis's own element, and the other axis's Crosses, which this
    element holds. Anything left out or spelled otherwise reads as unsupported rather than guessed."""

    def spelled(tag: str, names: dict[int, str]) -> str | None:
        node = element.find("c:" + tag, NS)
        value = node.attrib.get("val") if node is not None else None
        return value if value in names.values() else None

    axis.major_tick = spelled("majorTickMark", TICK_MARKS)
    axis.minor_tick = spelled("minorTickMark", TICK_MARKS)
    axis.label_position = spelled("tickLblPos", LABEL_POSITIONS)
    axis.major_gridlines = element.find("c:majorGridlines", NS) is not None
    axis.minor_gridlines = element.find("c:minorGridlines", NS) is not None

    def number(path: str) -> float | None:
        node = element.find(path, NS)
        try:
            value = float(node.attrib["val"]) if node is not None else None
        except (KeyError, ValueError, OverflowError):
            return None
        return value if value is not None and math.isfinite(value) and value > 0 else None

    axis.minor = number("c:minorUnit")
    axis.log_base = number("c:scaling/c:logBase")
    for field, path in (("label_skip", "c:tickLblSkip"), ("mark_skip", "c:tickMarkSkip")):
        skip = number(path)
        setattr(axis, field, int(skip) if skip is not None and skip.is_integer() else None)
    built_in = element.find("c:dispUnits/c:builtInUnit", NS)
    named = {name: number for number, name in DISPLAY_UNITS.items()}
    if built_in is not None and built_in.attrib.get("val") in named:
        axis.display_unit = named[built_in.attrib["val"]]
    elif number("c:dispUnits/c:custUnit") is not None:
        axis.display_unit, axis.display_custom = -4114, number("c:dispUnits/c:custUnit") or 1.0
    axis.display_label = axis.display_unit is not None and element.find("c:dispUnits/c:dispUnitsLbl", NS) is not None
    between = element.find("c:crossBetween", NS)
    if between is not None and between.attrib.get("val") in ("between", "midCat"):
        other.between = between.attrib["val"] == "between"
    other.crosses = spelled("crosses", CROSSINGS)
    custom = element.find("c:crossesAt", NS)
    if custom is not None:
        try:
            value = float(custom.attrib["val"])
        except (KeyError, ValueError, OverflowError):
            value = math.nan
        if other.crosses is None and math.isfinite(value):
            other.crosses, other.crosses_at = "custom", value
        else:
            other.crosses = None


FAR_SIDE = {"b": "t", "t": "b", "l": "r", "r": "l"}


#: BaseUnit and MajorUnitScale's xlDays, xlMonths and xlYears as the file spells them.
TIME_UNITS = ("days", "months", "years")


def _scaled_categories(raw: str, axis: AxisState, dated: bool) -> str:
    """A category axis as its CategoryType leaves it (tests/fixtures/chart_category_type.json): auto 0 once a type is
    chosen, and on a time scale, chosen or automatic over dates, a dateAx with a base time unit, and the major unit's
    scale after the major unit."""
    raw = raw.replace('<c:auto val="1"/>', f'<c:auto val="{1 if axis.category_type == -4105 else 0}"/>', 1)
    if not (axis.category_type == 3 or axis.category_type == -4105 and dated):
        return raw
    raw = raw.replace("<c:catAx>", "<c:dateAx>", 1).replace("</c:catAx>", "</c:dateAx>")
    raw = raw.replace('<c:lblAlgn val="ctr"/>', "", 1).replace('<c:noMultiLvlLbl val="0"/>', "", 1)
    unit = f'<c:baseTimeUnit val="{TIME_UNITS[axis.base_unit or 0]}"/>'
    raw = raw.replace('<c:lblOffset val="100"/>', '<c:lblOffset val="100"/>' + unit, 1)
    if axis.major_scale is not None and axis.major is not None:
        scale = f'<c:majorTimeUnit val="{TIME_UNITS[axis.major_scale]}"/>'
        raw = re.sub(r'(<c:majorUnit val="[^"]*"/>)', lambda found: found.group(1) + scale, raw, count=1)
    return raw


def generated(xml: str, result: ChartState, linked: Callable[[str], str] = lambda formula: "",
              dated: bool = False) -> str:
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
        other = next(entry for entry in result.axes.values() if entry is not axis)
        if other.reverse != (other.crosses == "max"):
            # An axis crosses the other at its start: reversing that one, or crossing at its maximum, moves this to
            # the far side, and doing both brings it back.
            raw = re.sub(r'<c:axPos val="([bltr])"/>', lambda found: '<c:axPos val="' + FAR_SIDE[found.group(1)] + '"/>',
                         raw, count=1)
        # This axis's element says where it crosses the other, which is the other's Crosses.
        crossing = (f'<c:crossesAt val="{format_number(other.crosses_at)}"/>' if other.crosses == "custom"
                    else f'<c:crosses val="{other.crosses or "autoZero"}"/>')
        raw = raw.replace('<c:crosses val="autoZero"/>', crossing, 1)
        for tag, value in (("majorTickMark", axis.major_tick), ("minorTickMark", axis.minor_tick),
                           ("tickLblPos", axis.label_position)):
            raw = re.sub(f'<c:{tag} val="[^"]*"/>', f'<c:{tag} val="{value}"/>', raw, count=1)
        gridlines = ("<c:majorGridlines/>" if axis.major_gridlines else "") + (
            "<c:minorGridlines/>" if axis.minor_gridlines else "")
        raw = re.sub(r'(<c:axPos val="[bltr]"/>)(?:<c:majorGridlines/>)?', lambda found: found.group(1) + gridlines,
                     raw, count=1)
        orientation = "maxMin" if axis.reverse else "minMax"
        scaling = "<c:scaling>"
        if axis.log_base is not None:
            scaling += '<c:logBase val="' + format_number(axis.log_base) + '"/>'
        scaling += '<c:orientation val="' + orientation + '"/>'
        for field, tag in [("maximum", "max"), ("minimum", "min")]:
            value = getattr(axis, field)
            if value is not None:
                scaling += "<c:" + tag + ' val="' + format_number(value) + '"/>'
        scaling += "</c:scaling>"
        raw = re.sub(r"<c:scaling>.*?</c:scaling>", lambda _match: scaling, raw, count=1)
        if axis.title is not None:
            look = axis.title.look
            title = title_element(look, text=axis.title.text if axis.title.typed else None,
                                  linked_text=linked(look.formula) if look.formula else "")
            # Title precedes numFmt/tick metadata in CT_CatAx and CT_ValAx.
            marker = re.search(r"<c:(?:numFmt|majorTickMark)\b", raw)
            if marker is not None:
                raw = raw[: marker.start()] + title + raw[marker.start() :]
        skips = "".join(f'<c:{tag} val="{value}"/>' for tag, value in (
            ("tickLblSkip", axis.label_skip), ("tickMarkSkip", axis.mark_skip)) if value is not None)
        raw = raw.replace("<c:noMultiLvlLbl ", skips + "<c:noMultiLvlLbl ", 1)
        if not other.numeric and other.between is not None:
            # The category axis says whether the value axis crosses between categories or on them.
            between = "between" if other.between else "midCat"
            raw = re.sub(r'<c:crossBetween val="[^"]*"/>', f'<c:crossBetween val="{between}"/>', raw, count=1)
        tail = ""
        if axis.major is not None:
            tail += '<c:majorUnit val="' + format_number(axis.major) + '"/>'
        if axis.minor is not None:
            tail += '<c:minorUnit val="' + format_number(axis.minor) + '"/>'
        if axis.display_unit is not None:
            unit = (f'<c:custUnit val="{format_number(axis.display_custom)}"/>' if axis.display_unit == -4114
                    else f'<c:builtInUnit val="{DISPLAY_UNITS[axis.display_unit]}"/>')
            tail += f'<c:dispUnits>{unit}{"<c:dispUnitsLbl/>" if axis.display_label else ""}</c:dispUnits>'
        if tail:
            end = raw.rfind("</c:")
            raw = raw[:end] + tail + raw[end:]
        if raw.startswith("<c:catAx>"):
            raw = _scaled_categories(raw, axis, dated)
        return raw

    return re.sub(r"<c:(?:catAx|valAx)>.*?</c:(?:catAx|valAx)>", update, xml)
