"""Native axis configuration/lifetime observations with explicit auto-layout exclusion."""

from __future__ import annotations

import hashlib
import importlib.util
import json
import os
import re
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._chart_axes import Axis, AxisTitle
from pyopenvba.apps.excel._charts import Chart, ChartObjects
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError

ROOT = Path(__file__).parent / "fixtures/chart_axes"
MANIFEST = json.loads((ROOT / "manifest.json").read_bytes())
INITIAL = json.loads((ROOT / "initial.json").read_bytes())
FOLLOWUP = json.loads((ROOT / "followup.json").read_bytes())
CONTROLS = [
    p
    for raw, key in [(INITIAL, "admitted_initial"), (FOLLOWUP, "admitted_followup")]
    for p in raw["probes"]
    if p["name"] in MANIFEST[key]
]


AXIS_OBSERVATIONS = """Function AxisSnap(a As Object) As String
Dim k As Variant,s As String,automatic As String
s="parent=" & ParentType(a)
For Each k In Array("Type","AxisGroup","HasTitle")
s=s & ";" & k & "=" & ReadP(a,CStr(k))
Next
For Each k In Array("MinimumScale","MaximumScale","MajorUnit")
automatic=ReadP(a,CStr(k) & "IsAuto")
s=s & ";" & k & "IsAuto=" & automatic
If automatic <> "True(11)" Then s=s & ";" & k & "=" & ReadP(a,CStr(k))
Next
s=s & ";ReversePlotOrder=" & ReadP(a,"ReversePlotOrder")
AxisSnap=s & ";title=[" & TitleRead(a) & "]"
End Function"""


def source(control: dict[str, Any]) -> str:
    text = control["source"].replace(' & ";parent=" & ParentType(s)', "")
    return re.sub(r"Function AxisSnap\(.*?End Function", AXIS_OBSERVATIONS, text, flags=re.DOTALL)


def project(text: str) -> str:
    # Only a numeric observation whose immediately preceding native auto flag
    # is True is excluded. False/manual values and category/error values stay.
    text = re.sub(
        r";(?P<field>MinimumScale|MaximumScale|MajorUnit)IsAuto=True\(11\);(?P=field)=[^;]*",
        lambda match: ";" + match["field"] + "IsAuto=True(11)",
        text,
    )
    return re.sub(r";parent=(?:ChartGroup|ERR:1004:Application-defined or object-defined error)", "", text)


@pytest.mark.parametrize("control", CONTROLS, ids=lambda c: c["name"])
def test_native_axis_configuration_and_lifetimes(control: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source(control), name="AxesControl")
    assert app.run("Probe") == project(control["raw_result"])


def test_raw_evidence_exact() -> None:
    for name, sha in {**MANIFEST["raw_sha256"], **MANIFEST["native_inputs"]}.items():
        assert hashlib.sha256((ROOT / name).read_bytes()).hexdigest() == sha
    assert len(CONTROLS) == 8


def new_chart() -> Chart:
    app = ExcelApplication()
    app.add_workbook()
    return ChartObjects(app.workbook.sheets_[0]).Add(0, 0, 300, 200).Chart()


def test_manual_values_auto_flags_and_exact_numeric_persistence(tmp_path: Path) -> None:
    chart = new_chart()
    axis = chart.Axes(2)
    assert isinstance(axis, Axis)
    axis.vba_set("MinimumScale", 0)
    axis.vba_set("MaximumScale", 40)
    axis.vba_set("MajorUnit", 10)
    assert [axis.MinimumScale(), axis.MaximumScale(), axis.MajorUnit()] == [0, 40, 10]
    assert [axis.MinimumScaleIsAuto(), axis.MaximumScaleIsAuto(), axis.MajorUnitIsAuto()] == [False] * 3
    axis.vba_set("ReversePlotOrder", True)
    axis.vba_set("HasTitle", True)
    title = axis.AxisTitle()
    assert isinstance(title, AxisTitle)
    title.vba_set("Text", 'Value & "margin" <2026>')
    path = tmp_path / "manual.xlsx"
    chart.sheet.book.save_to(path, file_format=51)
    reopened = ExcelApplication.open(path, with_vba=False)
    loaded = reopened.workbook.sheets_[0].ChartObjects().vba_get("Item", [1]).Chart().Axes(2)
    assert [loaded.MinimumScale(), loaded.MaximumScale(), loaded.MajorUnit()] == [0, 40, 10]
    assert loaded.ReversePlotOrder() is True
    assert loaded.AxisTitle().Text() == 'Value & "margin" <2026>'
    with pytest.raises(VBAUnsupportedError, match="imported"):
        loaded.vba_set("MinimumScale", -1)


@pytest.mark.parametrize("name", ["MinimumScale", "MaximumScale", "MajorUnit"])
def test_effective_auto_values_and_freezing_do_not_fabricate_layout(name: str) -> None:
    chart = new_chart()
    axis = chart.Axes(2)
    before = (chart.state.dirty, chart.sheet.drawing_dirty, chart.sheet.book.saved)
    with pytest.raises(VBAUnsupportedError, match="native layout"):
        axis.vba_get(name)
    with pytest.raises(VBAUnsupportedError, match="native layout"):
        axis.vba_set(name + "IsAuto", False)
    assert (chart.state.dirty, chart.sheet.drawing_dirty, chart.sheet.book.saved) == before


def test_invalid_manual_values_are_atomic() -> None:
    chart = new_chart()
    axis = chart.Axes(2)
    axis.vba_set("MajorUnit", 10)
    chart.state.dirty = False
    chart.sheet.drawing_dirty = False
    chart.sheet.book.saved = True
    with pytest.raises(VBARuntimeError) as failure:
        axis.vba_set("MajorUnit", 0)
    assert failure.value.number == 440 and axis.MajorUnit() == 10
    assert not chart.state.dirty and not chart.sheet.drawing_dirty and chart.sheet.book.saved
    with pytest.raises(VBARuntimeError) as failure:
        chart.Axes(1).MinimumScale()
    assert failure.value.number == 440


def test_native_absent_proxy_and_deleted_direct_vs_late_bound_errors() -> None:
    raw = json.loads((ROOT / "absence.json").read_bytes())
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(raw["source"], name="AxisTitleAbsence")
    assert app.run("Probe") == raw["raw_result"]


@pytest.mark.parametrize("name", ["manual-scale", "axis-titles", "bar-scatter"])
def test_native_input_and_generated_saved_axis_snapshot(name: str, tmp_path: Path) -> None:
    control = next(p for p in INITIAL["probes"] if p["name"] == name)
    expected = project(control["reopened"])
    native = ExcelApplication.open(ROOT / control["file"], with_vba=False)
    native.add_module(source(control), name="AxisSnapshot")
    assert native.run("Snapshot") == expected
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source(control), name="AxisSnapshot")
    app.run("Probe")
    saved = app.save(tmp_path / control["file"])
    reopened = ExcelApplication.open(saved, with_vba=False)
    reopened.add_module(source(control), name="AxisSnapshot")
    assert reopened.run("Snapshot") == expected


def test_native_conflicting_bounds_reset_other_auto_and_invalid_values_preserve() -> None:
    from pyopenvba.interpreter._values import NULL

    chart = new_chart()
    axis = chart.Axes(2)
    axis.vba_set("MinimumScale", 10)
    axis.vba_set("MaximumScale", 5)
    assert axis.MinimumScaleIsAuto() is True and axis.MaximumScale() == 5
    for name, value in [("MinimumScale", NULL), ("MajorUnit", NULL), ("MinimumScale", "bad")]:
        before = (axis.entry.minimum, axis.entry.maximum, axis.entry.major)
        with pytest.raises(VBARuntimeError) as failure:
            axis.vba_set(name, value)
        assert failure.value.number == 13
        assert before == (axis.entry.minimum, axis.entry.maximum, axis.entry.major)


def test_title_coercion_cannot_resurrect_deleted_proxy() -> None:
    from pyopenvba.interpreter._objects import VBAObject

    chart = new_chart()
    axis = chart.Axes(2)
    axis.vba_set("HasTitle", True)
    title = axis.AxisTitle()

    class DeletesTitle(VBAObject):
        def vba_value(self) -> object:
            title.Delete()
            return "resurrected"

    with pytest.raises(VBARuntimeError) as failure:
        title.vba_set("Text", DeletesTitle())
    assert failure.value.number == -2147467261 and axis.HasTitle() is False


def test_unknown_imported_axis_metadata_keeps_read_and_write_admission_closed(tmp_path: Path) -> None:
    from zipfile import ZipFile

    target = tmp_path / "rich-axis.xlsx"
    with ZipFile(ROOT / "manual-scale.xlsx") as original, ZipFile(target, "w") as result:
        for entry in original.infolist():
            raw = original.read(entry.filename)
            if entry.filename == "xl/charts/chart1.xml":
                raw = raw.replace(b"<c:scaling>", b'<c:scaling><c:orientation val="minMax"/>', 1)
            result.writestr(entry, raw)
    before = target.read_bytes()
    app = ExcelApplication.open(target, with_vba=False)
    chart = app.workbook.sheets_[0].ChartObjects().vba_get("Item", [1]).Chart()
    with pytest.raises(VBAUnsupportedError, match="rich imported"):
        chart.Axes(2)
    assert target.read_bytes() == before


def test_imported_logarithmic_axis_reads_its_base(tmp_path: Path) -> None:
    from zipfile import ZipFile

    target = tmp_path / "log-axis.xlsx"
    with ZipFile(ROOT / "manual-scale.xlsx") as original, ZipFile(target, "w") as result:
        for entry in original.infolist():
            raw = original.read(entry.filename)
            if entry.filename == "xl/charts/chart1.xml":
                head, tail = raw.split(b"<c:valAx>", 1)
                raw = head + b"<c:valAx>" + tail.replace(b"<c:scaling>", b'<c:scaling><c:logBase val="2"/>', 1)
            result.writestr(entry, raw)
    app = ExcelApplication.open(target, with_vba=False)
    axis = app.workbook.sheets_[0].ChartObjects().vba_get("Item", [1]).Chart().Axes(2)
    assert (axis.ScaleType(), axis.LogBase()) == (-4133, 2.0)


def test_detached_axis_and_title_cannot_mutate_chart_storage() -> None:
    chart = new_chart()
    axis = chart.Axes(2)
    axis.vba_set("HasTitle", True)
    title = axis.AxisTitle()
    before = (axis.entry.minimum, axis.entry.title)
    chart.Parent().Delete()
    for action in [
        lambda: axis.MinimumScaleIsAuto(),
        lambda: title.Text(),
        lambda: title.Delete(),
        lambda: axis.vba_set("MinimumScale", -1),
    ]:
        with pytest.raises(VBAUnsupportedError, match="detached"):
            action()
    assert before == (axis.entry.minimum, axis.entry.title)


@pytest.fixture(scope="module")
def native_axis_excel():
    harness = pytest.importorskip("pyvbaharness")
    with harness.ExcelSession() as excel:
        yield excel


@pytest.mark.skipif(
    os.environ.get("RUN_LIVE_EXCEL") != "1" or importlib.util.find_spec("pyvbaharness") is None,
    reason="opt-in generated primary-axis native gate",
)
@pytest.mark.parametrize("name", ["manual-scale", "axis-titles", "bar-scatter"])
def test_generated_axes_reopened_in_native(name: str, tmp_path: Path, native_axis_excel: Any) -> None:
    control = next(p for p in INITIAL["probes"] if p["name"] == name)
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source(control), name="AxisGenerated")
    assert app.run("Probe") == project(control["raw_result"])
    path = app.save(tmp_path / control["file"])
    native_axis_excel.open_document(path)
    result = native_axis_excel.run_vba(
        source(control), "Snapshot", module_name="AxisSnapshot", line_numbers=False, timeout=None, idle_timeout=None
    )
    assert result.ok and not result.dialogs, result
    assert result.value == project(control["reopened"])


@pytest.mark.parametrize("text", ["bad\x00text", "bad\x01text", "bad\ud800text"])
def test_unserializable_title_text_refuses_before_mutation(text: str) -> None:
    chart = new_chart()
    axis = chart.Axes(2)
    axis.vba_set("HasTitle", True)
    title = axis.AxisTitle()
    chart.state.dirty = False
    chart.sheet.drawing_dirty = False
    chart.sheet.book.saved = True
    with pytest.raises(VBAUnsupportedError, match="literal escapes"):
        title.vba_set("Text", text)
    assert title.Text() == "Axis Title" and axis.HasTitle() is True
    assert not chart.state.dirty and not chart.sheet.drawing_dirty and chart.sheet.book.saved


def test_native_hidden_axis_retains_manual_configuration_and_title_after_reopen(tmp_path: Path) -> None:
    raw = json.loads((ROOT / "hidden-save.json").read_bytes())
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source(raw), name="HiddenAxis")
    assert app.run("Probe") == project(raw["raw_result"])
    for saved in [ROOT / "hidden-save.xlsx", app.save(tmp_path / "hidden.xlsx")]:
        loaded = ExcelApplication.open(saved, with_vba=False)
        loaded.add_module(source(raw), name="HiddenAxisRead")
        assert loaded.run("Snapshot") == project(raw["reopened_result"])
        chart = loaded.workbook.sheets_[0].ChartObjects().vba_get("Item", [1]).Chart()
        assert chart.Axes().Count() == 1 and chart.HasAxis(2, 1) is False
        axis = chart.Axes(2)
        assert [axis.MinimumScale(), axis.MaximumScale(), axis.MajorUnit()] == [-10, 40, 10]
        assert axis.AxisTitle().Text() == "Hidden axis" and axis.ReversePlotOrder() is True


@pytest.mark.skipif(
    os.environ.get("RUN_LIVE_EXCEL") != "1" or importlib.util.find_spec("pyvbaharness") is None,
    reason="opt-in generated hidden-axis native gate",
)
def test_generated_hidden_axis_reopened_in_native(tmp_path: Path, native_axis_excel: Any) -> None:
    raw = json.loads((ROOT / "hidden-save.json").read_bytes())
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source(raw), name="HiddenAxis")
    assert app.run("Probe") == project(raw["raw_result"])
    output = app.save(tmp_path / "hidden.xlsx")
    native_axis_excel.open_document(output)
    result = native_axis_excel.run_vba(
        source(raw), "Snapshot", module_name="HiddenAxisRead", line_numbers=False, timeout=None, idle_timeout=None
    )
    assert result.ok and not result.dialogs, result
    assert result.value == project(raw["reopened_result"])


@pytest.mark.parametrize(
    "case",
    [
        "superscript",
        "arabic-digit",
        "duplicate-axId",
        "duplicate-scaling",
        "duplicate-delete",
        "duplicate-orientation",
        "duplicate-crossAx",
        "broken-crossAx",
        "numeric-alias-id",
    ],
)
def test_imported_axis_identity_and_known_fields_are_unambiguous_before_read(case: str, tmp_path: Path) -> None:
    from zipfile import ZipFile

    target = tmp_path / (case + ".xlsx")
    with ZipFile(ROOT / "manual-scale.xlsx") as native, ZipFile(target, "w") as result:
        for entry in native.infolist():
            raw = native.read(entry.filename)
            if entry.filename == "xl/charts/chart1.xml":
                text = raw.decode("utf8")
                ids = re.findall(r'<c:axId val="([0-9]+)"/>', text)
                first, second = ids[:2]
                if case in ["superscript", "arabic-digit"]:
                    text = text.replace(
                        'val="' + first + '"', 'val="' + ("²" if case == "superscript" else "\u0661") + '"'
                    )
                elif case == "numeric-alias-id":
                    text = text.replace('val="' + second + '"', 'val="0' + first + '"')
                elif case == "broken-crossAx":
                    text = text.replace('<c:crossAx val="' + first + '"/>', '<c:crossAx val="' + second + '"/>', 1)
                else:
                    tag = case.removeprefix("duplicate-")
                    pattern = r"<c:scaling>.*?</c:scaling>" if tag == "scaling" else r"<c:" + tag + r" [^>]*/>"
                    # Alter an actual axis definition, not its plot references.
                    axis_start = text.index("<c:catAx>")
                    before, axis = text[:axis_start], text[axis_start:]
                    match = re.search(pattern, axis)
                    assert match is not None
                    text = before + axis[: match.end()] + match.group() + axis[match.end() :]
                raw = text.encode("utf8")
            result.writestr(entry, raw)
    original = target.read_bytes()
    app = ExcelApplication.open(target, with_vba=False)
    chart = app.workbook.sheets_[0].ChartObjects().vba_get("Item", [1]).Chart()
    with pytest.raises(VBAUnsupportedError, match="rich imported"):
        chart.Axes(2)
    assert target.read_bytes() == original and not chart.state.dirty


def test_imported_axis_opaque_metadata_is_preserved_without_refusing_known_fields(tmp_path: Path) -> None:
    from zipfile import ZipFile

    target = tmp_path / "opaque-axis.xlsx"
    original_axis = b""
    with ZipFile(ROOT / "manual-scale.xlsx") as native, ZipFile(target, "w") as result:
        for entry in native.infolist():
            raw = native.read(entry.filename)
            if entry.filename == "xl/charts/chart1.xml":
                raw = raw.replace(b"<c:valAx>", b'<c:valAx xmlns:opaque="urn:axis-preserve" opaque:keep="a &gt; b">', 1)
                raw = raw.replace(
                    b"</c:valAx>",
                    b'<c:extLst><c:ext uri="urn:axis-preserve"><opaque:settings xmlns:opaque="urn:axis-preserve">keep &amp; unchanged</opaque:settings></c:ext></c:extLst></c:valAx>',
                    1,
                )
                original_axis = raw
            result.writestr(entry, raw)
    app = ExcelApplication.open(target, with_vba=False)
    chart = app.workbook.sheets_[0].ChartObjects().vba_get("Item", [1]).Chart()
    assert [chart.Axes(2).MinimumScale(), chart.Axes(2).MaximumScale(), chart.Axes(2).MajorUnit()] == [0, 40, 10]
    output = app.save(tmp_path / "opaque-saved.xlsx")
    with ZipFile(output) as result:
        assert result.read("xl/charts/chart1.xml") == original_axis


def test_title_extension_text_does_not_become_visible_axis_title(tmp_path: Path) -> None:
    from zipfile import ZipFile

    target = tmp_path / "opaque-title.xlsx"
    original_xml = b""
    with ZipFile(ROOT / "axis-titles.xlsx") as native, ZipFile(target, "w") as result:
        for entry in native.infolist():
            raw = native.read(entry.filename)
            if entry.filename == "xl/charts/chart1.xml":
                # Alter only the axis title, leaving the chart's independent title alone.
                start = raw.index(b"<c:valAx>")
                before, axis = raw[:start], raw[start:]
                axis = axis.replace(
                    b"</c:title>",
                    b'<c:extLst><c:ext uri="urn:opaque-title"><a:t>opaque only</a:t></c:ext></c:extLst></c:title>',
                    1,
                )
                raw = before + axis
                original_xml = raw
            result.writestr(entry, raw)
    app = ExcelApplication.open(target, with_vba=False)
    chart = app.workbook.sheets_[0].ChartObjects().vba_get("Item", [1]).Chart()
    assert chart.Axes(2).AxisTitle().Text() == "Final value"
    with ZipFile(app.save(tmp_path / "title-preserved.xlsx")) as result:
        assert result.read("xl/charts/chart1.xml") == original_xml


def test_native_projection_retains_manual_numbers_and_category_errors() -> None:
    raw = next(p["raw_result"] for p in INITIAL["probes"] if p["name"] == "manual-scale")
    projected = project(raw)
    manual = re.findall(r";(MinimumScale|MaximumScale|MajorUnit)IsAuto=False\(11\);\1=([^;]*)", raw)
    assert manual
    for field, value in manual:
        pair = ";" + field + "IsAuto=False(11);" + field + "=" + value
        assert projected.count(pair) == raw.count(pair)
    errors = re.findall(
        r";(MinimumScale|MaximumScale|MajorUnit)IsAuto=ERR:440:Automation error;\1=ERR:440:Automation error", raw
    )
    assert errors
    for field in errors:
        pair = ";" + field + "IsAuto=ERR:440:Automation error;" + field + "=ERR:440:Automation error"
        assert projected.count(pair) == raw.count(pair)
    assert re.search(r";(MinimumScale|MaximumScale|MajorUnit)IsAuto=True\(11\);\1=", projected) is None
    # These independently observed conflict adjustments remain in the oracle.
    assert ";MinimumScaleIsAuto=False(11);MinimumScale=80(5)" in projected
    assert ";MaximumScaleIsAuto=False(11);MaximumScale=40(5)" in projected
    assert ";MajorUnitIsAuto=False(11);MajorUnit=10(5)" in projected


@pytest.mark.parametrize("case", ["foreign-series-plot", "foreign-chart-space", "duplicate-chart", "duplicate-plot"])
def test_axis_read_requires_real_chart_namespace_and_single_direct_ownership(case: str, tmp_path: Path) -> None:
    from zipfile import ZipFile

    target = tmp_path / (case + ".xlsx")
    chart_bytes = b""
    with ZipFile(ROOT / "manual-scale.xlsx") as native, ZipFile(target, "w") as result:
        for entry in native.infolist():
            raw = native.read(entry.filename)
            if entry.filename == "xl/charts/chart1.xml":
                text = raw.decode("utf8")
                if case == "foreign-series-plot":
                    text = text.replace("<c:barChart>", '<foreign:barChart xmlns:foreign="urn:not-chart">', 1).replace(
                        "</c:barChart>", "</foreign:barChart>", 1
                    )
                elif case == "foreign-chart-space":
                    text = text.replace(
                        "<c:chartSpace ", '<foreign:chartSpace xmlns:foreign="urn:not-chart" ', 1
                    ).replace("</c:chartSpace>", "</foreign:chartSpace>", 1)
                elif case == "duplicate-chart":
                    chart = re.search(r"<c:chart>.*?</c:chart>", text)
                    assert chart is not None
                    text = text[: chart.end()] + chart.group() + text[chart.end() :]
                else:
                    plot = re.search(r"<c:plotArea>.*?</c:plotArea>", text)
                    assert plot is not None
                    text = text[: plot.end()] + plot.group() + text[plot.end() :]
                raw = text.encode("utf8")
                chart_bytes = raw
            result.writestr(entry, raw)
    original = target.read_bytes()
    app = ExcelApplication.open(target, with_vba=False)
    chart = app.workbook.sheets_[0].ChartObjects().vba_get("Item", [1]).Chart()
    with pytest.raises(VBAUnsupportedError, match="rich imported"):
        chart.Axes(2)
    assert target.read_bytes() == original
    with ZipFile(app.save(tmp_path / "unchanged.xlsx")) as saved:
        assert saved.read("xl/charts/chart1.xml") == chart_bytes


@pytest.mark.parametrize(
    "operation", ["scale", "auto", "reverse", "has-title", "title-text", "title-delete", "presence", "type"]
)
def test_axis_mutations_preserve_state_when_drawing_protection_refuses(operation: str) -> None:
    chart = new_chart()
    axis = chart.Axes(2)
    axis.vba_set("HasTitle", True)
    title = axis.AxisTitle()
    chart.sheet.Protect(DrawingObjects=True)
    before = (
        axis.entry.minimum,
        axis.entry.maximum,
        axis.entry.major,
        axis.entry.reverse,
        axis.entry.present,
        axis.entry.title,
        id(axis.entry.title),
        title.Text(),
        chart.state.chart_type,
        chart.state.dirty,
        chart.sheet.drawing_dirty,
        chart.sheet.book.saved,
    )
    actions = {
        "scale": lambda: axis.vba_set("MinimumScale", -10),
        "auto": lambda: axis.vba_set("MinimumScaleIsAuto", True),
        "reverse": lambda: axis.vba_set("ReversePlotOrder", True),
        "has-title": lambda: axis.vba_set("HasTitle", False),
        "title-text": lambda: title.vba_set("Text", "blocked"),
        "title-delete": lambda: title.Delete(),
        "presence": lambda: chart.vba_set("HasAxis", False, [2, 1]),
        "type": lambda: chart.vba_set("ChartType", 5),
    }
    with pytest.raises(VBAUnsupportedError, match="drawing objects are protected"):
        actions[operation]()
    assert before == (
        axis.entry.minimum,
        axis.entry.maximum,
        axis.entry.major,
        axis.entry.reverse,
        axis.entry.present,
        axis.entry.title,
        id(axis.entry.title),
        title.Text(),
        chart.state.chart_type,
        chart.state.dirty,
        chart.sheet.drawing_dirty,
        chart.sheet.book.saved,
    )
