"""Measured primary chart axes, with explicit manual scale configuration.

Effective automatic scale values require native layout and are deliberately not
invented. Imported chart contents remain protected by Chart.admit_change.
"""

from __future__ import annotations

import math
from collections.abc import Sequence
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.apps.excel._chart_title import TitleLook, TitleMembers
from pyopenvba.apps.excel._creator import Created
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, NULL, VBAInt, error, to_bool, to_integer, to_number, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts import Chart
    from pyopenvba.apps.excel._model import Worksheet

SUPPORTED = {51, 57, 4, 65, 1, -4169, 5, -4120}
NO_AXES = {5, -4120}
MAX_TITLE_TEXT = 32768


@dataclass
class TitleState:
    text: str = "Axis Title"
    #: False for Excel's automatic title, which the file holds with no text.
    typed: bool = False
    look: TitleLook = field(default_factory=TitleLook)


@dataclass
class AxisState:
    kind: int
    identifier: int
    present: bool = True
    numeric: bool = False
    reverse: bool = False
    minimum: float | None = None
    maximum: float | None = None
    major: float | None = None
    title: TitleState | None = None
    # The rest as the file spells them; None where an imported file spells them in a way the model does not read.
    major_tick: str | None = "out"
    minor_tick: str | None = "none"
    label_position: str | None = "nextTo"
    major_gridlines: bool = False
    minor_gridlines: bool = False
    #: Where the other axis crosses this one: autoZero, max, min, or custom at ``crosses_at``.
    crosses: str | None = "autoZero"
    crosses_at: float = 0.0


TICK_MARKS = {-4142: "none", 2: "in", 3: "out", 4: "cross"}
LABEL_POSITIONS = {-4127: "high", -4134: "low", 4: "nextTo", -4142: "none"}
CROSSINGS = {-4105: "autoZero", 2: "max", 4: "min", -4114: "custom"}


def states(chart: Chart) -> dict[int, AxisState]:
    chart.admit_read()
    if not chart.state.axes_readable:
        raise VBAUnsupportedError("Reading rich imported chart axis configuration is not implemented")
    if chart.state.chart_type not in SUPPORTED:
        raise VBAUnsupportedError("Primary axes for this chart type are not implemented")
    if not chart.state.axes:
        for kind, identifier in zip((1, 2), chart.state.axis_ids, strict=True):
            chart.state.axes[kind] = AxisState(
                kind,
                identifier,
                present=chart.state.chart_type not in NO_AXES,
                numeric=kind == 2 or chart.state.chart_type == -4169,
                major_gridlines=kind == 2,
            )
    return chart.state.axes


def transition(chart: Chart, wanted: int) -> None:
    if not chart.state.axes:
        return
    if wanted not in SUPPORTED:
        raise VBAUnsupportedError("Changing a configured primary axis to this chart type is not implemented")
    old = chart.state.chart_type
    if old in NO_AXES and wanted not in NO_AXES:
        for axis in chart.state.axes.values():
            axis.minimum = axis.maximum = axis.major = None
            axis.present = True
    elif wanted in NO_AXES:
        for axis in chart.state.axes.values():
            axis.present = False
    if wanted not in NO_AXES:
        for axis in chart.state.axes.values():
            axis.numeric = axis.kind == 2 or wanted == -4169


def axis(chart: Chart, kind: object, group: object = MISSING) -> Axis:
    kind_number = int(to_integer(kind, "Long"))
    group_number = 1 if group is MISSING else int(to_integer(group, "Long"))
    chart.admit_read()
    if kind_number not in (1, 2) or group_number != 1:
        if group_number == 2:
            raise VBAUnsupportedError("Secondary chart axes are not implemented")
        raise error(-2147467259, "Method 'Axes' of object '_Chart' failed")
    return Axis(chart, states(chart)[kind_number], valid=chart.state.chart_type not in NO_AXES)


class Axes(VBACollection, Created):
    vba_type_name = "Axes"

    def __init__(self, chart: Chart) -> None:
        self.chart = chart

    def vba_items(self) -> list[object]:
        return [Axis(self.chart, entry) for entry in states(self.chart).values() if entry.present]

    @member
    def Application(self) -> object:
        return self.chart.sheet.book.application

    @member(default=True)
    def Item(self, Type: object, AxisGroup: object = MISSING) -> object:  # noqa: N803 - native named arguments
        return axis(self.chart, Type, AxisGroup)

    @member
    def Parent(self) -> object:
        self.chart.admit_read()
        return self.chart


class Axis(Created):
    vba_type_name = "Axis"

    def __init__(self, chart: Chart, entry: AxisState, *, valid: bool = True) -> None:
        self.chart, self.entry, self.valid = chart, entry, valid

    def read(self) -> None:
        self.chart.admit_read()
        if not self.valid or self.chart.state.axes.get(self.entry.kind) is not self.entry:
            raise error(440, "Automation error")

    def change(self) -> None:
        self.read()
        self.chart.admit_change()

    def scale(self) -> None:
        self.read()
        if not self.entry.numeric:
            raise error(440, "Automation error")

    @member
    def Parent(self) -> object:
        self.chart.admit_read()
        return self.chart

    @member
    def Type(self) -> object:
        self.read()
        return VBAInt(self.entry.kind, "Long")

    @member
    def AxisGroup(self) -> object:
        self.read()
        return VBAInt(1, "Long")

    @member
    def HasTitle(self) -> object:
        self.read()
        return self.entry.title is not None

    @setter("HasTitle")
    def _title(self, value: object) -> None:
        self.change()
        wanted = to_bool(value)
        self.change()
        self.chart.changed()
        if wanted:
            if self.entry.title is None:
                self.entry.title = TitleState()
        else:
            self.entry.title = None

    @member
    def AxisTitle(self) -> object:
        self.chart.admit_read()
        if not self.valid or self.entry.title is None:
            return AxisTitle(self, None)
        return AxisTitle(self, self.entry.title)

    @member
    def ReversePlotOrder(self) -> object:
        self.read()
        return self.entry.reverse

    @setter("ReversePlotOrder")
    def _reverse(self, value: object) -> None:
        self.change()
        wanted = to_bool(value)
        self.change()
        self.chart.changed()
        self.entry.reverse = wanted

    def number(self, field: str) -> float:
        self.scale()
        value = getattr(self.entry, field)
        if value is None:
            raise VBAUnsupportedError("Effective automatic chart scales require native layout")
        return value

    def set_number(self, field: str, value: object) -> None:
        self.scale()
        self.change()
        if value is NULL:
            raise error(13)
        wanted = float(to_number(value))
        self.scale()
        self.change()
        if not math.isfinite(wanted):
            raise VBAUnsupportedError("Nonfinite chart scales are not implemented")
        if field == "major" and wanted <= 0:
            raise error(440, "Automation error")
        self.chart.changed()
        setattr(self.entry, field, wanted)
        if field == "minimum" and self.entry.maximum is not None and wanted >= self.entry.maximum:
            self.entry.maximum = None
        elif field == "maximum" and self.entry.minimum is not None and wanted <= self.entry.minimum:
            self.entry.minimum = None

    def automatic(self, field: str) -> bool:
        self.scale()
        return getattr(self.entry, field) is None

    def set_automatic(self, field: str, value: object) -> None:
        self.scale()
        self.change()
        wanted = to_bool(value)
        self.scale()
        self.change()
        if not wanted and getattr(self.entry, field) is None:
            raise VBAUnsupportedError("Freezing an effective automatic chart scale requires native layout")
        self.chart.changed()
        if wanted:
            setattr(self.entry, field, None)

    @member
    def MinimumScale(self) -> object:
        return self.number("minimum")

    @setter("MinimumScale")
    def _minimum(self, value: object) -> None:
        self.set_number("minimum", value)

    @member
    def MaximumScale(self) -> object:
        return self.number("maximum")

    @setter("MaximumScale")
    def _maximum(self, value: object) -> None:
        self.set_number("maximum", value)

    @member
    def MajorUnit(self) -> object:
        return self.number("major")

    @setter("MajorUnit")
    def _major(self, value: object) -> None:
        self.set_number("major", value)

    @member
    def MinimumScaleIsAuto(self) -> object:
        return self.automatic("minimum")

    @setter("MinimumScaleIsAuto")
    def _minimum_auto(self, value: object) -> None:
        self.set_automatic("minimum", value)

    @member
    def MaximumScaleIsAuto(self) -> object:
        return self.automatic("maximum")

    @setter("MaximumScaleIsAuto")
    def _maximum_auto(self, value: object) -> None:
        self.set_automatic("maximum", value)

    @member
    def MajorUnitIsAuto(self) -> object:
        return self.automatic("major")

    @setter("MajorUnitIsAuto")
    def _major_auto(self, value: object) -> None:
        self.set_automatic("major", value)

    def spelled(self, field: str, names: dict[int, str]) -> object:
        self.read()
        value = getattr(self.entry, field)
        if value is None:
            raise VBAUnsupportedError("Reading this imported chart axis setting is not implemented")
        return VBAInt(next(number for number, name in names.items() if name == value), "Long")

    def set_spelled(self, field: str, names: dict[int, str], value: object) -> None:
        self.change()
        wanted = names.get(int(to_integer(value, "Long")))
        if wanted is None:
            raise error(-2147467259, "Automation error\nUnspecified error")
        self.chart.changed()
        setattr(self.entry, field, wanted)

    @member
    def MajorTickMark(self) -> object:
        return self.spelled("major_tick", TICK_MARKS)

    @setter("MajorTickMark")
    def _major_tick(self, value: object) -> None:
        self.set_spelled("major_tick", TICK_MARKS, value)

    @member
    def MinorTickMark(self) -> object:
        return self.spelled("minor_tick", TICK_MARKS)

    @setter("MinorTickMark")
    def _minor_tick(self, value: object) -> None:
        self.set_spelled("minor_tick", TICK_MARKS, value)

    @member
    def TickLabelPosition(self) -> object:
        return self.spelled("label_position", LABEL_POSITIONS)

    @setter("TickLabelPosition")
    def _label_position(self, value: object) -> None:
        self.set_spelled("label_position", LABEL_POSITIONS, value)

    @member
    def HasMajorGridlines(self) -> object:
        self.read()
        return self.entry.major_gridlines

    @setter("HasMajorGridlines")
    def _major_gridlines(self, value: object) -> None:
        self.change()
        wanted = to_bool(value)
        self.chart.changed()
        self.entry.major_gridlines = wanted

    @member
    def HasMinorGridlines(self) -> object:
        self.read()
        return self.entry.minor_gridlines

    @setter("HasMinorGridlines")
    def _minor_gridlines(self, value: object) -> None:
        self.change()
        wanted = to_bool(value)
        self.chart.changed()
        self.entry.minor_gridlines = wanted

    @member
    def Crosses(self) -> object:
        return self.spelled("crosses", CROSSINGS)

    @setter("Crosses")
    def _crosses(self, value: object) -> None:
        self.set_spelled("crosses", CROSSINGS, value)
        if self.entry.crosses == "autoZero":
            self.entry.crosses_at = 0.0

    @member
    def CrossesAt(self) -> object:
        self.read()
        if self.entry.crosses in ("max", "min") or self.entry.crosses is None:
            raise VBAUnsupportedError("Where an axis crosses at its drawn maximum or minimum requires native layout")
        return self.entry.crosses_at

    @setter("CrossesAt")
    def _crosses_at(self, value: object) -> None:
        self.change()
        if value is NULL:
            raise error(13)
        wanted = float(to_number(value))
        if not math.isfinite(wanted):
            raise VBAUnsupportedError("Nonfinite chart axis crossings are not implemented")
        if not self.entry.numeric and not wanted.is_integer():
            raise VBAUnsupportedError("Crossing a category axis between categories is not implemented")
        self.chart.changed()
        self.entry.crosses, self.entry.crosses_at = "custom", wanted


class AxisTitle(TitleMembers):
    vba_type_name = "AxisTitle"

    def __init__(self, axis: Axis, entry: TitleState | None) -> None:
        self.axis, self.entry = axis, entry

    def read(self) -> None:
        self.axis.chart.admit_read()
        if self.entry is None:
            raise error(424, "Object required")
        self.axis.read()
        if self.axis.entry.title is not self.entry:
            raise error(-2147467261, "This object is no longer valid.")

    def vba_call_by_name(
        self, name: str, args: Sequence[object], *, assignment: bool = False, by_ref: bool = False
    ) -> object:
        self.axis.chart.admit_read()
        if (
            self.entry is not None
            and self.axis.entry.title is not self.entry
            and not (assignment and args and args[-1] is NULL)
        ):
            raise error(440, "Automation error")
        return super().vba_call_by_name(name, args, assignment=assignment, by_ref=by_ref)

    @member
    def Parent(self) -> object:
        self.read()
        return self.axis

    def title_look(self) -> TitleLook:
        self.read()
        assert self.entry is not None
        return self.entry.look

    def title_restyle(self) -> TitleLook:
        self.read()
        self.axis.change()
        self.axis.chart.changed()
        assert self.entry is not None
        return self.entry.look

    def title_sheet(self) -> Worksheet:
        self.read()
        self.axis.change()
        return self.axis.chart.sheet

    def title_link(self, formula: str) -> None:
        self.axis.chart.changed()
        assert self.entry is not None
        self.entry.typed, self.entry.look.formula = False, formula

    def title_angle(self) -> int:
        # The title of an axis drawn upright reads upward until its orientation is set.
        upright = (self.axis.entry.kind == 1) == (self.axis.chart.state.chart_type == 57)
        return 90 if upright else 0

    @member
    def Text(self) -> object:
        from pyopenvba.apps.excel._charts import linked_title

        self.read()
        assert self.entry is not None
        if self.entry.look.formula:
            return linked_title(self.axis.chart.sheet, self.entry.look.formula)
        return self.entry.text

    @setter("Text")
    def _text(self, value: object) -> None:
        if value is NULL:
            raise error(13)
        self.read()
        self.axis.change()
        wanted = to_text(value)
        if len(wanted) > MAX_TITLE_TEXT:
            raise VBAUnsupportedError("Axis titles exceed the supported text budget")
        if any(
            not (
                ord(char) in (9, 10, 13)
                or 32 <= ord(char) <= 0xD7FF
                or 0xE000 <= ord(char) <= 0xFFFD
                or 0x10000 <= ord(char) <= 0x10FFFF
            )
            for char in wanted
        ):
            raise VBAUnsupportedError("Axis title characters requiring Office literal escapes are not implemented")
        self.read()
        self.axis.change()
        self.axis.chart.changed()
        if wanted:
            assert self.entry is not None
            self.entry.text, self.entry.typed, self.entry.look.formula = wanted, True, ""
        else:
            self.axis.entry.title = None

    @member
    def Caption(self) -> object:
        return self.Text()

    @setter("Caption")
    def _caption(self, value: object) -> None:
        self._text(value)

    @method
    def Delete(self) -> object:
        self.read()
        self.axis.chart.changed()
        self.axis.entry.title = None
        return EMPTY
