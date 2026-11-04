"""Shape.Shadow, Glow, SoftEdge and Reflection, as live Excel answers them (tests/fixtures/shape_effects.json).

What is set is kept on the shape's look, in points, EMU and fractions as the file spells them. A shadow's offsets are
kept as the file's distance and direction, so that what reads back has been through them, as in Excel. A shadow that is
not visible but has been changed goes in the file's hidden effects."""
from __future__ import annotations

import math
from collections.abc import Sequence
from typing import TYPE_CHECKING, cast

from pyopenvba.apps.excel._creator import Created
from pyopenvba.apps.excel._held import HeldPart
from pyopenvba.apps.excel._shape_format import STYLE, DEFAULT_THEME
from pyopenvba.apps.excel._styles import bgr
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import method, member, setter
from pyopenvba.interpreter._values import EMPTY, VBAInt, VBASingle, error, to_integer, to_number
from pyopenvba.shapes._values import FillLine, ShapeEffects

if TYPE_CHECKING:
    from pyopenvba.apps.excel._shapes import ShapeObject, ShapeRange

_MIXED = -2147483648
_EMU_PER_POINT = 12700
_ANGLE = 60000
#: What a shadow reads for its offsets until they are set.
_FIRST_OFFSET = 2.079952
_OUT_OF_RANGE = (-2147024809, 'The specified value is out of range.')
_SOFT_EDGES = {1: 12700, 2: 31750, 3: 63500, 4: 127000, 5: 317500, 6: 635000}
#: Reflection.Type: start alpha, end alpha, end position and distance. The rest is the same in each.
_REFLECTIONS = {1: (52000, 300, 35000, 0), 2: (50000, 300, 55000, 0), 3: (50000, 300, 90000, 0),
                4: (50000, 300, 38500, 50800), 5: (50000, 300, 55500, 50800), 6: (50000, 300, 90000, 50800),
                7: (50000, 275, 40000, 101600), 8: (50000, 300, 55500, 101600), 9: (50000, 295, 92000, 101600)}
_ACCENTS = {5: 'accent1', 6: 'accent2', 7: 'accent3', 8: 'accent4', 9: 'accent5', 10: 'accent6'}


def _fraction(value: object) -> float:
    wanted = float(VBASingle(float(to_number(value))))
    if not 0 <= wanted <= 1:
        raise error(*_OUT_OF_RANGE)
    return wanted


class _Effect(HeldPart):
    """What the four share: the shape, and the effects a change writes."""

    vba_type_name = ''

    def __init__(self, owner: ShapeObject) -> None:
        self.owner = owner

    def admit(self) -> None:
        shape = self.owner.shape
        if shape.source or shape.kind not in STYLE:
            raise VBAUnsupportedError('The effects of this shape are not implemented')
        if self.owner.sheet.book.stylesheet.colors.theme != DEFAULT_THEME:
            raise VBAUnsupportedError('The effects of a shape under another theme are not implemented')

    def held(self) -> ShapeEffects:
        """What is set, for a read: nothing yet if nothing was."""
        self.admit()
        look = self.owner.shape.fill_line
        return ShapeEffects() if look is None or look.extra is None else look.extra

    def change(self) -> ShapeEffects:
        self.admit()
        self.owner.guard('Effects')
        shape = self.owner.shape
        if shape.fill_line is None:
            shape.fill_line = FillLine()
        if shape.fill_line.extra is None:
            shape.fill_line.extra = ShapeEffects()
        self.owner.restyled()
        if not isinstance(self, ShadowFormat):
            # Another effect brings out a shadow changed while it was hidden.
            shape.fill_line.extra.touched = True
            shape.fill_line.extra.shadow_visible = shape.fill_line.extra.shadow_on
        return shape.fill_line.extra

    @member
    def Application(self) -> object:
        return self.owner.sheet.book.application

    def theme_colour(self, scheme: str) -> int:
        colours = self.owner.sheet.book.stylesheet.colors.theme
        return bgr(colours[3 + int(scheme[-1])])


class EffectColor(HeldPart):
    """A shadow's or glow's colour."""

    vba_type_name = 'ColorFormat'

    def __init__(self, owner: ShadowFormat | GlowFormat) -> None:
        self.owner = owner

    @member
    def RGB(self) -> object:
        return self.owner.read_colour()

    @setter('RGB')
    def _rgb(self, value: object) -> None:
        colour = int(to_integer(value, 'Long')) & 0xFFFFFF
        red, green, blue = colour & 255, (colour >> 8) & 255, (colour >> 16) & 255
        self.owner.set_colour(f'{red:02X}{green:02X}{blue:02X}')

    @member
    def ObjectThemeColor(self) -> object:
        # Mixed with the RGB while there is no colour yet, 0 for an RGB; a glow's theme colour is not read back here.
        if self.owner.read_colour() == -2147483648:
            return VBAInt(-2, 'Long')
        if isinstance(self.owner, GlowFormat) and self.owner.held().glow_scheme:
            raise VBAUnsupportedError('The theme colour of a glow is not implemented')
        return VBAInt(0, 'Long')

    @setter('ObjectThemeColor')
    def _object_theme_color(self, value: object) -> None:
        wanted = int(to_integer(value, 'Long'))
        if wanted not in _ACCENTS or not isinstance(self.owner, GlowFormat):
            raise VBAUnsupportedError('This theme colour of an effect is not implemented')
        self.owner.set_scheme(_ACCENTS[wanted])


class ShadowFormat(_Effect):
    vba_type_name = 'ShadowFormat'

    @member
    def Parent(self) -> object:
        return self.owner

    @member
    def Visible(self) -> object:
        return VBAInt(-1 if self.held().shadow_visible else 0, 'Long')

    @setter('Visible')
    def _visible(self, value: object) -> None:
        wanted = int(to_integer(value, 'Long')) != 0
        if not wanted and not self.held().shadow_on:
            return
        effects = self.change()
        effects.shadow_visible = wanted
        effects.shadow_on = True

    def worked(self) -> ShapeEffects:
        """The effects, with the shadow marked as having been changed."""
        effects = self.change()
        effects.shadow_on = True
        return effects

    @member
    def ForeColor(self) -> object:
        return EffectColor(self)

    def read_colour(self) -> object:
        effects = self.held()
        if not effects.shadow_on:
            return VBAInt(_MIXED, 'Long')
        return VBAInt(bgr(effects.shadow_rgb) if effects.shadow_rgb else 0, 'Long')

    def set_colour(self, colour: str) -> None:
        effects = self.change()
        effects.list_only = effects.list_only or not effects.shadow_on
        if effects.shadow_on:
            effects.shadow_rgb = colour

    def _offsets(self) -> tuple[float, float]:
        effects = self.held()
        if not effects.shadow_moved:
            return _FIRST_OFFSET, _FIRST_OFFSET
        distance = effects.shadow_distance / _EMU_PER_POINT
        angle = math.radians(effects.shadow_direction / _ANGLE)
        return distance * math.cos(angle), distance * math.sin(angle)

    def _moved(self, across: float, down: float) -> None:
        effects = self.worked()
        effects.shadow_moved = True
        effects.shadow_distance = int(math.hypot(across, down) * _EMU_PER_POINT)
        turned = math.degrees(math.atan2(down, across)) % 360
        effects.shadow_direction = round(turned * _ANGLE)

    @member
    def OffsetX(self) -> object:
        return VBASingle(self._offsets()[0])

    @setter('OffsetX')
    def _offset_x(self, value: object) -> None:
        self._moved(float(to_number(value)), self._offsets()[1])

    @member
    def OffsetY(self) -> object:
        return VBASingle(self._offsets()[1])

    @setter('OffsetY')
    def _offset_y(self, value: object) -> None:
        self._moved(self._offsets()[0], float(to_number(value)))

    @method
    def IncrementOffsetX(self, Increment: object) -> object:  # noqa: N803 - native named argument
        across, down = self._offsets()
        self._moved(across + float(to_number(Increment)), down)
        return EMPTY

    @method
    def IncrementOffsetY(self, Increment: object) -> object:  # noqa: N803 - native named argument
        across, down = self._offsets()
        self._moved(across, down + float(to_number(Increment)))
        return EMPTY

    @member
    def Blur(self) -> object:
        return VBASingle(self.held().shadow_blur / _EMU_PER_POINT)

    @setter('Blur')
    def _blur(self, value: object) -> None:
        wanted = float(to_number(value))
        if wanted < 0:
            raise error(*_OUT_OF_RANGE)
        self.worked().shadow_blur = round(wanted * _EMU_PER_POINT)

    @member
    def Transparency(self) -> object:
        effects = self.held()
        if not effects.shadow_on:
            return VBASingle(float(_MIXED))
        return VBASingle(0.0 if effects.shadow_alpha is None else (100000 - effects.shadow_alpha) / 100000)

    @setter('Transparency')
    def _transparency(self, value: object) -> None:
        wanted = _fraction(value)
        self.worked().shadow_alpha = round((1 - wanted) * 100000)

    @member
    def Size(self) -> object:
        size = self.held().shadow_size
        return VBASingle(100.0 if size is None else size / 1000)

    @setter('Size')
    def _size(self, value: object) -> None:
        wanted = float(to_number(value))
        if not 0 <= wanted <= 200:
            raise error(*_OUT_OF_RANGE)
        self.worked().shadow_size = round(wanted * 1000)

    @member
    def Style(self) -> object:
        return VBAInt(2 if self.held().shadow_on else -2, 'Long')

    @setter('Style')
    def _style(self, value: object) -> None:
        if int(to_integer(value, 'Long')) != 2:
            raise VBAUnsupportedError('A shadow but an outer one is not implemented')
        self.worked()

    @member
    def Type(self) -> object:
        self.admit()
        return VBAInt(-2, 'Long')

    @setter('Type')
    def _type(self, value: object) -> None:
        raise VBAUnsupportedError('A preset shadow is not implemented')

    @member
    def Obscured(self) -> object:
        return VBAInt(-1 if self.held().shadow_obscured else 0, 'Long')

    @setter('Obscured')
    def _obscured(self, value: object) -> None:
        self.change().shadow_obscured = int(to_integer(value, 'Long')) != 0

    @member
    def RotateWithShape(self) -> object:
        self.admit()
        return VBAInt(0, 'Long')

    @setter('RotateWithShape')
    def _rotate_with_shape(self, value: object) -> None:
        if int(to_integer(value, 'Long')) != 0:
            raise VBAUnsupportedError('A shadow that rotates with its shape is not implemented')
        self.worked()


class GlowFormat(_Effect):
    vba_type_name = 'GlowFormat'

    @member
    def Parent(self) -> object:
        raise error(438, "Object doesn't support this property or method")

    @member
    def Radius(self) -> object:
        radius = self.held().glow_radius
        return VBASingle(0.0 if radius is None else radius / _EMU_PER_POINT)

    @setter('Radius')
    def _radius(self, value: object) -> None:
        wanted = float(to_number(value))
        if not 0 <= wanted <= 150:
            raise error(*_OUT_OF_RANGE)
        effects = self.change()
        effects.glow = wanted != 0
        effects.glow_radius = round(wanted * _EMU_PER_POINT) if wanted else None
        if not wanted:
            effects.glow_rgb, effects.glow_scheme, effects.glow_alpha = None, None, None

    @member
    def Color(self) -> object:
        return EffectColor(self)

    def read_colour(self) -> object:
        effects = self.held()
        if not effects.glow:
            return VBAInt(_MIXED, 'Long')
        if effects.glow_scheme:
            return VBAInt(self.theme_colour(effects.glow_scheme), 'Long')
        return VBAInt(bgr(effects.glow_rgb) if effects.glow_rgb else 0, 'Long')

    def set_colour(self, colour: str) -> None:
        effects = self.change()
        effects.glow, effects.glow_rgb, effects.glow_scheme = True, colour, None

    def set_scheme(self, scheme: str) -> None:
        effects = self.change()
        effects.glow, effects.glow_rgb, effects.glow_scheme = True, None, scheme

    @member
    def Transparency(self) -> object:
        effects = self.held()
        if not effects.glow:
            return VBASingle(float(_MIXED))
        return VBASingle(0.0 if effects.glow_alpha is None else (100000 - effects.glow_alpha) / 100000)

    @setter('Transparency')
    def _transparency(self, value: object) -> None:
        wanted = _fraction(value)
        effects = self.change()
        effects.glow, effects.glow_alpha = True, round((1 - wanted) * 100000)


class SoftEdgeFormat(_Effect):
    vba_type_name = 'SoftEdgeFormat'

    @member
    def Parent(self) -> object:
        raise error(438, "Object doesn't support this property or method")

    @member
    def Radius(self) -> object:
        return VBASingle(self.held().soft_edge / _EMU_PER_POINT)

    @setter('Radius')
    def _radius(self, value: object) -> None:
        wanted = float(to_number(value))
        if not 0 <= wanted <= 100:
            raise error(*_OUT_OF_RANGE)
        self.change().soft_edge = round(wanted * _EMU_PER_POINT)

    @member
    def Type(self) -> object:
        radius = self.held().soft_edge
        return VBAInt(next((kind for kind, size in _SOFT_EDGES.items() if size == radius), 0), 'Long')

    @setter('Type')
    def _type(self, value: object) -> None:
        wanted = int(to_integer(value, 'Long'))
        if wanted != 0 and wanted not in _SOFT_EDGES:
            raise error(*_OUT_OF_RANGE)
        self.change().soft_edge = _SOFT_EDGES.get(wanted, 0)


class ReflectionFormat(_Effect):
    vba_type_name = 'ReflectionFormat'

    @member
    def Type(self) -> object:
        return VBAInt(self.held().reflection_type, 'Long')

    @setter('Type')
    def _type(self, value: object) -> None:
        wanted = int(to_integer(value, 'Long'))
        if wanted != 0 and wanted not in _REFLECTIONS:
            raise error(*_OUT_OF_RANGE)
        effects = self.change()
        effects.reflection_type = wanted
        effects.reflection = wanted != 0
        start, end, place, distance = _REFLECTIONS.get(wanted, (None, None, None, None))
        effects.reflection_start, effects.reflection_end, effects.reflection_place = start, end, place
        effects.reflection_distance = distance or None
        effects.reflection_blur = 6350 if wanted else None

    def worked(self) -> ShapeEffects:
        """The effects, with the reflection its own, no longer one of the presets."""
        effects = self.change()
        effects.reflection, effects.reflection_type = True, -2
        return effects

    @member
    def Transparency(self) -> object:
        effects = self.held()
        if not effects.reflection:
            return VBASingle(float(_MIXED))
        return VBASingle(0.0 if effects.reflection_start is None else (100000 - effects.reflection_start) / 100000)

    @setter('Transparency')
    def _transparency(self, value: object) -> None:
        wanted = _fraction(value)
        self.worked().reflection_start = round((1 - wanted) * 100000)

    @member
    def Size(self) -> object:
        place = self.held().reflection_place
        return VBASingle(100.0 if place is None else place / 1000)

    @setter('Size')
    def _size(self, value: object) -> None:
        wanted = float(to_number(value))
        if not 0 <= wanted <= 100:
            raise error(*_OUT_OF_RANGE)
        self.worked().reflection_place = round(wanted * 1000)

    @member
    def Offset(self) -> object:
        effects = self.held()
        if not effects.reflection:
            return VBASingle(float(_MIXED))
        return VBASingle((effects.reflection_distance or 0) / _EMU_PER_POINT)

    @setter('Offset')
    def _offset(self, value: object) -> None:
        # Less than nothing is taken as nothing.
        self.worked().reflection_distance = round(max(float(to_number(value)), 0.0) * _EMU_PER_POINT) or None

    @member
    def Blur(self) -> object:
        return VBASingle((self.held().reflection_blur or 0) / _EMU_PER_POINT)

    @setter('Blur')
    def _blur(self, value: object) -> None:
        wanted = float(to_number(value))
        if wanted < 0:
            raise error(*_OUT_OF_RANGE)
        self.worked().reflection_blur = round(wanted * _EMU_PER_POINT) or None


class _RangeEffect(_Effect):
    """An effect through a ShapeRange (tests/fixtures/shape_range_effects.json): a set reaches every shape, and a read
    gives what they share, or the mixed value for its type."""

    #: The format each of its shapes has, and what a measurement they do not share reads.
    single: type[_Effect] = _Effect
    mixed_measure = -2147483648.0

    def __init__(self, shapes: ShapeRange) -> None:
        super().__init__(shapes.shapes[0])
        self.range = shapes

    def parts(self) -> list[_Effect]:
        return [self.single(one) for one in self.range.shapes]

    def vba_get(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        if name == 'Parent':
            return self.range
        if name == 'Application':
            return super().vba_get(name, args, named)
        found = [part.vba_get(name, args, named) for part in self.parts()]
        if isinstance(found[0], EffectColor):
            return RangeEffectColor(self, name)
        return self.shared(name, found)

    def shared(self, name: str, found: list[object]) -> object:
        if all(one == found[0] for one in found[1:]):
            return found[0]
        return VBASingle(self.mixed_measure) if isinstance(found[0], VBASingle) else VBAInt(-2, 'Long')

    def vba_set(self, name: str, value: object, args: Sequence[object] = (), named: dict[str, object] | None = None,
                *, by_ref: bool = False) -> None:
        for part in self.parts():
            part.vba_set(name, value, args, named, by_ref=by_ref)


class RangeEffectColor(Created):
    """A range's shadow or glow colour."""

    vba_type_name = 'ColorFormat'

    def __init__(self, effect: _RangeEffect, name: str) -> None:
        self.effect, self.name = effect, name

    def colours(self) -> list[EffectColor]:
        return [cast(EffectColor, part.vba_get(self.name)) for part in self.effect.parts()]

    def vba_get(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        found = [one.vba_get(name, args, named) for one in self.colours()]
        if all(one == found[0] for one in found[1:]):
            return found[0]
        return VBAInt(-2147483648 if name == 'RGB' else -2, 'Long')

    def vba_set(self, name: str, value: object, args: Sequence[object] = (), named: dict[str, object] | None = None,
                *, by_ref: bool = False) -> None:
        for one in self.colours():
            one.vba_set(name, value, args, named, by_ref=by_ref)


class RangeShadow(_RangeEffect, ShadowFormat):
    """Any shape's shadow showing makes the range's Visible true. While none shows, a read gives the last shape's."""

    vba_type_name = 'ShadowFormat'
    single = ShadowFormat

    def shared(self, name: str, found: list[object]) -> object:
        visible = [part.vba_get('Visible') for part in self.parts()]
        if name == 'Visible':
            return VBAInt(-1 if any(one == -1 for one in visible) else 0, 'Long')
        if not any(one == -1 for one in visible):
            return found[-1]
        if len(set(visible)) > 1:
            # One showing beside one hidden: these do not count as shared even where they are equal.
            if name in ('OffsetX', 'OffsetY'):
                return VBASingle(-2147483648.0)
            if name == 'RotateWithShape':
                return VBAInt(-2, 'Long')
        return super().shared(name, found)


class RangeGlow(_RangeEffect, GlowFormat):
    """A glow on one shape and none on another reads its Radius as mixed, even where neither has one set."""

    vba_type_name = 'GlowFormat'
    single = GlowFormat

    def shared(self, name: str, found: list[object]) -> object:
        if name == 'Radius' and len({part.held().glow for part in self.parts() if isinstance(part, GlowFormat)}) > 1:
            return VBASingle(self.mixed_measure)
        return super().shared(name, found)


class RangeSoftEdge(_RangeEffect, SoftEdgeFormat):
    """Radii that differ read -1, and types the last shape's."""

    vba_type_name = 'SoftEdgeFormat'
    single = SoftEdgeFormat
    mixed_measure = -1.0

    def shared(self, name: str, found: list[object]) -> object:
        return found[-1] if name == 'Type' else super().shared(name, found)


class RangeReflection(_RangeEffect, ReflectionFormat):
    """Types that differ read mixed, and measurements the last shape's."""

    vba_type_name = 'ReflectionFormat'
    single = ReflectionFormat

    def shared(self, name: str, found: list[object]) -> object:
        return found[-1] if isinstance(found[0], VBASingle) else super().shared(name, found)
