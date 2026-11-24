"""Shape.ThreeD and ShapeRange.ThreeD, as live Excel answers them (tests/fixtures/three_d_round_1.json).

A shape nothing was set on writes nothing and reads fixed answers. The first change writes a scene3d: a camera straight on
and three-point light from the top. A change to the camera or the light writes only that. A change to the body (a bevel,
the depth, the contour, the extrusion colour, Z or ProjectText) also writes an sp3d, and the material reads a real value
from then on. Hiding the format keeps both in the file's extensions, and every member reads as if nothing were set."""
from __future__ import annotations

import math
import struct
from collections.abc import Sequence
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._creator import Created
from pyopenvba.apps.excel._held import HeldPart
from pyopenvba.apps.excel._shape_format import STYLE, DEFAULT_THEME
from pyopenvba.apps.excel._styles import bgr
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import method, member, setter
from pyopenvba.interpreter._values import EMPTY, VBAInt, VBASingle, error, to_integer, to_number
from pyopenvba.shapes._values import FillLine, ShapeThreeD

if TYPE_CHECKING:
    from pyopenvba.apps.excel._shapes import ShapeObject, ShapeRange

_EMU_PER_POINT = 12700
_ANGLE = 60000
_TURN = 360 * _ANGLE
_MIXED = -2147483648
_BAD_VALUE = (-2147024809, 'The parameter is incorrect.')
#: A bevel's depth and inset until set, in EMU.
_BEVEL_SIZE = 76200
#: The depth an extrusion colour gives a shape that has none.
_COLOURED_DEPTH = 76200
#: BevelTopType and BevelBottomType by the prst the file writes; circle is the default and writes none.
_BEVELS = {2: 'relaxedInset', 3: None, 4: 'slope', 5: 'cross', 6: 'angle', 7: 'softRound', 8: 'convex', 9: 'coolSlant', 10: 'divot',
           11: 'riblet', 12: 'hardEdge', 13: 'artDeco'}
#: What each member reads where nothing is set or the format is hidden.
_UNSET: dict[str, object] = {
    'Visible': VBAInt(0, 'Long'), 'BevelTopType': VBAInt(1, 'Long'), 'BevelBottomType': VBAInt(1, 'Long'),
    'BevelTopDepth': VBASingle(0.0), 'BevelTopInset': VBASingle(0.0), 'BevelBottomDepth': VBASingle(0.0), 'BevelBottomInset': VBASingle(0.0),
    'Depth': VBASingle(36.0), 'ContourWidth': VBASingle(0.0), 'ExtrusionColorType': VBAInt(1, 'Long'), 'FieldOfView': VBASingle(45.0),
    'LightAngle': VBASingle(0.0), 'Perspective': VBAInt(-2, 'Long'), 'PresetCamera': VBAInt(-2, 'Long'),
    'PresetExtrusionDirection': VBAInt(-2, 'Long'), 'PresetLighting': VBAInt(-2, 'Long'), 'PresetLightingDirection': VBAInt(-2, 'Long'),
    'PresetLightingSoftness': VBAInt(-2, 'Long'), 'PresetMaterial': VBAInt(-2, 'Long'), 'PresetThreeDFormat': VBAInt(-2, 'Long'),
    'ProjectText': VBAInt(-1, 'Long'), 'RotationX': VBASingle(0.0), 'RotationY': VBASingle(0.0), 'RotationZ': VBASingle(0.0),
    'Z': VBASingle(float(_MIXED)),
}
_COLOURS_UNSET = {'ContourColor': 0, 'ExtrusionColor': 16777215}


def _single(value: float) -> float:
    return float(struct.unpack('f', struct.pack('f', value))[0])


def _degrees(angle: int, sign: int) -> VBASingle:
    """An angle of the file's as a rotation reads it, from -180 up to 180."""
    turned = sign * angle / _ANGLE
    return VBASingle((turned + 180) % 360 - 180)


def _stored(degrees: float) -> int:
    return round((degrees % 360) * _ANGLE) % _TURN


def _bounded(value: object, largest: float, refused: float, what: str) -> float:
    """A size Excel takes up to its largest, refusing with the parameter error at the value measured; between them the
    answer is not known."""
    wanted = float(VBASingle(float(to_number(value))))
    if wanted >= refused:
        raise error(*_BAD_VALUE)
    if wanted > largest or wanted < 0:
        raise VBAUnsupportedError(f'{what} of {wanted} is not implemented')
    return wanted


class ThreeDFormat(HeldPart):
    vba_type_name = 'ThreeDFormat'

    def __init__(self, owner: ShapeObject) -> None:
        self.owner = owner

    def admit(self) -> None:
        shape = self.owner.shape
        if shape.source or shape.kind not in STYLE:
            raise VBAUnsupportedError('The ThreeD of this shape is not implemented')
        if self.owner.sheet.book.stylesheet.colors.theme != DEFAULT_THEME:
            raise VBAUnsupportedError('The ThreeD of a shape under another theme is not implemented')

    def shown(self) -> ShapeThreeD | None:
        """What is set, for a read, while the format shows."""
        self.admit()
        look = self.owner.shape.fill_line
        three_d = None if look is None else look.three_d
        return three_d if three_d is not None and three_d.visible else None

    def change(self, *, body: bool = False) -> ShapeThreeD:
        self.admit()
        look = self.owner.shape.fill_line
        if look is not None and look.three_d is not None and not look.three_d.visible:
            raise VBAUnsupportedError('A change to a hidden ThreeD is not implemented')
        self.owner.guard('ThreeD')
        shape = self.owner.shape
        if shape.fill_line is None:
            shape.fill_line = FillLine()
        if shape.fill_line.three_d is None:
            shape.fill_line.three_d = ShapeThreeD()
        self.owner.restyled()
        three_d = shape.fill_line.three_d
        three_d.body = three_d.body or body
        return three_d

    @member
    def Application(self) -> object:
        return self.owner.sheet.book.application

    @member
    def Parent(self) -> object:
        return self.owner

    @member
    def Visible(self) -> object:
        return VBAInt(0 if self.shown() is None else -1, 'Long')

    @setter('Visible')
    def _visible(self, value: object) -> None:
        self.admit()
        look = self.owner.shape.fill_line
        three_d = None if look is None else look.three_d
        if three_d is None:
            # Neither way changes a format nothing was set on.
            return
        self.owner.guard('ThreeD')
        three_d.visible = int(to_integer(value, 'Long')) != 0
        self.owner.restyled()

    def _bevel(self, top: bool) -> tuple[str | None, int | None, int | None] | None:
        three_d = self.shown()
        return None if three_d is None else (three_d.bevel_top if top else three_d.bevel_bottom)

    def _set_bevel(self, top: bool, at: int, value: object) -> None:
        three_d = self.change(body=True)
        was = (three_d.bevel_top if top else three_d.bevel_bottom) or (None, None, None)
        bevel = (*was[:at], value, *was[at + 1:])
        if top:
            three_d.bevel_top = bevel  # type: ignore[assignment]
        else:
            three_d.bevel_bottom = bevel  # type: ignore[assignment]

    def _bevel_type(self, top: bool) -> object:
        bevel = self._bevel(top)
        if bevel is None:
            return VBAInt(1, 'Long')
        return VBAInt(next(kind for kind, preset in _BEVELS.items() if preset == bevel[0]), 'Long')

    def _set_bevel_type(self, top: bool, value: object) -> None:
        wanted = int(to_integer(value, 'Long'))
        if wanted not in _BEVELS:
            raise VBAUnsupportedError(f'A bevel of type {wanted} is not implemented')
        self._set_bevel(top, 0, _BEVELS[wanted])

    def _bevel_size(self, top: bool, at: int) -> object:
        bevel = self._bevel(top)
        if bevel is None:
            return VBASingle(0.0)
        size = bevel[at]
        return VBASingle((_BEVEL_SIZE if size is None else size) / _EMU_PER_POINT)

    def _set_bevel_size(self, top: bool, at: int, value: object) -> None:
        wanted = _bounded(value, 1584, 2000, 'A bevel')
        self._set_bevel(top, at, round(wanted * _EMU_PER_POINT))

    @member
    def BevelTopType(self) -> object:
        return self._bevel_type(True)

    @setter('BevelTopType')
    def _bevel_top_type(self, value: object) -> None:
        self._set_bevel_type(True, value)

    @member
    def BevelTopInset(self) -> object:
        return self._bevel_size(True, 1)

    @setter('BevelTopInset')
    def _bevel_top_inset(self, value: object) -> None:
        self._set_bevel_size(True, 1, value)

    @member
    def BevelTopDepth(self) -> object:
        return self._bevel_size(True, 2)

    @setter('BevelTopDepth')
    def _bevel_top_depth(self, value: object) -> None:
        self._set_bevel_size(True, 2, value)

    @member
    def BevelBottomType(self) -> object:
        return self._bevel_type(False)

    @setter('BevelBottomType')
    def _bevel_bottom_type(self, value: object) -> None:
        self._set_bevel_type(False, value)

    @member
    def BevelBottomInset(self) -> object:
        return self._bevel_size(False, 1)

    @setter('BevelBottomInset')
    def _bevel_bottom_inset(self, value: object) -> None:
        self._set_bevel_size(False, 1, value)

    @member
    def BevelBottomDepth(self) -> object:
        return self._bevel_size(False, 2)

    @setter('BevelBottomDepth')
    def _bevel_bottom_depth(self, value: object) -> None:
        self._set_bevel_size(False, 2, value)

    @member
    def Depth(self) -> object:
        three_d = self.shown()
        if three_d is None or not three_d.body:
            return _UNSET['Depth']
        return VBASingle((three_d.extrusion_h or 0) / _EMU_PER_POINT)

    @setter('Depth')
    def _depth(self, value: object) -> None:
        wanted = float(VBASingle(float(to_number(value))))
        if wanted < 0 and wanted != -1:
            raise VBAUnsupportedError(f'A depth of {wanted} is not implemented')
        # -1 writes 1 pt.
        size = 1.0 if wanted == -1 else _bounded(wanted, 1584, 4000, 'A depth')
        self.change(body=True).extrusion_h = round(size * _EMU_PER_POINT)

    @member
    def ContourWidth(self) -> object:
        three_d = self.shown()
        return VBASingle(0.0 if three_d is None else (three_d.contour_w or 0) / _EMU_PER_POINT)

    @setter('ContourWidth')
    def _contour_width(self, value: object) -> None:
        wanted = _bounded(value, 1584, 4000, 'A contour')
        self.change(body=True).contour_w = round(wanted * _EMU_PER_POINT)

    @member
    def ContourColor(self) -> object:
        self.admit()
        return ThreeDColor(self, 'ContourColor')

    @member
    def ExtrusionColor(self) -> object:
        self.admit()
        return ThreeDColor(self, 'ExtrusionColor')

    def read_colour(self, which: str) -> int:
        three_d = self.shown()
        rgb = None if three_d is None else (three_d.contour_rgb if which == 'ContourColor' else three_d.extrusion_rgb)
        return _COLOURS_UNSET[which] if rgb is None else bgr(rgb)

    def set_colour(self, which: str, rgb: str) -> None:
        three_d = self.change(body=True)
        if which == 'ContourColor':
            three_d.contour_rgb = rgb
            if three_d.contour_w is None:
                three_d.contour_w = _EMU_PER_POINT
        else:
            three_d.extrusion_rgb = rgb
            if three_d.extrusion_h is None:
                three_d.extrusion_h = _COLOURED_DEPTH

    @member
    def ExtrusionColorType(self) -> object:
        three_d = self.shown()
        return VBAInt(2 if three_d is not None and three_d.extrusion_rgb is not None else 1, 'Long')

    @setter('ExtrusionColorType')
    def _extrusion_color_type(self, value: object) -> None:
        wanted = int(to_integer(value, 'Long'))
        three_d = self.shown()
        now = 2 if three_d is not None and three_d.extrusion_rgb is not None else 1
        # Custom with no colour yet changes nothing, as setting what it already is does.
        if wanted != now and not (wanted == 2 and now == 1):
            raise VBAUnsupportedError('An extrusion colour type set back to automatic is not implemented')

    @member
    def FieldOfView(self) -> object:
        three_d = self.shown()
        return VBASingle(65.0 if three_d is not None and three_d.perspective else 45.0)

    @setter('FieldOfView')
    def _field_of_view(self, value: object) -> None:
        _bounded(value, 180, 200, 'A field of view')
        three_d = self.change()
        if three_d.perspective:
            raise VBAUnsupportedError('The field of view of a perspective camera is not implemented')

    @member
    def LightAngle(self) -> object:
        three_d = self.shown()
        return VBASingle(0.0 if three_d is None or three_d.light_rev is None else three_d.light_rev / _ANGLE)

    @setter('LightAngle')
    def _light_angle(self, value: object) -> None:
        wanted = float(VBASingle(float(to_number(value))))
        if not 0 <= wanted < 360:
            raise VBAUnsupportedError(f'A light angle of {wanted} is not implemented')
        self.change().light_rev = round(wanted * _ANGLE)

    @member
    def Perspective(self) -> object:
        three_d = self.shown()
        return _UNSET['Perspective'] if three_d is None else VBAInt(-1 if three_d.perspective else 0, 'Long')

    @setter('Perspective')
    def _perspective(self, value: object) -> None:
        if int(to_integer(value, 'Long')) == 0:
            raise VBAUnsupportedError('Perspective set off is not implemented')
        three_d = self.change()
        if three_d.rotation is not None:
            raise VBAUnsupportedError('Perspective on a turned shape is not implemented')
        three_d.perspective = True

    @member
    def PresetCamera(self) -> object:
        three_d = self.shown()
        return _UNSET['PresetCamera'] if three_d is None else VBAInt(14 if three_d.perspective else 19, 'Long')

    @member
    def PresetExtrusionDirection(self) -> object:
        return _UNSET['PresetExtrusionDirection'] if self.shown() is None else VBAInt(5, 'Long')

    @member
    def PresetLighting(self) -> object:
        return _UNSET['PresetLighting'] if self.shown() is None else VBAInt(13, 'Long')

    @member
    def PresetLightingDirection(self) -> object:
        three_d = self.shown()
        return VBAInt(2 if three_d is not None and three_d.light_rev is None else -2, 'Long')

    @member
    def PresetLightingSoftness(self) -> object:
        self.admit()
        return VBAInt(-2, 'Long')

    @member
    def PresetMaterial(self) -> object:
        three_d = self.shown()
        return VBAInt(8 if three_d is not None and three_d.body else -2, 'Long')

    @member
    def PresetThreeDFormat(self) -> object:
        self.admit()
        return VBAInt(-2, 'Long')

    @member
    def ProjectText(self) -> object:
        self.admit()
        return VBAInt(-1, 'Long')

    @setter('ProjectText')
    def _project_text(self, value: object) -> None:
        if int(to_integer(value, 'Long')) == 0:
            raise VBAUnsupportedError('ProjectText set off is not implemented')
        self.change(body=True)

    @member
    def Z(self) -> object:
        three_d = self.shown()
        if three_d is None or not three_d.body:
            return _UNSET['Z']
        return VBASingle((three_d.z or 0) / _EMU_PER_POINT)

    @setter('Z')
    def _z(self, value: object) -> None:
        wanted = float(VBASingle(float(to_number(value))))
        if not 0 <= wanted <= 1584:
            raise VBAUnsupportedError(f'A Z of {wanted} is not implemented')
        self.change(body=True).z = round(wanted * _EMU_PER_POINT)

    # The camera's rotation: RotationX turns its lon the other way, RotationY is its lat and RotationZ its rev the other way.

    def _rotation(self) -> tuple[int, int, int]:
        three_d = self.shown()
        return (0, 0, 0) if three_d is None or three_d.rotation is None else three_d.rotation

    def _turn(self, at: int, angle: int) -> None:
        three_d = self.change()
        rotation = list(three_d.rotation or (0, 0, 0))
        rotation[at] = angle
        three_d.rotation = (rotation[0], rotation[1], rotation[2])

    @member
    def RotationX(self) -> object:
        return _degrees(self._rotation()[1], -1)

    @setter('RotationX')
    def _rotation_x(self, value: object) -> None:
        self._turn(1, _stored(-float(VBASingle(float(to_number(value))))))

    @member
    def RotationY(self) -> object:
        return _degrees(self._rotation()[0], 1)

    @setter('RotationY')
    def _rotation_y(self, value: object) -> None:
        self._turn(0, _stored(float(VBASingle(float(to_number(value))))))

    @member
    def RotationZ(self) -> object:
        return _degrees(self._rotation()[2], -1)

    @setter('RotationZ')
    def _rotation_z(self, value: object) -> None:
        self._turn(2, _stored(-float(VBASingle(float(to_number(value))))))

    def _turned(self, at: int, sign: int) -> float:
        return float(_degrees(self._rotation()[at], sign))

    @method
    def IncrementRotationX(self, Increment: object) -> object:  # noqa: N803 - native named argument
        self._turn(1, _stored(-(self._turned(1, -1) + float(to_number(Increment)))))
        return EMPTY

    @method
    def IncrementRotationY(self, Increment: object) -> object:  # noqa: N803 - native named argument
        self._turn(0, _stored(self._turned(0, 1) + float(to_number(Increment))))
        return EMPTY

    @method
    def IncrementRotationZ(self, Increment: object) -> object:  # noqa: N803 - native named argument
        self._turn(2, _stored(-(self._turned(2, -1) + float(to_number(Increment)))))
        return EMPTY

    @method
    def IncrementRotationHorizontal(self, Increment: object) -> object:  # noqa: N803 - native named argument
        self._turn(1, _stored(-(self._turned(1, -1) - float(to_number(Increment)))))
        return EMPTY

    @method
    def IncrementRotationVertical(self, Increment: object) -> object:  # noqa: N803 - native named argument
        # Through radians in Single and cut, not rounded: 10 from nothing writes lat 20999999, read as -10.00002.
        turned = _single(_single(float(to_number(Increment))) * _single(math.pi)) / 180 * 180 / math.pi
        degrees = (self._turned(0, 1) - turned) % 360
        self._turn(0, int(degrees * _ANGLE) % _TURN)
        return EMPTY

    @method
    def ResetRotation(self) -> object:
        look = self.owner.shape.fill_line
        if look is None or look.three_d is None:
            raise VBAUnsupportedError('ResetRotation on a shape with no ThreeD is not implemented')
        self.change().rotation = None
        return EMPTY

    @method
    def SetThreeDFormat(self, PresetThreeDFormat: object) -> object:  # noqa: N803 - native named argument
        raise VBAUnsupportedError('ThreeDFormat.SetThreeDFormat is not implemented')

    @method
    def SetExtrusionDirection(self, PresetExtrusionDirection: object) -> object:  # noqa: N803 - native named argument
        raise VBAUnsupportedError('ThreeDFormat.SetExtrusionDirection is not implemented')

    @method
    def SetPresetCamera(self, PresetCamera: object) -> object:  # noqa: N803 - native named argument
        raise VBAUnsupportedError('ThreeDFormat.SetPresetCamera is not implemented')

    def vba_set(self, name: str, value: object, args: Sequence[object] = (), named: dict[str, object] | None = None,
                *, by_ref: bool = False) -> None:
        if name.lower() in _PRESETS:
            raise VBAUnsupportedError(f'ThreeDFormat.{name} set is not implemented')
        super().vba_set(name, value, args, named, by_ref=by_ref)


_PRESETS = frozenset(name.lower() for name in ('PresetCamera', 'PresetExtrusionDirection', 'PresetLighting', 'PresetLightingDirection',
                                               'PresetLightingSoftness', 'PresetMaterial', 'PresetThreeDFormat'))


class ThreeDColor(HeldPart):
    """A ThreeD's contour or extrusion colour: its RGB alone."""

    vba_type_name = 'ColorFormat'

    def __init__(self, owner: ThreeDFormat, which: str) -> None:
        self.owner, self.which = owner, which

    @member
    def RGB(self) -> object:
        return VBAInt(self.owner.read_colour(self.which), 'Long')

    @setter('RGB')
    def _rgb(self, value: object) -> None:
        colour = int(to_integer(value, 'Long')) & 0xFFFFFF
        red, green, blue = colour & 255, (colour >> 8) & 255, (colour >> 16) & 255
        self.owner.set_colour(self.which, f'{red:02X}{green:02X}{blue:02X}')

    def vba_get(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        if name.lower() != 'rgb':
            raise VBAUnsupportedError(f'ColorFormat.{name} of a ThreeD is not implemented')
        return super().vba_get(name, args, named)

    def vba_set(self, name: str, value: object, args: Sequence[object] = (), named: dict[str, object] | None = None,
                *, by_ref: bool = False) -> None:
        if name.lower() != 'rgb':
            raise VBAUnsupportedError(f'ColorFormat.{name} of a ThreeD is not implemented')
        super().vba_set(name, value, args, named, by_ref=by_ref)


class RangeThreeD(Created):
    """A ShapeRange's ThreeD: a set or a method reaches every shape, and a read gives what they share, or the mixed value for its
    type. Not measured on a range yet."""

    vba_type_name = 'ThreeDFormat'

    def __init__(self, shapes: ShapeRange) -> None:
        self.range = shapes

    def parts(self) -> list[ThreeDFormat]:
        return [ThreeDFormat(one) for one in self.range.shapes]

    def vba_get(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        if name.lower() == 'parent':
            return self.range
        found = [part.vba_get(name, args, named) for part in self.parts()]
        if name.lower() == 'application':
            return found[0]
        if isinstance(found[0], ThreeDColor):
            raise VBAUnsupportedError(f'ThreeDFormat.{name} of a range is not implemented')
        if all(one == found[0] for one in found[1:]):
            return found[0]
        return VBASingle(float(_MIXED)) if isinstance(found[0], VBASingle) else VBAInt(-2, 'Long')

    def vba_set(self, name: str, value: object, args: Sequence[object] = (), named: dict[str, object] | None = None,
                *, by_ref: bool = False) -> None:
        for part in self.parts():
            part.vba_set(name, value, args, named, by_ref=by_ref)

    def vba_call(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        for part in self.parts():
            part.vba_call(name, args, named)
        return EMPTY
