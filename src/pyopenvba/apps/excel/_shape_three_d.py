"""Shape.ThreeD and ShapeRange.ThreeD, as live Excel answers them (tests/fixtures/three_d_round_1.json, three_d_round_2.json).

A shape nothing was set on writes nothing and reads fixed answers. The file has two parts, each written once a member of it was
set: the scene3d, a camera and a light rig, and the sp3d, the body (a bevel, the depth, the contour, the material, the
extrusion colour, Z or ProjectText). Visible reads whether the scene3d is written. The first change but BevelTopType = msoBevelNone
writes a scene3d with a camera straight on and three-point light from the top. A preset of the old 3-D settings (the light's
direction or softness, the material, the extrusion direction) set on a plain shape starts from Excel's old default: an oblique
camera from the top right, flat light from the right, a 36 pt extrusion of matte and an angle bevel each side. Hiding the
format puts both parts in the file's extensions; what is set after that is written as a part of its own."""
from __future__ import annotations

import struct
from collections.abc import Sequence
from dataclasses import replace
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._creator import Created
from pyopenvba.apps.excel._held import HeldPart
from pyopenvba.apps.excel._shape_format import STYLE, DEFAULT_THEME
from pyopenvba.apps.excel._styles import bgr
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import BINDING, method, member, setter
from pyopenvba.interpreter._values import EMPTY, NULL, VBAInt, VBASingle, error, to_integer, to_number
from pyopenvba.shapes._read_look import read_three_d
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
#: PresetCamera: the camera's prst, the field of view it reads, whether it reads as Perspective, the PresetExtrusionDirection it
#: reads, and the rotations it reads where the file has no rot of its own.
_CAMERAS: dict[int, tuple[str, int, bool, int, tuple[float, float, float] | None]] = {
    1: ('legacyObliqueTopLeft', 45, False, 9, None), 2: ('legacyObliqueTop', 45, False, 8, None),
    3: ('legacyObliqueTopRight', 45, False, 7, None), 4: ('legacyObliqueLeft', 45, False, 6, None),
    5: ('legacyObliqueFront', 45, False, 5, None), 6: ('legacyObliqueRight', 45, False, 4, None),
    7: ('legacyObliqueBottomLeft', 45, False, 3, None), 8: ('legacyObliqueBottom', 45, False, 2, None),
    9: ('legacyObliqueBottomRight', 45, False, 1, None), 10: ('legacyPerspectiveTopLeft', 65, True, 9, None),
    11: ('legacyPerspectiveTop', 65, True, 8, None), 12: ('legacyPerspectiveTopRight', 65, True, 7, None),
    13: ('legacyPerspectiveLeft', 65, True, 6, None), 14: ('legacyPerspectiveFront', 65, True, 5, None),
    15: ('legacyPerspectiveRight', 65, True, 4, None), 16: ('legacyPerspectiveBottomLeft', 65, True, 3, None),
    17: ('legacyPerspectiveBottom', 65, True, 2, None), 18: ('legacyPerspectiveBottomRight', 65, True, 1, None),
    19: ('orthographicFront', 45, False, 5, None),
    20: ('isometricTopUp', 45, False, -2, (45.2809, -35.39627, -60.16242)),
    21: ('isometricTopDown', 45, False, -2, (-45.28088, -35.39627, 60.16245)),
    22: ('isometricBottomUp', 45, False, -2, (-45.28088, 35.39627, -60.16242)),
    23: ('isometricBottomDown', 45, False, -2, (45.2809, 35.39627, 60.16245)),
    24: ('isometricLeftUp', 45, False, 1, (-45.0, -35.00002, 0.0)), 25: ('isometricLeftDown', 45, False, 7, (-45.0, 35.0, 0.0)),
    26: ('isometricRightUp', 45, False, 3, (45.00003, 35.0, 0.0)), 27: ('isometricRightDown', 45, False, 9, (45.00003, -35.00002, 0.0)),
    28: ('isometricOffAxis1Left', 45, False, 6, (-64.0, 18.0, 0.0)), 29: ('isometricOffAxis1Right', 45, False, 4, (26.00003, 18.0, 0.0)),
    30: ('isometricOffAxis1Top', 45, False, 8, (53.45425, -58.73812, -57.6425)),
    31: ('isometricOffAxis2Left', 45, False, 6, (-26.0, 18.0, 0.0)), 32: ('isometricOffAxis2Right', 45, False, 4, (64.00003, 18.0, 0.0)),
    33: ('isometricOffAxis2Top', 45, False, 8, (-53.45424, -58.73812, 57.6425)),
    34: ('isometricOffAxis3Left', 45, False, 6, (-64.0, -18.00002, 0.0)),
    35: ('isometricOffAxis3Right', 45, False, 4, (26.00003, -18.00002, 0.0)),
    36: ('isometricOffAxis3Bottom', 45, False, 2, (53.45425, 58.73808, 57.64255)),
    37: ('isometricOffAxis4Left', 45, False, 6, (-26.0, -18.00002, 0.0)),
    38: ('isometricOffAxis4Right', 45, False, 4, (64.00003, -18.00002, 0.0)),
    39: ('isometricOffAxis4Bottom', 45, False, 2, (-53.45424, 58.73808, -57.64252)),
    40: ('obliqueTopLeft', 45, False, 9, None), 41: ('obliqueTop', 45, False, 8, None), 42: ('obliqueTopRight', 45, False, 7, None),
    43: ('obliqueLeft', 45, False, 6, None), 44: ('obliqueRight', 45, False, 4, None), 45: ('obliqueBottomLeft', 45, False, 3, None),
    46: ('obliqueBottom', 45, False, 2, None), 47: ('obliqueBottomRight', 45, False, 1, None),
    48: ('perspectiveFront', 45, True, 5, None), 49: ('perspectiveLeft', 45, True, 4, (-20.0, 0.0, 0.0)),
    50: ('perspectiveRight', 45, True, 6, (20.00002, 0.0, 0.0)), 51: ('perspectiveAbove', 45, True, 2, (0.0, -20.00002, 0.0)),
    52: ('perspectiveBelow', 45, True, 8, (0.0, 20.0, 0.0)),
    53: ('perspectiveAboveLeftFacing', 45, True, 7, (-14.33758, 39.30647, 18.87617)),
    54: ('perspectiveAboveRightFacing', 45, True, 9, (14.33762, 39.30647, -18.87615)),
    55: ('perspectiveContrastingLeftFacing', 45, True, 4, (-43.93887, 10.39642, 3.553517)),
    56: ('perspectiveContrastingRightFacing', 45, True, 6, (43.9389, 10.39642, -3.553517)),
    57: ('perspectiveHeroicLeftFacing', 45, True, 1, (-14.29302, -11.01627, -2.626817)),
    58: ('perspectiveHeroicRightFacing', 45, True, 3, (14.29307, -11.01627, 2.626767)),
    59: ('perspectiveHeroicExtremeLeftFacing', 80, True, 4, (-34.46068, 8.12245, 2.908583)),
    60: ('perspectiveHeroicExtremeRightFacing', 80, True, 6, (34.46073, 8.12245, -2.9086)),
    61: ('perspectiveRelaxed', 45, True, 2, (0.0, -50.43993, 0.0)),
    62: ('perspectiveRelaxedModerately', 45, True, 2, (0.0, -35.15602, 0.0)),
}
_CAMERA_NUMBERS = {camera[0]: number for number, camera in _CAMERAS.items()}
#: SetThreeDFormat: the camera, its rot, the extrusion and the material.
_FORMATS: dict[int, tuple[str, tuple[int, int, int] | None, int, str]] = {
    1: ('legacyObliqueTopRight', None, 457200, 'legacyMatte'), 2: ('legacyObliqueTopLeft', None, 457200, 'legacyMatte'),
    3: ('legacyObliqueBottomLeft', None, 457200, 'legacyMatte'), 4: ('legacyObliqueTopRight', None, 457200, 'legacyWireframe'),
    5: ('perspectiveFront', (20100000, 1500000, 0), 457200, 'legacyMatte'),
    6: ('perspectiveFront', (20100000, 20100000, 0), 457200, 'legacyMatte'),
    7: ('legacyPerspectiveBottom', None, 914400, 'legacyMatte'), 8: ('legacyPerspectiveBottomLeft', None, 914400, 'legacyMatte'),
    9: ('perspectiveFront', (1500000, 1500000, 0), 457200, 'legacyMatte'),
    10: ('perspectiveFront', (1500000, 20100000, 0), 457200, 'legacyMatte'),
    11: ('legacyPerspectiveTopRight', None, 914400, 'legacyMatte'), 12: ('legacyPerspectiveTop', None, 914400, 'legacyMatte'),
    13: ('perspectiveFront', (20100000, 1500000, 0), 914400, 'legacyMatte'),
    14: ('perspectiveFront', (20100000, 20100000, 0), 914400, 'legacyMatte'),
    15: ('legacyPerspectiveBottom', None, 121920000, 'legacyMatte'), 16: ('legacyPerspectiveBottomLeft', None, 121920000, 'legacyMatte'),
    17: ('perspectiveFront', (1500000, 1500000, 0), 914400, 'legacyMatte'),
    18: ('perspectiveFront', (1500000, 20100000, 0), 914400, 'legacyMatte'),
    19: ('legacyPerspectiveTopRight', None, 121920000, 'legacyMatte'), 20: ('legacyPerspectiveTop', None, 121920000, 'legacyMatte'),
}
#: PresetLighting by the light rig.
_RIGS = ['legacyFlat1', 'legacyFlat2', 'legacyFlat3', 'legacyFlat4', 'legacyNormal1', 'legacyNormal2', 'legacyNormal3', 'legacyNormal4',
         'legacyHarsh1', 'legacyHarsh2', 'legacyHarsh3', 'legacyHarsh4', 'threePt', 'balanced', 'soft', 'harsh', 'flood',
         'contrasting', 'morning', 'sunrise', 'sunset', 'chilly', 'freezing', 'flat', 'twoPt', 'glow', 'brightRoom']
#: PresetLightingSoftness by an old rig's kind.
_SOFTNESS = {'Harsh': 1, 'Normal': 2, 'Flat': 3}
#: An old rig's direction at dir t by its number, and the directions a quarter turn clockwise takes them through.
#: An old rig's PresetLightingDirection by its number and dir, each as Excel was seen to write or read it: not a turn of the rig,
#: as number 2 from the left reads left (tests/fixtures/three_d_round_6.json).
_OLD_DIRECTIONS = {(1, 't'): 5, (2, 't'): 1, (3, 't'): 4, (4, 't'): 7, (3, 'r'): 2, (3, 'b'): 6, (3, 'l'): 8, (4, 'b'): 3, (2, 'b'): 9,
                   (2, 'l'): 4, (3, 'tl'): 1}
#: A light's angle clockwise from the top by PresetLightingDirection, the front read as the top.
_DEGREES = {2: 0, 3: 45, 6: 90, 9: 135, 8: 180, 7: 225, 4: 270, 1: 315, 5: 0}
_QUARTERS = {'t': 0, 'r': 1, 'b': 2, 'l': 3}
#: PresetLightingDirection by eighths of a turn clockwise from the top.
_COMPASS = (2, 3, 6, 9, 8, 7, 4, 1)
#: PresetLightingDirection set on an old rig: the rig's number and dir Excel writes.
_DIRECTED = {1: (2, 't'), 2: (3, 'r'), 3: (4, 'b'), 4: (3, 't'), 5: (1, 't'), 6: (3, 'b'), 7: (4, 't'), 8: (3, 'l'), 9: (2, 'b')}
#: PresetMaterial by the file's prstMaterial; warm matte is the default and writes none.
_MATERIALS = ['legacyMatte', 'legacyPlastic', 'legacyMetal', 'legacyWireframe', 'matte', 'plastic', 'metal', None, 'translucentPowder',
              'powder', 'dkEdge', 'softEdge', 'clear', 'flat', 'softmetal']
#: SetExtrusionDirection: the side the old camera looks from.
_SIDES = {1: 'BottomRight', 2: 'Bottom', 3: 'BottomLeft', 4: 'Right', 5: 'Front', 6: 'Left', 7: 'TopRight', 8: 'Top', 9: 'TopLeft'}
_ACCENTS = {5: 'accent1', 6: 'accent2', 7: 'accent3', 8: 'accent4', 9: 'accent5', 10: 'accent6'}
_COLOURS_UNSET = {'ContourColor': 0, 'ExtrusionColor': 16777215}


def _single(value: object) -> float:
    return float(struct.unpack('f', struct.pack('f', float(VBASingle(float(to_number(value))))))[0])


def _degrees(angle: int, sign: int) -> float:
    """An angle of the file's as a rotation reads it, from -180 up to 180."""
    turned = sign * angle / _ANGLE
    return (turned + 180) % 360 - 180


def _stored(degrees: float) -> int:
    return round((degrees % 360) * _ANGLE) % _TURN


def _size(value: object, largest: float, refused: float, what: str) -> float:
    """A size Excel takes up to its largest, refusing with the parameter error from the value measured; between them, and
    below nothing, the answer is not known."""
    wanted = _single(value)
    if wanted >= refused or wanted == -1:
        # -1 was measured as True (tests/fixtures/three_d_round_11.json).
        raise error(*_BAD_VALUE)
    if wanted > largest or wanted < 0:
        raise VBAUnsupportedError(f'{what} of {wanted} is not implemented')
    return wanted


def _choice(value: object) -> int:
    """An enumerated member's value: msoFalse and msoTrue, as Empty and True are, are the parameter error for each of them
    (tests/fixtures/three_d_round_11.json to three_d_round_13.json)."""
    wanted = int(to_integer(value, 'Long'))
    if wanted in (0, -1):
        raise error(*_BAD_VALUE)
    return wanted


def _drop_hidden_scene(three_d: ShapeThreeD) -> None:
    """Perspective and SetThreeDFormat on a hidden format write a scene3d in place of the hidden one
    (tests/fixtures/three_d_round_12.json)."""
    if three_d.hidden is None:
        return
    three_d.scene = True
    _move(ShapeThreeD(), three_d.hidden, _PARTS[0][0])
    if three_d.hidden.blank():
        three_d.hidden = None


def _plain_rotation(camera: str, rotation: tuple[int, int, int] | None) -> tuple[int, int, int] | None:
    """A camera's rot, None where it is the one the camera reads without it."""
    implied = _CAMERAS[_CAMERA_NUMBERS[camera]][4] or (0.0, 0.0, 0.0)
    if rotation == (_stored(implied[1]), _stored(-implied[0]), _stored(-implied[2])):
        return None
    return rotation


def _old_default(three_d: ShapeThreeD) -> None:
    """Excel's old 3-D default, which the old presets start from in place of what was set."""
    _move(ShapeThreeD(scene=True, camera='legacyObliqueTopRight', rig='legacyFlat3', rig_dir='b', body=True, extrusion_h=457200,
                      material='legacyMatte', bevel_top=('angle', 13500, 13500), bevel_bottom=('angle', 13500, 13500)),
          three_d, _PARTS[0][0] + _PARTS[1][0])


def _loaded(source: str) -> ShapeThreeD | None:
    """What a shape's file holds of its ThreeD, where every name in it is one Excel was seen to write."""
    three_d = read_three_d(source)
    for one in (three_d, None if three_d is None else three_d.hidden):
        if one is None:
            continue
        bevels = [bevel[0] for bevel in (one.bevel_top, one.bevel_bottom) if bevel is not None]
        schemes = [scheme for scheme in (one.extrusion_scheme, one.contour_scheme) if scheme is not None]
        if (one.camera not in _CAMERA_NUMBERS or one.rig not in _RIGS or one.rig_dir not in _QUARTERS or one.material not in _MATERIALS
                or any(bevel not in _BEVELS.values() for bevel in bevels) or any(scheme not in _ACCENTS.values() for scheme in schemes)):
            raise VBAUnsupportedError("This ThreeD of a shape from a file is not implemented")
    if three_d is not None and three_d.body and three_d.hidden is not None and three_d.hidden.body:
        # Excel drops the hidden body when it opens such a file, and writes it so when it saves, changed or not.
        raise VBAUnsupportedError("A ThreeD from a file with a hidden body beside its own is not implemented")
    return three_d


#: The scene3d's members and the sp3d's, each with the flag that says it is written.
_PARTS = ((('scene', 'camera', 'fov', 'rotation', 'rig', 'rig_dir', 'light_rev'), 'scene'),
          (('body', 'z', 'extrusion_h', 'contour_w', 'material', 'bevel_top', 'bevel_bottom', 'extrusion_rgb', 'extrusion_scheme',
            'contour_rgb', 'contour_scheme'), 'body'))
#: ProjectText is the text's own and is in neither: hiding the format leaves it.


def _move(source: ShapeThreeD, target: ShapeThreeD, part: tuple[str, ...]) -> None:
    for name in part:
        setattr(target, name, getattr(source, name))


class ThreeDFormat(HeldPart):
    vba_type_name = 'ThreeDFormat'

    def __init__(self, owner: ShapeObject) -> None:
        self.owner = owner

    def vba_set(self, name: str, value: object, args: Sequence[object] = (), named: dict[str, object] | None = None,
                *, by_ref: bool = False) -> None:
        # Null reaches Excel late bound, which refuses it with 13; typed, VBA's own conversion refuses it with 94
        # (tests/fixtures/three_d_round_11.json).
        if value is NULL and BINDING.late:
            raise error(13)
        super().vba_set(name, value, args, named, by_ref=by_ref)

    def vba_get(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        # A method's call comes here too.
        if BINDING.late and any(one is NULL for one in [*args, *(named or {}).values()]):
            raise error(13)
        return super().vba_get(name, args, named)

    def admit(self) -> None:
        shape = self.owner.shape
        if shape.kind not in STYLE:
            raise VBAUnsupportedError('The ThreeD of this shape is not implemented')
        if self.owner.sheet.book.stylesheet.colors.theme != DEFAULT_THEME:
            raise VBAUnsupportedError('The ThreeD of a shape under another theme is not implemented')

    def held(self) -> ShapeThreeD | None:
        """What is set: a shape from a file holds what its file says until a change keeps it."""
        self.admit()
        shape = self.owner.shape
        look = shape.fill_line
        if look is not None and look.three_d is not None:
            return look.three_d
        return _loaded(shape.source) if shape.source else None

    def keep(self) -> ShapeThreeD:
        """What is set, kept on the shape for a change."""
        shape = self.owner.shape
        if shape.source and not any(shape is one for one in self.owner.sheet.shapes_):
            # A save writes the group's markup as it came, so the change would be lost.
            raise VBAUnsupportedError('Changing the ThreeD of a shape in a group from a file is not implemented')
        three_d = self.held() or ShapeThreeD()
        self.owner.guard('ThreeD')
        if shape.fill_line is None:
            shape.fill_line = FillLine()
        shape.fill_line.three_d = three_d
        return three_d

    def scene(self) -> ShapeThreeD | None:
        """What is set, where the scene3d is written."""
        three_d = self.held()
        return three_d if three_d is not None and three_d.scene else None

    def body(self) -> ShapeThreeD | None:
        """What is set, where the sp3d is written."""
        three_d = self.held()
        return three_d if three_d is not None and three_d.body else None

    def change(self, *, scene: bool = True, body: bool = False) -> ShapeThreeD:
        """What is set, for a change to the scene3d, the sp3d or both. A change to the body of a hidden format writes the sp3d
        alone; a change to its camera or light writes a scene3d of its own."""
        three_d = self.keep()
        if body:
            three_d.body = True
            three_d.scene = three_d.scene or scene and three_d.hidden is None
        elif scene:
            three_d.scene = True
        self.owner.restyled()
        return three_d

    def old_preset(self, needs: str, *, body: bool = False) -> ShapeThreeD | None:
        """What is set, for one of the old presets: from the old default where the part it needs, the scene3d or the sp3d, is
        not set. A light preset on a body with no scene3d changes nothing, and answers None."""
        held = self.held()
        if held is None or not getattr(held, needs):
            if held is not None and held.hidden is not None:
                raise VBAUnsupportedError('An old 3-D preset on a hidden ThreeD is not implemented')
            if needs == 'scene' and held is not None and held.body:
                return None
            three_d = self.change(body=True)
            _old_default(three_d)
            return three_d
        return self.change(body=body)

    @member
    def Application(self) -> object:
        return self.owner.sheet.book.application

    @member
    def Parent(self) -> object:
        return self.owner

    @member
    def Visible(self) -> object:
        return VBAInt(0 if self.scene() is None else -1, 'Long')

    @setter('Visible')
    def _visible(self, value: object) -> None:
        number = int(to_integer(value, 'Long'))
        three_d = self.held()
        if three_d is None:
            # Neither way changes a format nothing was set on.
            return
        if number in (2, 3):
            # Neither shows nor hides (tests/fixtures/three_d_round_12.json).
            return
        if number not in (0, -1):
            raise VBAUnsupportedError(f'Visible = {number} on a ThreeD that is set is not implemented')
        wanted = number == -1
        if wanted and (three_d.hidden is None or three_d.scene):
            # Showing a format that shows changes nothing, what is hidden staying hidden.
            return
        if not wanted and three_d.blank():
            return
        three_d = self.keep()
        if wanted:
            # Each part comes back where nothing was set in its place since; what was set wins.
            assert three_d.hidden is not None
            for part, flag in _PARTS:
                if getattr(three_d.hidden, flag) and not getattr(three_d, flag):
                    _move(three_d.hidden, three_d, part)
            three_d.hidden = None
        else:
            # Each part shown goes in place of what was hidden of it before.
            hidden = three_d.hidden or ShapeThreeD()
            for part, flag in _PARTS:
                if getattr(three_d, flag):
                    _move(three_d, hidden, part)
                    _move(ShapeThreeD(), three_d, part)
            three_d.hidden = hidden
        self.owner.restyled()

    # The bevels.

    def _bevel(self, top: bool) -> tuple[str | None, int | None, int | None] | None:
        three_d = self.body()
        return None if three_d is None else (three_d.bevel_top if top else three_d.bevel_bottom)

    def _put_bevel(self, three_d: ShapeThreeD, top: bool, bevel: tuple[str | None, int | None, int | None] | None) -> None:
        if top:
            three_d.bevel_top = bevel
        else:
            three_d.bevel_bottom = bevel

    def _set_bevel(self, top: bool, at: int, value: str | int | None) -> None:
        three_d = self.change(body=True)
        was = list((three_d.bevel_top if top else three_d.bevel_bottom) or (None, None, None))
        was[at] = value
        self._put_bevel(three_d, top, (was[0], was[1], was[2]))  # type: ignore[arg-type]

    def _bevel_type(self, top: bool) -> object:
        bevel = self._bevel(top)
        if bevel is None:
            return VBAInt(1, 'Long')
        return VBAInt(next(kind for kind, preset in _BEVELS.items() if preset == bevel[0]), 'Long')

    def _set_bevel_type(self, top: bool, value: object) -> None:
        wanted = _choice(value)
        if wanted == 1:
            # None takes the bevel away, and writes the sp3d without a scene of its own.
            self._put_bevel(self.change(scene=False, body=True), top, None)
            return
        if wanted not in _BEVELS:
            raise VBAUnsupportedError(f'A bevel of type {wanted} is not implemented')
        self._set_bevel(top, 0, _BEVELS[wanted])

    def _bevel_size(self, top: bool, at: int) -> object:
        bevel = self._bevel(top)
        if bevel is None:
            return VBASingle(0.0)
        size = bevel[at]
        return VBASingle((_BEVEL_SIZE if size is None else int(size)) / _EMU_PER_POINT)

    def _set_bevel_size(self, top: bool, at: int, value: object) -> None:
        wanted = _single(value)
        if not 0 <= wanted <= 1584:
            raise error(*_BAD_VALUE)
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

    # The body.

    @member
    def Depth(self) -> object:
        three_d = self.body()
        return VBASingle(36.0 if three_d is None else (three_d.extrusion_h or 0) / _EMU_PER_POINT)

    @setter('Depth')
    def _depth(self, value: object) -> None:
        # Less than nothing is taken as its size; nothing writes no extrusion.
        wanted = abs(_single(value))
        if wanted > 1585.5:
            raise VBAUnsupportedError(f'A depth of {wanted} is not implemented')
        self.change(body=True).extrusion_h = round(wanted * _EMU_PER_POINT) or None

    @member
    def ContourWidth(self) -> object:
        three_d = self.body()
        return VBASingle(0.0 if three_d is None else (three_d.contour_w or 0) / _EMU_PER_POINT)

    @setter('ContourWidth')
    def _contour_width(self, value: object) -> None:
        wanted = _size(value, 1584, 2000, 'A contour')
        self.change(body=True).contour_w = round(wanted * _EMU_PER_POINT) or None

    @member
    def Z(self) -> object:
        three_d = self.body()
        return VBASingle(float(_MIXED) if three_d is None else (three_d.z or 0) / _EMU_PER_POINT)

    @setter('Z')
    def _z(self, value: object) -> None:
        wanted = _single(value)
        if abs(wanted) > 2000:
            raise VBAUnsupportedError(f'A Z of {wanted} is not implemented')
        # Nothing writes no z (tests/fixtures/three_d_round_11.json).
        self.change(body=True).z = round(wanted * _EMU_PER_POINT) or None

    @member
    def ProjectText(self) -> object:
        three_d = self.held()
        return VBAInt(0 if three_d is not None and not three_d.project_text else -1, 'Long')

    @setter('ProjectText')
    def _project_text(self, value: object) -> None:
        # Off is the text body's flatTx (tests/fixtures/three_d_round_4.json); a line, with no text, still reads on.
        wanted = int(to_integer(value, 'Long')) != 0
        three_d = self.change(body=True)
        three_d.project_text = wanted or self.owner.shape.kind == 'line'

    @member
    def PresetMaterial(self) -> object:
        three_d = self.body()
        return VBAInt(-2 if three_d is None else _MATERIALS.index(three_d.material) + 1, 'Long')

    @setter('PresetMaterial')
    def _preset_material(self, value: object) -> None:
        wanted = _choice(value)
        if not 1 <= wanted <= len(_MATERIALS):
            raise VBAUnsupportedError(f'PresetMaterial {wanted} is not implemented')
        three_d = self.old_preset('body', body=True)
        assert three_d is not None
        three_d.material = _MATERIALS[wanted - 1]

    # The colours.

    @member
    def ContourColor(self) -> object:
        self.admit()
        return ThreeDColor(self, 'ContourColor')

    @member
    def ExtrusionColor(self) -> object:
        self.admit()
        return ThreeDColor(self, 'ExtrusionColor')

    def read_colour(self, which: str) -> int:
        three_d = self.body()
        if three_d is None:
            return _COLOURS_UNSET[which]
        rgb, scheme = ((three_d.contour_rgb, three_d.contour_scheme) if which == 'ContourColor'
                       else (three_d.extrusion_rgb, three_d.extrusion_scheme))
        if scheme is not None:
            colours = self.owner.sheet.book.stylesheet.colors.theme
            return bgr(colours[3 + int(scheme[-1])])
        return _COLOURS_UNSET[which] if rgb is None else bgr(rgb)

    def set_colour(self, which: str, rgb: str | None, scheme: str | None) -> None:
        three_d = self.change(body=True)
        if which == 'ContourColor':
            three_d.contour_rgb, three_d.contour_scheme = rgb, scheme
            if three_d.contour_w is None:
                three_d.contour_w = _EMU_PER_POINT
        else:
            three_d.extrusion_rgb, three_d.extrusion_scheme = rgb, scheme
            if three_d.extrusion_h is None:
                three_d.extrusion_h = _COLOURED_DEPTH

    @member
    def ExtrusionColorType(self) -> object:
        three_d = self.body()
        coloured = three_d is not None and (three_d.extrusion_rgb or three_d.extrusion_scheme) is not None
        return VBAInt(2 if coloured else 1, 'Long')

    @setter('ExtrusionColorType')
    def _extrusion_color_type(self, value: object) -> None:
        wanted = _choice(value)
        if wanted == 3:
            raise error(*_BAD_VALUE)
        if wanted not in (1, 2):
            raise VBAUnsupportedError(f'ExtrusionColorType {wanted} is not implemented')
        if wanted == 1:
            # Automatic takes the colour away and leaves the depth it gave; on a plain shape it writes the sp3d alone, as
            # BevelTopType = msoBevelNone does (tests/fixtures/three_d_round_7.json).
            three_d = self.change(scene=False, body=True)
            three_d.extrusion_rgb = three_d.extrusion_scheme = None
        # Custom with no colour yet changes nothing.

    # The camera.

    def _camera(self) -> tuple[str, int, bool, int, tuple[float, float, float] | None] | None:
        three_d = self.scene()
        return None if three_d is None else _CAMERAS[_CAMERA_NUMBERS[three_d.camera]]

    @member
    def PresetCamera(self) -> object:
        three_d = self.scene()
        return VBAInt(-2 if three_d is None else _CAMERA_NUMBERS[three_d.camera], 'Long')

    @method
    def SetPresetCamera(self, PresetCamera: object) -> object:  # noqa: N803 - native named argument
        wanted = _choice(PresetCamera)
        if wanted not in _CAMERAS:
            raise VBAUnsupportedError(f'SetPresetCamera {wanted} is not implemented')
        # In place of what the camera was, its rot and field of view with it.
        three_d = self.change()
        three_d.camera, three_d.rotation, three_d.fov = _CAMERAS[wanted][0], None, None
        return EMPTY

    @member
    def Perspective(self) -> object:
        camera = self._camera()
        return VBAInt(-2 if camera is None else (-1 if camera[2] else 0), 'Long')

    @setter('Perspective')
    def _perspective(self, value: object) -> None:
        number = int(to_integer(value, 'Long'))
        held = self.scene()
        now = 'orthographicFront' if held is None else held.camera
        if number in (-2, 2, 3):
            # 2.5, "3" and a date are the parameter error (tests/fixtures/three_d_round_11.json), as -2 is; 1 is msoFalse
            # (three_d_round_12.json, three_d_round_13.json).
            raise error(*_BAD_VALUE)
        if number == -3 and not (held is not None and _CAMERAS[_CAMERA_NUMBERS[now]][2]):
            # The toggle was measured from a camera with perspective only, where it takes it away.
            raise VBAUnsupportedError('Perspective = msoTriStateToggle without perspective is not implemented')
        if number not in (-3, -1, 0, 1):
            raise VBAUnsupportedError(f'Perspective = {number} is not implemented')
        wanted = number == -1
        # From any old camera, from whichever side, the camera faces front.
        if now != 'orthographicFront' and not now.startswith(('legacyOblique', 'legacyPerspective')):
            raise VBAUnsupportedError(f'Perspective on the {now} camera is not implemented')
        three_d = self.change()
        _drop_hidden_scene(three_d)
        three_d.camera = 'legacyPerspectiveFront' if wanted else 'legacyObliqueFront'
        # Either way the camera's rot is written out.
        three_d.rotation = three_d.rotation or (0, 0, 0)
        if not wanted:
            three_d.fov = None

    @member
    def FieldOfView(self) -> object:
        three_d = self.scene()
        if three_d is None:
            return VBASingle(45.0)
        return VBASingle(three_d.fov / _ANGLE if three_d.fov is not None else float(_CAMERAS[_CAMERA_NUMBERS[three_d.camera]][1]))

    @setter('FieldOfView')
    def _field_of_view(self, value: object) -> None:
        wanted = _size(value, 50, 179, 'A field of view')
        three_d = self.change()
        if _CAMERAS[_CAMERA_NUMBERS[three_d.camera]][2]:
            three_d.fov = round(wanted * _ANGLE)
        # A camera without perspective keeps none.

    @member
    def PresetExtrusionDirection(self) -> object:
        camera = self._camera()
        return VBAInt(-2 if camera is None else camera[3], 'Long')

    @method
    def SetExtrusionDirection(self, PresetExtrusionDirection: object) -> object:  # noqa: N803 - native named argument
        wanted = _choice(PresetExtrusionDirection)
        if wanted not in _SIDES:
            raise VBAUnsupportedError(f'SetExtrusionDirection {wanted} is not implemented')
        three_d = self.old_preset('body')
        assert three_d is not None
        if three_d.rotation is not None:
            raise VBAUnsupportedError('SetExtrusionDirection on a turned camera is not implemented')
        family = 'legacyPerspective' if three_d.camera.startswith('legacyPerspective') else 'legacyOblique'
        three_d.camera = family + _SIDES[wanted]
        return EMPTY

    @member
    def PresetThreeDFormat(self) -> object:
        # Whichever preset what is set is, as a file opened again reads it.
        three_d = self.held()
        if three_d is None or not three_d.scene or not three_d.body:
            return VBAInt(-2, 'Long')
        # A contour beside it still reads the preset (tests/fixtures/three_d_round_5.json); the old default, with its bevels, does
        # not (three_d_round_2.json).
        if any(bevel is not None and bevel[2] != 0 for bevel in (three_d.bevel_top, three_d.bevel_bottom)):
            # A light set after it still reads the preset (tests/fixtures/three_d_round_6.json), and so does a bevel of no
            # depth, as does a rot the camera already reads (three_d_round_11.json).
            return VBAInt(-2, 'Long')
        found = (three_d.camera, _plain_rotation(three_d.camera, three_d.rotation), three_d.extrusion_h, three_d.material)
        return VBAInt(next((number for number, preset in _FORMATS.items()
                            if (preset[0], _plain_rotation(preset[0], preset[1]), *preset[2:]) == found), -2), 'Long')

    @method
    def SetThreeDFormat(self, PresetThreeDFormat: object) -> object:  # noqa: N803 - native named argument
        wanted = _choice(PresetThreeDFormat)
        if wanted not in _FORMATS:
            raise VBAUnsupportedError(f'SetThreeDFormat {wanted} is not implemented')
        # In place of the camera, its rot, the depth and the material; the rest stays.
        three_d = self.change(body=True)
        _drop_hidden_scene(three_d)
        three_d.camera, three_d.rotation, three_d.extrusion_h, three_d.material = _FORMATS[wanted]
        three_d.fov = None
        return EMPTY

    # The light.

    @member
    def PresetLighting(self) -> object:
        three_d = self.scene()
        return VBAInt(-2 if three_d is None else _RIGS.index(three_d.rig) + 1, 'Long')

    @setter('PresetLighting')
    def _preset_lighting(self, value: object) -> None:
        wanted = _choice(value)
        if not 1 <= wanted <= len(_RIGS):
            raise VBAUnsupportedError(f'PresetLighting {wanted} is not implemented')
        held = self.scene()
        if held is not None and held.rig_dir != 't':
            raise VBAUnsupportedError('PresetLighting over a light from another side is not implemented')
        self.change().rig = _RIGS[wanted - 1]

    def _old_rig(self, three_d: ShapeThreeD) -> tuple[str, int]:
        if not three_d.rig.startswith('legacy'):
            raise VBAUnsupportedError(f'The direction and softness of the {three_d.rig} light are not implemented')
        return three_d.rig[6:-1], int(three_d.rig[-1])

    @member
    def PresetLightingDirection(self) -> object:
        three_d = self.scene()
        if three_d is None:
            return VBAInt(-2, 'Long')
        if three_d.rig.startswith('legacy'):
            number = self._old_rig(three_d)[1]
            found = _OLD_DIRECTIONS.get((number, three_d.rig_dir))
            if found is None:
                raise VBAUnsupportedError(f'The direction of the {three_d.rig} light from {three_d.rig_dir} is not implemented')
            side = _DEGREES[found]
            if three_d.light_rev is not None:
                # Turned, numbers 1 and 4 from the top read their angle (tests/fixtures/three_d_round_7.json). Number 3 from
                # the left read 180 turned 270 and 45 turned 45 (three_d_round_8.json): too few to tell a rule by.
                if number not in (1, 4) or three_d.rig_dir != 't':
                    raise VBAUnsupportedError(f'The direction of the {three_d.rig} light turned is not implemented')
                side = 0
        elif three_d.rig_dir in _QUARTERS:
            side = _QUARTERS[three_d.rig_dir] * 90
        else:
            raise VBAUnsupportedError(f'The direction of a light from {three_d.rig_dir} is not implemented')
        if three_d.light_rev is None:
            return VBAInt(found if three_d.rig.startswith('legacy') else _COMPASS[side // 45], 'Long')
        # The side it lights from turned by its angle, clockwise from the top, where that is a side or a corner; an angle
        # past 135 and short of 270 reads none (tests/fixtures/three_d_round_5.json, three_d_round_6.json).
        angle = three_d.light_rev / _ANGLE
        degrees = (side + angle) % 360
        within = angle <= 135 or angle >= 270
        return VBAInt(_COMPASS[int(degrees // 45)] if within and degrees % 45 == 0 else -2, 'Long')

    @setter('PresetLightingDirection')
    def _preset_lighting_direction(self, value: object) -> None:
        wanted = _choice(value)
        if wanted not in _DIRECTED:
            raise VBAUnsupportedError(f'PresetLightingDirection {wanted} is not implemented')
        three_d = self.old_preset('scene')
        if three_d is None:
            return
        if not three_d.rig.startswith('legacy'):
            raise VBAUnsupportedError(f'PresetLightingDirection of the {three_d.rig} light is not implemented')
        if three_d.camera not in ('orthographicFront', 'legacyObliqueTopRight'):
            # Excel writes the old light by where the camera looks from (tests/fixtures/three_d_round_4.json).
            raise VBAUnsupportedError(f'PresetLightingDirection with the {three_d.camera} camera is not implemented')
        kind, now = self._old_rig(three_d)
        number, side = _DIRECTED[wanted]
        if now != 3 and number != now:
            # From another number Excel keeps the number where it can (tests/fixtures/three_d_round_6.json).
            raise VBAUnsupportedError(f'PresetLightingDirection {wanted} from the {three_d.rig} light is not implemented')
        three_d.rig, three_d.rig_dir = f'legacy{kind}{number}', side

    @member
    def PresetLightingSoftness(self) -> object:
        three_d = self.scene()
        if three_d is None or not three_d.rig.startswith('legacy'):
            return VBAInt(-2, 'Long')
        return VBAInt(_SOFTNESS[self._old_rig(three_d)[0]], 'Long')

    @setter('PresetLightingSoftness')
    def _preset_lighting_softness(self, value: object) -> None:
        wanted = _choice(value)
        kinds = {number: kind for kind, number in _SOFTNESS.items()}
        if wanted not in kinds:
            raise VBAUnsupportedError(f'PresetLightingSoftness {wanted} is not implemented')
        three_d = self.old_preset('scene')
        if three_d is None:
            return
        _kind, number = self._old_rig(three_d)
        three_d.rig = f'legacy{kinds[wanted]}{number}'

    @member
    def LightAngle(self) -> object:
        three_d = self.scene()
        return VBASingle(0.0 if three_d is None or three_d.light_rev is None else three_d.light_rev / _ANGLE)

    @setter('LightAngle')
    def _light_angle(self, value: object) -> None:
        self.change().light_rev = _stored(_single(value))

    # The camera's rotation: RotationX turns its lon the other way, RotationY is its lat and RotationZ its rev the other way.

    def _rotation(self) -> tuple[int, int, int] | None:
        """The rot the camera turns from: its own, or the one its preset reads where that is whole degrees; None where the
        preset's is not."""
        three_d = self.scene()
        if three_d is None:
            return (0, 0, 0)
        if three_d.rotation is not None:
            return three_d.rotation
        implied = _CAMERAS[_CAMERA_NUMBERS[three_d.camera]][4]
        if implied is None:
            return (0, 0, 0)
        # A turn writes the preset's rotation out, as Singles, with the one turned.
        return (_stored(_single(implied[1])), _stored(-_single(implied[0])), _stored(-_single(implied[2])))

    def _read_turn(self, at: int, sign: int, implied_at: int) -> object:
        three_d = self.scene()
        if three_d is not None and three_d.rotation is None:
            implied = _CAMERAS[_CAMERA_NUMBERS[three_d.camera]][4]
            if implied is not None:
                return VBASingle(implied[implied_at])
        rotation = self._rotation()
        assert rotation is not None
        return VBASingle(_degrees(rotation[at], sign))

    def _start(self) -> tuple[int, int, int]:
        rotation = self._rotation()
        if rotation is None:
            raise VBAUnsupportedError('Turning a camera whose preset is turned by part of a degree is not implemented')
        return rotation

    def _turned(self, at: int, sign: int) -> float:
        return _degrees(self._start()[at], sign)

    def _turn(self, at: int, angle: int) -> None:
        rotation = list(self._start())
        three_d = self.change()
        rotation[at] = angle
        three_d.rotation = (rotation[0], rotation[1], rotation[2])

    @member
    def RotationX(self) -> object:
        return self._read_turn(1, -1, 0)

    @setter('RotationX')
    def _rotation_x(self, value: object) -> None:
        self._turn(1, _stored(-_single(value)))

    @member
    def RotationY(self) -> object:
        return self._read_turn(0, 1, 1)

    @setter('RotationY')
    def _rotation_y(self, value: object) -> None:
        self._turn(0, _stored(_single(value)))

    @member
    def RotationZ(self) -> object:
        return self._read_turn(2, -1, 2)

    @setter('RotationZ')
    def _rotation_z(self, value: object) -> None:
        self._turn(2, _stored(-_single(value)))

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
        # Excel writes the lon one short some of the time, past nothing or not (-15 from 0, 20 from 60; not 10 or 45 from 0):
        # not known.
        raise VBAUnsupportedError('IncrementRotationHorizontal is not implemented')

    @method
    def IncrementRotationVertical(self, Increment: object) -> object:  # noqa: N803 - native named argument
        # Excel writes the lat one short some of the time, past nothing or not (10 and 20 from 0, -20 from 300; not 45 from 0
        # or 10 from 30): not known.
        raise VBAUnsupportedError('IncrementRotationVertical is not implemented')

    @method
    def ResetRotation(self) -> object:
        # On a plain shape it writes the scene3d, as any change does; the old perspective camera goes back to straight on.
        held = self.scene()
        if held is not None and held.camera not in ('orthographicFront', 'legacyPerspectiveFront'):
            raise VBAUnsupportedError(f'ResetRotation of the {held.camera} camera is not implemented')
        three_d = self.change()
        three_d.rotation = None
        if three_d.camera == 'legacyPerspectiveFront':
            three_d.camera, three_d.fov = 'orthographicFront', None
        return EMPTY


class ThreeDColor(HeldPart):
    """A ThreeD's contour or extrusion colour: its RGB, or a theme colour by ObjectThemeColor."""

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
        self.owner.set_colour(self.which, f'{red:02X}{green:02X}{blue:02X}', None)

    def vba_get(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        if name.lower() != 'rgb':
            raise VBAUnsupportedError(f'ColorFormat.{name} of a ThreeD is not implemented')
        return super().vba_get(name, args, named)

    def vba_set(self, name: str, value: object, args: Sequence[object] = (), named: dict[str, object] | None = None,
                *, by_ref: bool = False) -> None:
        if name.lower() == 'objectthemecolor':
            wanted = int(to_integer(value, 'Long'))
            if wanted not in _ACCENTS:
                raise VBAUnsupportedError(f'ObjectThemeColor {wanted} of a ThreeD is not implemented')
            self.owner.set_colour(self.which, None, _ACCENTS[wanted])
            return
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
            return RangeThreeDColor(self, name)
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


class RangeThreeDColor(Created):
    """A range's contour or extrusion colour: what its shapes share; shapes of different colours are not measured."""

    vba_type_name = 'ColorFormat'

    def __init__(self, effect: RangeThreeD, name: str) -> None:
        self.effect, self.name = effect, name

    def colours(self) -> list[ThreeDColor]:
        return [ThreeDColor(part, self.name) for part in self.effect.parts()]

    def vba_get(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        found = [one.vba_get(name, args, named) for one in self.colours()]
        if any(one != found[0] for one in found[1:]):
            raise VBAUnsupportedError(f'ColorFormat.{name} of a range whose shapes differ is not implemented')
        return found[0]

    def vba_set(self, name: str, value: object, args: Sequence[object] = (), named: dict[str, object] | None = None,
                *, by_ref: bool = False) -> None:
        for one in self.colours():
            one.vba_set(name, value, args, named, by_ref=by_ref)
