"""Worksheet.Buttons, CheckBoxes, OptionButtons, DropDowns and ListBoxes: the old names of the sheet's form controls, which the
macro recorder writes (scripts/measure_old_controls.py, tests/fixtures/old_controls.json).

Each is a view of the form controls on the sheet, not a model of its own: an item answers through the shape and its
ControlFormat, so what a macro sets one way reads back the other. A collection reads a member of its first item and sets it on
all of them. What only reads here, the alignment of a button's caption among them, is what Excel answers for a new control;
setting those is refused until what they write is measured.
"""
from __future__ import annotations

from collections.abc import Sequence
from typing import TYPE_CHECKING, Any, ClassVar, cast

from pyopenvba.apps.excel._creator import Created
from pyopenvba.apps.excel._held import HeldPart, shape_here
from pyopenvba.apps.excel._shapes import ShapeObject
from pyopenvba.shapes._xlsx import on_pixel
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import MemberSpec, VBACollection, VBAObject, member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, NOTHING, NULL, VBAArray, VBAErrorValue, VBAInt, VBASingle, error, to_bool, to_number, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet
    from pyopenvba.shapes._values import Shape

_NOT_THERE = (1004, 'Unable to get the property')
_NO_MEMBER = (438, "Object doesn't support this property or method")
#: The items' types by what a shape's control calls itself, and AddFormControl's Type for each.
_TYPES = {'Button': 0, 'CheckBox': 1, 'Radio': 7, 'Drop': 2, 'List': 6}


def _object(value: object) -> VBAObject:
    return cast(VBAObject, value)


class OldObject(HeldPart):
    """A shape under its old name, a form control's or a drawing's (see _old_drawings). Held after the shape has gone, it answers
    only what needs no shape (tests/fixtures/shape_held_deleted.json, tests/fixtures/old_drawings.json)."""

    vba_type_name = ''
    kind = ''
    held_free = frozenset({'parent', 'application', 'creator', 'shaperange', 'addindent', 'interior', 'border'})

    def __init__(self, sheet: Worksheet, shape: Shape) -> None:
        self.sheet, self.shape = sheet, shape
        self.drawn = ShapeObject(sheet, shape)

    def held_gone(self) -> bool:
        return not shape_here(self.sheet, self.shape)

    def held_error(self, name: str) -> int:
        return 1004

    def holds(self, shape: Shape) -> bool:
        """Whether a shape is one its own collection counts."""
        return _of_kind(shape, self.kind)

    @member
    def Parent(self) -> object:
        return self.sheet

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @member
    def Visible(self) -> object:
        return bool(to_number(self.drawn.vba_get('Visible', [])))

    @setter('Visible')
    def _visible(self, value: object) -> None:
        self.drawn.vba_set('Visible', -1 if to_bool(value) else 0)

    @member
    def Locked(self) -> object:
        return bool(self.drawn.vba_get('Locked', []))

    @setter('Locked')
    def _locked(self, value: object) -> None:
        self.drawn.vba_set('Locked', to_bool(value))

    @member
    def Index(self) -> object:
        """Its place among the controls of its kind, in the order they lie on the sheet."""
        return VBAInt(1 + [one for one in self.sheet.shapes_ if self.holds(one)].index(self.shape), 'Long')

    @member
    def ZOrder(self) -> object:
        return self.drawn.vba_get('ZOrderPosition', [])

    @method
    def BringToFront(self) -> object:
        self.drawn.vba_call('ZOrder', [VBAInt(0, 'Long')])
        return True

    @method
    def SendToBack(self) -> object:
        self.drawn.vba_call('ZOrder', [VBAInt(1, 'Long')])
        return True

    @member
    def ShapeRange(self) -> object:
        return _object(self.sheet.vba_get('Shapes', [])).vba_get('Range', [self.drawn.vba_get('Name', [])])

    @method
    def Select(self, Replace: object = MISSING) -> object:  # noqa: N803 - native named argument
        return self.drawn.vba_call('Select', [] if Replace is MISSING else [Replace])

    @method
    def Delete(self) -> object:
        return self.drawn.vba_call('Delete', [])

    @method
    def Duplicate(self) -> object:
        found = self.drawn.vba_call('Duplicate', [])
        return item_of(self.sheet, getattr(found, 'shape'))


class OldControl(OldObject):
    """A form control under its old name."""

    vba_type_name = ''

    def __init__(self, sheet: Worksheet, shape: Shape) -> None:
        super().__init__(sheet, shape)
        self.control = _object(self.drawn.vba_get('ControlFormat', []))


def _of_kind(shape: Shape, kind: str) -> bool:
    return shape.kind == 'formControl' and shape.control is not None and shape.control.kind == kind


def _through(owner: str, name: str, *, writes: bool = True) -> MemberSpec:
    """A member of the shape or of its ControlFormat, under the same name."""

    def get(item: OldObject) -> object:
        return getattr(item, owner).vba_get(name, [])

    def put(item: OldObject, value: object) -> None:
        getattr(item, owner).vba_set(name, value)

    return MemberSpec(name=name, kind='property', getter=get, setter=put if writes else None)


def _box_side(name: str) -> MemberSpec:
    """Left, Top, Width or Height, a Double. A new form control's is the pixel its file will hold, whatever the Shape reads. A width set here stays as it was given on the
    Shape (tests/fixtures/new_control_old_places.json)."""

    def get(item: OldObject) -> object:
        shape = item.shape
        found = float(VBASingle(shape.top)) if name == 'Top' else float(cast(float, item.drawn.vba_get(name, [])))
        if shape.kind != 'formControl' or shape.source:
            return found
        return on_pixel(found) if name in ('Left', 'Top') else round(found * 4 / 3) * 0.75

    def put(item: OldObject, value: object) -> None:
        shape = item.shape
        if name == 'Width' and shape.kind == 'formControl' and not shape.source and shape.control is not None and shape.control.kind != 'Scroll':
            shape.width = float(VBASingle(float(to_number(value))))
            item.sheet.drawing_changed()
            return
        item.drawn.vba_set(name, value)
        if name == 'Height' and shape.kind == 'formControl' and not shape.source:
            shape.unsure_top = True

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


def _new_control(name: str, value: object) -> MemberSpec:
    """What a new control answers; setting it is refused until what it writes is measured."""

    def get(item: OldControl) -> object:
        del item
        return value

    def put(item: OldControl, wanted: object) -> None:
        del item, wanted
        raise VBAUnsupportedError(f'Setting {name} on a form control is not implemented')

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


def _taken_on_new(name: str, value: object) -> MemberSpec:
    """A setting a new control takes and keeps nothing of, reading ``value`` before and after and writing nothing
    (tests/fixtures/new_control_edits.json); on a control loaded from a file what a set writes is not measured, and it is refused."""

    def get(item: OldControl) -> object:
        del item
        return value

    def put(item: OldControl, wanted: object) -> None:
        if item.shape.control is None or item.shape.source:
            raise VBAUnsupportedError(f'Setting {name} on a form control in a file is not implemented')
        to_bool(wanted)

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


def _taken_on_loaded(name: str, value: object, values: tuple[int, ...] | None = None) -> MemberSpec:
    """A setting a control takes and keeps nothing of, reading ``value`` before and after and writing nothing, as Excel's does on a
    control loaded from a file and on a new one (tests/fixtures/button_text.json, new_control_edits.json)."""

    def get(item: OldControl) -> object:
        del item
        return value

    def put(item: OldControl, wanted: object) -> None:
        if item.shape.control is None:
            raise VBAUnsupportedError(f'Setting {name} on this form control is not implemented')
        if values is None:
            to_text(wanted)
        elif int(to_number(wanted)) not in values:
            raise VBAUnsupportedError(f'Setting {name} to {wanted} on a form control is not implemented')

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


def _raising(name: str, failure: tuple[int, str]) -> MemberSpec:
    def get(item: OldControl) -> object:
        del item
        raise error(*failure)

    def put(item: OldControl, wanted: object) -> None:
        del item, wanted
        raise error(*failure)

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


for _spec in (*(_box_side(_name) for _name in ('Left', 'Top', 'Width', 'Height')), *(_through('drawn', _name) for _name in ('Name', 'Placement', 'OnAction')), _through('drawn', 'TopLeftCell', writes=False), _through('drawn', 'BottomRightCell', writes=False)):
    OldObject.vba_add_member(_spec)
for _spec in (_through('control', 'Enabled'), _through('control', 'PrintObject')):
    OldControl.vba_add_member(_spec)


class _Captioned(OldControl):
    """A control with a caption: a button, a check box or an option button."""

    vba_type_name = ''

    def _text_of(self) -> VBAObject:
        return _object(_object(self.drawn.vba_get('TextFrame', [])).vba_call('Characters', []))

    def _caption(self) -> str:
        return to_text(self._text_of().vba_get('Text', []))

    def _set_caption(self, value: object) -> None:
        from pyopenvba.apps.excel._protection import enforced

        if enforced(self.sheet) is not None:
            raise error(1004, 'Unable to set the Caption property')
        self._text_of().vba_set('Text', to_text(value))

    @member
    def Caption(self) -> object:
        return self._caption()

    @setter('Caption')
    def _caption_set(self, value: object) -> None:
        self._set_caption(value)

    @member
    def Text(self) -> object:
        return self._caption()

    @setter('Text')
    def _text_set(self, value: object) -> None:
        self._set_caption(value)

    @method
    def Characters(self, Start: object = MISSING, Length: object = MISSING) -> object:  # noqa: N803 - native named arguments
        return ControlCharacters(self, Start, Length)

    @member
    def Accelerator(self) -> object:
        return ''

    @setter('Accelerator')
    def _accelerator(self, value: object) -> None:
        # It takes one and goes on answering none.
        to_text(value)


for _spec in (_through('control', 'LockedText'), _taken_on_loaded('PhoneticAccelerator', '')):
    _Captioned.vba_add_member(_spec)


class ControlCharacters(Created):
    """Some of a caption. A check box's or an option button's take a start and no length."""

    vba_type_name = 'Characters'

    def __init__(self, item: _Captioned, start: object, length: object) -> None:
        self.item = item
        self.start = 0 if start is MISSING else max(0, int(to_number(start)) - 1)
        self.length = None if length is MISSING else int(to_number(length))

    def _span(self, writing: bool = False) -> tuple[int, int]:
        text = self.item._caption()  # pyright: ignore[reportPrivateUsage]
        length = self.length if writing or isinstance(self.item, Button) else None
        return self.start, len(text) if length is None else min(len(text), self.start + length)

    @member
    def Text(self) -> object:
        begin, end = self._span()
        return self.item._caption()[begin:end]  # pyright: ignore[reportPrivateUsage]

    @setter('Text')
    def _text(self, value: object) -> None:
        begin, end = self._span(writing=True)
        whole = self.item._caption()  # pyright: ignore[reportPrivateUsage]
        self.item._set_caption(whole[:begin] + to_text(value) + whole[end:])  # pyright: ignore[reportPrivateUsage]

    @member
    def Count(self) -> object:
        begin, end = self._span()
        return VBAInt(end - begin, 'Long')

    @member
    def Font(self) -> object:
        if isinstance(self.item, Button):
            return self.item.vba_get('Font', [])
        return UnreadFont()


class UnreadFont(Created):
    """The font of a check box's or an option button's caption, whose every member reads Null."""

    vba_type_name = 'Font'

    @member
    def Bold(self) -> object:
        return NULL


class Button(_Captioned):
    vba_type_name = 'Button'
    kind = 'Button'

    @member
    def Font(self) -> object:
        return self._text_of().vba_get('Font', [])


#: A button's flags and the VML element each is written as.
_BUTTON_FLAGS = {'AddIndent': 'JustLastX', 'DefaultButton': 'Default', 'CancelButton': 'Cancel', 'DismissButton': 'Dismiss', 'HelpButton': 'Help'}


def _button_flag(name: str) -> MemberSpec:
    """One of a button's flags. Set on a button loaded from a file or a new one, it reads True for the session but AddIndent, which reads
    False, and its VML gets it as Excel's does; Excel reads none back from a file (tests/fixtures/loaded_control_edits.json,
    button_flags.json, new_control_edits.json)."""
    element = _BUTTON_FLAGS[name]

    def get(item: OldControl) -> object:
        control = item.shape.control
        return name != 'AddIndent' and control is not None and element in control.button_flags

    def put(item: OldControl, wanted: object) -> None:
        from pyopenvba.apps.excel._protection import guard_drawing

        control = item.shape.control
        if control is None:
            raise VBAUnsupportedError(f'Setting {name} on this form control is not implemented')
        guard_drawing(item.sheet, 'ControlFormat')
        flags = set(control.button_flags)
        if to_bool(wanted):
            flags.add(element)
            if element in _SOLE_ROLES:
                # One button on the sheet holds each of these; the other roles stay (tests/fixtures/button_roles.json).
                for other in item.sheet.shapes_:
                    held = other.control
                    if other is not item.shape and held is not None and element in held.button_flags:
                        held.button_flags = held.button_flags - {element}
        else:
            flags.discard(element)
        control.button_flags = frozenset(flags)
        item.sheet.drawing_changed()

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


#: The roles a sheet gives one button at a time.
_SOLE_ROLES = frozenset({'Default', 'Cancel', 'Help'})


def _text_frame(name: str) -> MemberSpec:
    """A button's HorizontalAlignment, VerticalAlignment or Orientation: what its DrawingML holds, or what it was set to this session. Set
    on a new button, or on one loaded from a file whose text is laid out as the one measured, it is written as Excel writes it, one of
    the three at a time (tests/fixtures/button_text.json, new_control_edits.json); anything else is refused, as what it writes is not
    measured."""
    from pyopenvba.shapes._xlsx import TEXT_FRAME_VALUES, control_text_frame, text_frame_measured

    def get(item: OldControl) -> object:
        control = item.shape.control
        if control is not None and control.text_frame is not None and control.text_frame[0] == name:
            return VBAInt(control.text_frame[1], 'Long')
        return VBAInt(control_text_frame(item.shape.source)[name], 'Long')

    def put(item: OldControl, wanted: object) -> None:
        from pyopenvba.apps.excel._protection import guard_drawing

        control = item.shape.control
        if control is None:
            raise VBAUnsupportedError(f'Setting {name} on this form control is not implemented')
        value = int(to_number(wanted))
        if value not in TEXT_FRAME_VALUES[name]:
            raise VBAUnsupportedError(f'Setting {name} to {value} on a form control is not implemented')
        # A new button is laid out as the one measured (tests/fixtures/new_control_edits.json).
        if control.text_frame is not None and control.text_frame[0] != name or item.shape.source and not text_frame_measured(item.shape.source):
            raise VBAUnsupportedError(f'Setting {name} on a button laid out otherwise than the one measured is not implemented')
        guard_drawing(item.sheet, 'ControlFormat')
        control.text_frame = (name, value)
        item.sheet.drawing_changed()

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


for _spec in (_new_control('AutoSize', False), _taken_on_new('AutoScaleFont', False), *(_button_flag(_name) for _name in _BUTTON_FLAGS),
              *(_text_frame(_name) for _name in ('HorizontalAlignment', 'VerticalAlignment', 'Orientation')),
              _taken_on_loaded('ReadingOrder', VBAInt(27, 'Long'), (-5004, -5003, -5002)), _new_control('Formula', '')):
    Button.vba_add_member(_spec)


class _Valued(OldControl):
    """A control that holds a value and may be linked to a cell."""

    vba_type_name = ''

    @member
    def Value(self) -> object:
        """On or mixed is a Double, and off xlOff, a Long (tests/fixtures/old_collection_members.json)."""
        found = self.control.vba_get('Value', [])
        number = to_number(found)
        return found if number == -4146 else float(number)

    @setter('Value')
    def _value(self, value: object) -> None:
        self.control.vba_set('Value', value)


def _three_d() -> MemberSpec:
    """Display3DShading as the control's part holds it. Set on a control loaded from a file or a new one, it is written where Excel writes
    it (tests/fixtures/loaded_control_edits.json, new_control_edits.json)."""

    def get(item: OldControl) -> object:
        control = item.shape.control
        return bool(control.three_d) if control is not None else False

    def put(item: OldControl, wanted: object) -> None:
        from pyopenvba.apps.excel._protection import guard_drawing

        control = item.shape.control
        if control is None:
            raise VBAUnsupportedError('Setting Display3DShading on this form control is not implemented')
        guard_drawing(item.sheet, 'ControlFormat')
        control.three_d = to_bool(wanted)
        item.sheet.drawing_changed()

    return MemberSpec(name='Display3DShading', kind='property', getter=get, setter=put)


for _spec in (_through('control', 'LinkedCell'), _three_d()):
    _Valued.vba_add_member(_spec)


class CheckBox(_Valued, _Captioned):
    vba_type_name = 'CheckBox'
    kind = 'CheckBox'


class OptionButton(_Valued, _Captioned):
    vba_type_name = 'OptionButton'
    kind = 'Radio'

    @member
    def GroupBox(self) -> object:
        if any(one.kind == 'formControl' and one.control is not None and one.control.kind == 'GBox' for one in self.sheet.shapes_):
            raise VBAUnsupportedError('OptionButton.GroupBox on a sheet with group boxes is not implemented')
        return NOTHING


class _Listed(_Valued):
    """A control holding a list."""

    vba_type_name = ''

    @member
    def List(self, Index: object = MISSING) -> object:  # noqa: N803 - native named argument
        found = self.control.vba_get('List', [] if Index is MISSING else [Index])
        return NULL if Index is MISSING and not int(to_number(self.control.vba_get('ListCount', []))) else found

    @setter('List')
    def _list(self, value: object) -> None:
        self.control.vba_set('List', value)

    @method
    def AddItem(self, Text: object = MISSING, Index: object = MISSING) -> object:  # noqa: N803 - native named arguments
        return self.control.vba_call('AddItem', [Text] if Index is MISSING else [Text, Index])

    @method
    def RemoveItem(self, Index: object = MISSING, Count: object = MISSING) -> object:  # noqa: N803 - native named arguments
        return self.control.vba_call('RemoveItem', [Index] if Count is MISSING else [Index, Count])

    @method
    def RemoveAllItems(self) -> object:
        return self.control.vba_call('RemoveAllItems', [])

    @member
    def Value(self) -> object:
        return self.control.vba_get('Value', [])

    @setter('Value')
    def _value(self, value: object) -> None:
        if not int(to_number(self.control.vba_get('ListCount', []))):
            raise error(1004, 'Unable to set the Value property')
        self.control.vba_set('Value', value)

    @member
    def Selected(self, Index: object = MISSING) -> object:  # noqa: N803 - native named argument
        chosen = int(to_number(self.control.vba_get('ListIndex', [])))
        if Index is MISSING:
            count = int(to_number(self.control.vba_get('ListCount', [])))
            if not count or isinstance(self, DropDown):
                raise error(*_NOT_THERE)
            return VBAArray([(1, count)], items=[chosen == one for one in range(1, count + 1)])
        return chosen == int(to_number(Index))

    @setter('Selected')
    def _selected(self, value: object, Index: object = MISSING) -> None:  # noqa: N803 - native named argument
        if Index is MISSING:
            raise error(*_NOT_THERE)
        if to_bool(value):
            self.control.vba_set('ListIndex', Index)


for _spec in (_through('control', 'ListFillRange'), _through('control', 'ListCount', writes=False), _through('control', 'ListIndex'),
              _new_control('LinkedObject', '')):
    _Listed.vba_add_member(_spec)


class DropDown(_Listed):
    vba_type_name = 'DropDown'
    kind = 'Drop'

    @method
    def Characters(self, Start: object = MISSING, Length: object = MISSING) -> object:  # noqa: N803 - native named arguments
        del Start, Length
        return DropCharacters()

    def _no_text(self) -> str:
        raise error(*_NOT_THERE)

    @member
    def Caption(self) -> object:
        return self._no_text()

    @setter('Caption')
    def _caption_set(self, value: object) -> None:
        # It takes one, and nothing can read it back.
        to_text(value)

    @member
    def Text(self) -> object:
        return self._no_text()

    @setter('Text')
    def _text_set(self, value: object) -> None:
        to_text(value)


class DropCharacters(Created):
    """A drop-down's Characters, which is there and cannot be read."""

    vba_type_name = 'Characters'

    @member
    def Count(self) -> object:
        raise error(*_NOT_THERE)

    @member
    def Text(self) -> object:
        return VBAErrorValue(2042)

    @setter('Text')
    def _text(self, value: object) -> None:
        # It takes it, and nothing can read it back.
        to_text(value)

    @member
    def Font(self) -> object:
        return UnreadFont()


DropDown.vba_add_member(_through('control', 'DropDownLines'))


class ListBox(_Listed):
    vba_type_name = 'ListBox'
    kind = 'List'


for _spec in (_through('control', 'MultiSelect'), _raising('Caption', _NO_MEMBER), _raising('Text', _NO_MEMBER)):
    ListBox.vba_add_member(_spec)


_ITEMS: dict[str, type[OldControl]] = {'Button': Button, 'CheckBox': CheckBox, 'Radio': OptionButton, 'Drop': DropDown, 'List': ListBox}


def item_of(sheet: Worksheet, shape: Shape) -> OldObject:
    """The old object a shape answers as."""
    if shape.kind != 'formControl':
        from pyopenvba.apps.excel._old_drawings import drawing_of

        return drawing_of(sheet, shape)
    kind = '' if shape.control is None else shape.control.kind
    if kind not in _ITEMS:
        raise VBAUnsupportedError(f'The old object of a {kind or "shape"} control is not implemented')
    return _ITEMS[kind](sheet, shape)


class OLEFormat(Created):
    """A form control's OLEFormat, whose Object is the old object."""

    vba_type_name = 'OLEFormat'

    def __init__(self, sheet: Worksheet, shape: Shape) -> None:
        self.sheet, self.shape = sheet, shape

    @member
    def Object(self) -> object:
        return item_of(self.sheet, self.shape)

    @member
    def Parent(self) -> object:
        return ShapeObject(self.sheet, self.shape)


class SpreadPart(Created):
    """A collection's own Font, Interior, Border or Characters, answering for all its items. Changing one is refused until what it writes is measured."""

    vba_type_name = ''

    def __init__(self, owner: OldControls, part: str) -> None:
        self.owner, self.part = owner, part
        self.vba_type_name = part

    def vba_get(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        if name.lower() == 'creator':
            return super().vba_get(name, args, named)
        if name.lower() == 'parent':
            return self.owner
        if name.lower() == 'application':
            return self.owner.sheet.book.application
        return self.owner.read_part(self.part, name)

    def vba_set(self, name: str, value: object, args: Sequence[object] = (), named: dict[str, object] | None = None, *, by_ref: bool = False) -> None:
        raise VBAUnsupportedError(f'Setting {self.part}.{name} on {self.owner.vba_type_name} is not implemented')


def _spread_getter(name: str) -> Any:
    def read(self: OldControls) -> object:
        return self.vba_get(name)

    return read


def _spread_setter(name: str) -> Any:
    def write(self: OldControls, value: object) -> None:
        self.vba_set(name, value)

    return write


class OldControls(VBACollection, Created):
    """One kind of a sheet's form controls, or some of them picked by an Array. Asked for a member of its items, a collection of one
    answers for the item, and of several fails, but for what these say (tests/fixtures/old_collection_members.json)."""

    vba_type_name = ''
    kind: ClassVar[str] = ''
    #: What it gives as an object of its own.
    parts: ClassVar[frozenset[str]] = frozenset()
    #: Whether one item alone still answers through that object rather than its own.
    spread_alone: ClassVar[bool] = True
    #: What several items answer.
    several: ClassVar[dict[str, object]] = {'visible': True}
    #: What the collection of one gives as a Double, where its item gives a Long.
    doubled: ClassVar[frozenset[str]] = frozenset({'zorder', 'value', 'listindex', 'dropdownlines', 'listcount'})
    #: Its items' members it answers for them as Excel does, as is, alike and different, typed and late, on a saved and a made workbook
    #: (tests/fixtures/old_collection_reads.json). Each is registered, so the class lists it, and read as the rest are.
    spread_reads: ClassVar[frozenset[str]] = frozenset()
    #: Those of them a set on the whole collection answers for as Excel does, typed and late, in the session and after a save and a reopen
    #: (tests/fixtures/old_collection_lets.json): registered with a setter, which goes through ``vba_set`` as every set here does.
    spread_sets: ClassVar[frozenset[str]] = frozenset()
    #: Members a set on the whole collection takes and loses, its items reading as before, in the session and after a save.
    lost_sets: ClassVar[frozenset[str]] = frozenset()

    def __init_subclass__(cls, **kwargs: Any) -> None:
        super().__init_subclass__(**kwargs)
        assert cls.spread_sets <= cls.spread_reads, cls.spread_sets - cls.spread_reads
        for name in cls.spread_reads:
            cls.vba_add_member(MemberSpec(name=name, kind='property', getter=_spread_getter(name),
                                          setter=_spread_setter(name) if name in cls.spread_sets else None))

    def __init__(self, sheet: Worksheet, picked: list[Shape] | None = None) -> None:
        self.sheet, self.picked = sheet, picked

    def holds(self, shape: Shape) -> bool:
        return _of_kind(shape, self.kind)

    def item_for(self, shape: Shape) -> object:
        return item_of(self.sheet, shape)

    def _shapes(self) -> list[Shape]:
        return self.picked if self.picked is not None else [one for one in self.sheet.shapes_ if self.holds(one)]

    def vba_items(self) -> list[object]:
        return [self.item_for(one) for one in self._shapes()]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, VBAArray):
            shapes = [getattr(self.vba_lookup(one, items), 'shape') for one in index.items]
            return type(self)(self.sheet, shapes)
        if isinstance(index, str):
            for item in items:
                if to_text(getattr(item, 'shape').name).lower() == index.lower():
                    return item
            raise error(*_NOT_THERE)
        position = int(to_number(index))
        if not 1 <= position <= len(items):
            raise error(*_NOT_THERE)
        return items[position - 1]

    @member
    def Parent(self) -> object:
        return self.sheet

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @method
    def Add(self, Left: object = MISSING, Top: object = MISSING, Width: object = MISSING, Height: object = MISSING) -> object:  # noqa: N803 - native named arguments
        from pyopenvba.apps.excel._protection import enforced

        if MISSING in (Left, Top, Width, Height):
            raise error(449, 'Argument not optional')
        if any(isinstance(one, str) for one in (Left, Top, Width, Height)) or enforced(self.sheet) is not None:
            raise error(*_NOT_THERE)
        wanted = [max(0.0, float(to_number(one))) for one in (Left, Top, Width, Height)]
        if self.kind in ('CheckBox', 'Radio', 'List') and not wanted[2] and not wanted[3]:
            raise VBAUnsupportedError(f'{self.vba_type_name}.Add with no size, which takes a size of its own, is not implemented')
        found = _object(self.sheet.vba_get('Shapes', [])).vba_call('AddFormControl', [VBAInt(_TYPES[self.kind], 'Long'), *wanted])
        shape = getattr(found, 'shape')
        if self.kind in ('Button', 'CheckBox', 'Radio'):
            shape.text = shape.name
        if self.kind != 'Button':
            ShapeObject(self.sheet, shape).vba_set('Placement', VBAInt(2, 'Long'))
        return item_of(self.sheet, shape)

    @method
    def Delete(self) -> object:
        for shape in list(self._shapes()):
            ShapeObject(self.sheet, shape).vba_call('Delete', [])
        return EMPTY

    @method
    def Select(self, Replace: object = MISSING) -> object:  # noqa: N803 - native named argument
        for index, shape in enumerate(self._shapes()):
            ShapeObject(self.sheet, shape).vba_call('Select', [False] if index else ([] if Replace is MISSING else [Replace]))
        return True

    @method
    def Group(self) -> object:
        if len(self._shapes()) < 2:
            raise error(*_NOT_THERE)
        raise VBAUnsupportedError(f'Grouping {self.vba_type_name} is not implemented')

    @method
    def BringToFront(self) -> object:
        """Every item to the front, one at a time in order, so they keep their order among themselves on top of the rest
        (tests/fixtures/old_collection_methods.json)."""
        for shape in list(self._shapes()):
            ShapeObject(self.sheet, shape).vba_call('ZOrder', [VBAInt(0, 'Long')])
        return True

    @method
    def SendToBack(self) -> object:
        """Every item to the back, keeping their order among themselves beneath the rest (tests/fixtures/old_collection_methods.json)."""
        for shape in reversed(self._shapes()):
            ShapeObject(self.sheet, shape).vba_call('ZOrder', [VBAInt(1, 'Long')])
        return True

    @method
    def Duplicate(self) -> object:
        """A copy of every item, answered as a collection of the same kind holding the copies (tests/fixtures/old_collection_methods.json)."""
        copies = [getattr(ShapeObject(self.sheet, shape).vba_call('Duplicate', []), 'shape') for shape in list(self._shapes())]
        return type(self)(self.sheet, copies)

    def read_part(self, part: str, name: str) -> object:
        """A member read through what ``part`` gives."""
        if part == 'Interior' and name.lower() in ('colorindex', 'pattern'):
            return VBAInt(-4142, 'Long')
        raise error(*_NOT_THERE)

    def nothing_held(self, name: str) -> object:
        raise VBAUnsupportedError(f'{self.vba_type_name}.{name} with nothing in the collection is not implemented')

    def vba_get(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        """A member of the collection, or else what its items answer."""
        lower = name.lower()
        if lower in {key.lower() for key in self._vba_members} and lower not in {key.lower() for key in self.spread_reads}:
            return super().vba_get(name, args, named)
        return self.items_read(name, args, named)

    def items_read(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        """What the collection answers for its items' member ``name``."""
        lower = name.lower()
        items = self.vba_items()
        if not items:
            return self.nothing_held(name)
        if lower == 'shaperange':
            names = [getattr(one, 'shape').name for one in items]
            found = _object(_object(self.sheet.vba_get('Shapes', [])).vba_get('Range', [VBAArray([(1, len(names))], items=names)]))
            return found.vba_get('Item', args) if args else found
        part = next((one for one in self.parts if one.lower() == lower), None)
        if part is not None and (self.spread_alone or len(items) > 1):
            return SpreadPart(self, part)
        if lower == 'list' and not args:
            raise error(*_NOT_THERE)
        if len(items) == 1:
            if lower == 'zorder' and self.doubled:
                # Among the sheet's form controls only.
                return float(1 + [one for one in self.sheet.shapes_ if one.kind == 'formControl'].index(getattr(items[0], 'shape')))
            found = getattr(items[0], 'vba_get')(name, args, named)
            return float(found) if lower in self.doubled and isinstance(found, int) and found > 0 else found
        if getattr(items[0], 'vba_member')(name) is None:
            return getattr(items[0], 'vba_get')(name, args, named)
        if lower not in self.several:
            raise error(*_NOT_THERE)
        if lower == 'visible':
            # The first item's, whatever the rest hold: all hidden reads False, the second alone hidden True (tests/fixtures/old_collection_reads.json).
            return not getattr(items[0], 'shape').hidden
        return self.several[lower]

    def vba_set(self, name: str, value: object, args: Sequence[object] = (), named: dict[str, object] | None = None, *, by_ref: bool = False) -> None:
        """A member set on every item, though not a name, which they cannot share, nor a whole List, which is 1004 before any item changes, as
        its read is. A list with nothing in it takes no Value or ListIndex, and the others still do. What ``lost_sets`` names is taken and
        changes nothing (tests/fixtures/old_collection_lets.json)."""
        lower = name.lower()
        if lower == 'name':
            raise error(438, "Object doesn't support this property or method")
        items = self.vba_items()
        if not items:
            raise VBAUnsupportedError(f'{self.vba_type_name}.{name} with nothing in the collection is not implemented')
        if lower == 'list' and not args:
            raise error(*_NOT_THERE)
        if lower in self.lost_sets:
            return
        for item in items:
            if not (lower in ('value', 'listindex') and isinstance(item, _Listed) and not int(to_number(item.control.vba_get('ListCount', [])))):
                getattr(item, 'vba_set')(name, value, args, named, by_ref=by_ref)


class Buttons(OldControls):
    vba_type_name = 'Buttons'
    kind = 'Button'
    spread_reads = frozenset({'Accelerator', 'AddIndent', 'CancelButton', 'Caption', 'Characters', 'DefaultButton', 'DismissButton', 'Enabled', 'Font', 'HelpButton', 'HorizontalAlignment', 'Locked', 'LockedText', 'Orientation', 'PhoneticAccelerator', 'Placement', 'PrintObject', 'ReadingOrder', 'ShapeRange', 'Text', 'VerticalAlignment', 'Visible', 'ZOrder'})
    spread_sets = frozenset({'Accelerator', 'AddIndent', 'CancelButton', 'Caption', 'DefaultButton', 'DismissButton', 'Enabled', 'HelpButton', 'HorizontalAlignment', 'Locked', 'LockedText', 'Orientation', 'PhoneticAccelerator', 'Placement', 'PrintObject', 'ReadingOrder', 'Text', 'VerticalAlignment', 'Visible'})
    parts = frozenset({'Characters', 'Font'})
    several = {'visible': True, 'autoscalefont': False, 'readingorder': VBAInt(27, 'Long')}
    lost_sets = frozenset({'autoscalefont'})


class ValueDefault(OldControls):
    """Those whose hidden _Default the type library gives As Long: their items' Value, read and set as Value is over them, typed. Late bound
    Excel answers 438 for it, as for every hidden member; the model follows the typed answers (tests/fixtures/old_collection_defaults.json)."""

    vba_type_name = ''

    @member(name='_Default')
    def value_default(self) -> object:
        return self.vba_get('Value')

    @setter('_Default')
    def _value_default(self, value: object) -> None:
        self.vba_set('Value', value)

    def vba_set(self, name: str, value: object, args: Sequence[object] = (), named: dict[str, object] | None = None, *, by_ref: bool = False) -> None:
        if name.lower() != '_default':
            super().vba_set(name, value, args, named, by_ref=by_ref)
            return
        if self.vba_type_name == 'DrawingObjects':
            # Excel sets every control's Value but leaves the sheet's two option buttons both off, by a rule not worked out.
            raise VBAUnsupportedError('DrawingObjects._Default set is not implemented')
        for item in self.vba_items():
            if getattr(item, 'vba_member')('Value') is not None:
                getattr(item, 'vba_set')('Value', value)

    def vba_value(self) -> object:
        # Read where a value is wanted, the collection is its _Default, typed and late alike.
        return self.vba_get('Value')


class CheckBoxes(ValueDefault):
    vba_type_name = 'CheckBoxes'
    kind = 'CheckBox'
    spread_reads = frozenset({'Accelerator', 'Border', 'Caption', 'Characters', 'Display3DShading', 'Enabled', 'Interior', 'LinkedCell', 'Locked', 'LockedText', 'PhoneticAccelerator', 'Placement', 'PrintObject', 'ShapeRange', 'Text', 'Value', 'Visible', 'ZOrder'})
    spread_sets = frozenset({'Accelerator', 'Caption', 'Display3DShading', 'Enabled', 'LinkedCell', 'Locked', 'LockedText', 'PhoneticAccelerator', 'Placement', 'PrintObject', 'Text', 'Value', 'Visible'})
    parts = frozenset({'Border', 'Characters', 'Interior'})


class OptionButtons(ValueDefault):
    vba_type_name = 'OptionButtons'
    kind = 'Radio'
    spread_reads = frozenset({'Accelerator', 'Border', 'Caption', 'Characters', 'Display3DShading', 'Enabled', 'GroupBox', 'Interior', 'LinkedCell', 'Locked', 'LockedText', 'PhoneticAccelerator', 'Placement', 'PrintObject', 'ShapeRange', 'Text', 'Value', 'Visible', 'ZOrder'})
    spread_sets = frozenset({'Accelerator', 'Caption', 'Display3DShading', 'Enabled', 'LinkedCell', 'Locked', 'LockedText', 'PhoneticAccelerator', 'Placement', 'PrintObject', 'Text', 'Value', 'Visible'})
    parts = frozenset({'Border', 'Characters', 'Interior'})
    several = {'visible': True, 'groupbox': NOTHING}


class ListedControls(OldControls):
    """A collection whose items include lists. Its list methods reach every item that holds one (tests/fixtures/old_collection_methods.json)."""

    vba_type_name = ''

    def _lists(self) -> list[_Listed]:
        return [one for one in self.vba_items() if isinstance(one, _Listed)]

    @method
    def AddItem(self, Text: object = MISSING, Index: object = MISSING) -> object:  # noqa: N803 - native named arguments
        for one in self._lists():
            one.vba_call('AddItem', [Text] if Index is MISSING else [Text, Index])
        return True

    @method
    def RemoveItem(self, Index: object = MISSING, Count: object = MISSING) -> object:  # noqa: N803 - native named arguments
        for one in self._lists():
            one.vba_call('RemoveItem', [Index] if Count is MISSING else [Index, Count])
        return True

    @method
    def RemoveAllItems(self) -> object:
        for one in self._lists():
            one.vba_call('RemoveAllItems', [])
        return True


class DropDowns(ValueDefault, ListedControls):
    vba_type_name = 'DropDowns'
    kind = 'Drop'
    spread_reads = frozenset({'Caption', 'Characters', 'Display3DShading', 'DropDownLines', 'Enabled', 'LinkedCell', 'List', 'ListFillRange', 'ListIndex', 'Locked', 'Placement', 'PrintObject', 'Selected', 'ShapeRange', 'Text', 'Value', 'Visible', 'ZOrder'})
    spread_sets = frozenset({'Caption', 'Display3DShading', 'DropDownLines', 'Enabled', 'LinkedCell', 'List', 'ListFillRange', 'ListIndex', 'Locked', 'Placement', 'PrintObject', 'Text', 'Value', 'Visible'})
    parts = frozenset({'Characters'})


class ListBoxes(ValueDefault, ListedControls):
    vba_type_name = 'ListBoxes'
    kind = 'List'
    spread_reads = frozenset({'Display3DShading', 'Enabled', 'LinkedCell', 'List', 'ListFillRange', 'ListIndex', 'Locked', 'MultiSelect', 'Placement', 'PrintObject', 'Selected', 'ShapeRange', 'Value', 'Visible', 'ZOrder'})
    spread_sets = frozenset({'Display3DShading', 'Enabled', 'LinkedCell', 'List', 'ListFillRange', 'ListIndex', 'Locked', 'MultiSelect', 'Placement', 'PrintObject', 'Value', 'Visible'})


def collection(sheet: Worksheet, kind: type[OldControls], index: object) -> Any:
    found = kind(sheet)
    return found if index is MISSING else found.vba_get('Item', [index])


__all__ = ['Button', 'Buttons', 'CheckBox', 'CheckBoxes', 'DropDown', 'DropDowns', 'ListBox', 'ListBoxes', 'OLEFormat', 'OptionButton', 'OptionButtons', 'collection', 'item_of']
