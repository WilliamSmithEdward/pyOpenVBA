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
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import MemberSpec, VBACollection, VBAObject, member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, NULL, VBAArray, VBAInt, error, to_bool, to_number, to_text

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


def _new_control(name: str, value: object) -> MemberSpec:
    """What a new control answers; setting it is refused until what it writes is measured."""

    def get(item: OldControl) -> object:
        del item
        return value

    def put(item: OldControl, wanted: object) -> None:
        del item, wanted
        raise VBAUnsupportedError(f'Setting {name} on a form control is not implemented')

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


def _raising(name: str, failure: tuple[int, str]) -> MemberSpec:
    def get(item: OldControl) -> object:
        del item
        raise error(*failure)

    def put(item: OldControl, wanted: object) -> None:
        del item, wanted
        raise error(*failure)

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


for _spec in (*(_through('drawn', _name) for _name in ('Name', 'Left', 'Top', 'Width', 'Height', 'Placement', 'OnAction')), _through('drawn', 'TopLeftCell', writes=False), _through('drawn', 'BottomRightCell', writes=False)):
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


for _spec in (_through('control', 'LockedText'), _new_control('PhoneticAccelerator', '')):
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

    @member(default=True)
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


for _spec in (*(_new_control(_name, False) for _name in ('AutoSize', 'AddIndent', 'AutoScaleFont', 'CancelButton', 'DefaultButton', 'DismissButton', 'HelpButton')), _new_control('HorizontalAlignment', VBAInt(-4108, 'Long')),
              _new_control('VerticalAlignment', VBAInt(-4108, 'Long')), _new_control('Orientation', VBAInt(-4128, 'Long')), _new_control('ReadingOrder', VBAInt(27, 'Long')), _new_control('Formula', '')):
    Button.vba_add_member(_spec)


class _Valued(OldControl):
    """A control that holds a value and may be linked to a cell."""

    vba_type_name = ''


for _spec in (_through('control', 'Value'), _through('control', 'LinkedCell'), _new_control('Display3DShading', False)):
    _Valued.vba_add_member(_spec)


class CheckBox(_Valued, _Captioned):
    vba_type_name = 'CheckBox'
    kind = 'CheckBox'


class OptionButton(_Valued, _Captioned):
    vba_type_name = 'OptionButton'
    kind = 'Radio'


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
        if not int(to_number(self.control.vba_get('ListCount', []))) or isinstance(self, DropDown):
            raise error(1004, 'Unable to set the Value property')
        self.control.vba_set('Value', value)

    @member
    def Selected(self, Index: object = MISSING) -> object:  # noqa: N803 - native named argument
        if Index is MISSING:
            raise error(*_NOT_THERE)
        return int(to_number(self.control.vba_get('ListIndex', []))) == int(to_number(Index))

    @setter('Selected')
    def _selected(self, value: object, Index: object = MISSING) -> None:  # noqa: N803 - native named argument
        if Index is MISSING:
            raise error(*_NOT_THERE)
        if to_bool(value):
            self.control.vba_set('ListIndex', Index)


for _spec in (_through('control', 'ListFillRange'), _through('control', 'ListCount', writes=False), _through('control', 'ListIndex')):
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
    def Text(self) -> object:
        raise error(13, 'Type mismatch')

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


class OldControls(VBACollection, Created):
    """One kind of a sheet's form controls, or some of them picked by an Array."""

    vba_type_name = ''
    kind: ClassVar[str] = ''

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

    def vba_get(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        """A member of the collection, or else of its first item."""
        if name.lower() in self._vba_members or name.lower() in {key.lower() for key in self._vba_members}:
            return super().vba_get(name, args, named)
        items = self.vba_items()
        if not items:
            raise VBAUnsupportedError(f'{self.vba_type_name}.{name} with nothing in the collection is not implemented')
        return getattr(items[0], 'vba_get')(name, args, named)

    def vba_set(self, name: str, value: object, args: Sequence[object] = (), named: dict[str, object] | None = None, *, by_ref: bool = False) -> None:
        """A member set on every item."""
        items = self.vba_items()
        if not items:
            raise VBAUnsupportedError(f'{self.vba_type_name}.{name} with nothing in the collection is not implemented')
        for item in items:
            getattr(item, 'vba_set')(name, value, args, named, by_ref=by_ref)


class Buttons(OldControls):
    vba_type_name = 'Buttons'
    kind = 'Button'


class CheckBoxes(OldControls):
    vba_type_name = 'CheckBoxes'
    kind = 'CheckBox'


class OptionButtons(OldControls):
    vba_type_name = 'OptionButtons'
    kind = 'Radio'


class DropDowns(OldControls):
    vba_type_name = 'DropDowns'
    kind = 'Drop'


class ListBoxes(OldControls):
    vba_type_name = 'ListBoxes'
    kind = 'List'


def collection(sheet: Worksheet, kind: type[OldControls], index: object) -> Any:
    found = kind(sheet)
    return found if index is MISSING else found.vba_get('Item', [index])


__all__ = ['Button', 'Buttons', 'CheckBox', 'CheckBoxes', 'DropDown', 'DropDowns', 'ListBox', 'ListBoxes', 'OLEFormat', 'OptionButton', 'OptionButtons', 'collection', 'item_of']
