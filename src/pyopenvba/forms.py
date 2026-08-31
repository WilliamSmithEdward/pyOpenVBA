"""
UserForm designer streams: the control tree a form's code-behind does not carry.

pyOpenVBA reads and writes a project's *code*.  A UserForm's *design* --
which controls exist, how they nest, and what each one's properties are --
lives in a separate storage inside the same CFB, named for the form and
sitting beside ``VBA/``::

    /VBA/EntryForm            the code-behind (handled by vba.py)
    /EntryForm/f              the sites: which controls, in what order
    /EntryForm/o              each control's own property record
    /EntryForm/\\x03VBFrame    text: the form's own non-default properties
    /EntryForm/i06/{f,o}      a Frame: its children live in their own storage
    /EntryForm/i06/i08/{f,o}  a Page inside a MultiPage

This module reads that tree and writes it back ([MS-OFORMS]).  Property
names come from :mod:`pyopenvba._oforms_records`, which carries one table
per control class.

MSForms stores a property only when it differs from that control's
default, so the set of stored properties is the set the developer chose.
That is not something a live host can tell you: a sited control reports
inherited, default and chosen values indistinguishably.

Conservative by construction.  Every structure here is length-prefixed or
counted, so a misreading collapses immediately rather than yielding a
plausible-looking control list, and this raises :class:`FormParseError`
instead of guessing.  Four independent checks have to agree:

1. the site count is consistent with ``SiteDepthsAndTypes``,
2. ``CountOfBytes`` runs exactly to the end of the ``f`` stream,
3. the per-site ``ObjectStreamSize`` values sum to exactly ``len(o)``,
4. every child storage is claimed by a site that can contain one.

Writing is lossless: an unedited form serializes to the bytes it was read
from, which is the gate the fixtures pin.  Editing a property rewrites
only the affected control's record and patches that site's
``ObjectStreamSize`` in place -- a fixed-width field, so no length in the
``f`` stream moves.
"""

from __future__ import annotations

import struct
from collections.abc import Sequence
from copy import deepcopy
from dataclasses import dataclass, field

from pyopenvba._oforms_pages import (
    EMPTY_PAGE_PROPERTIES,
    NEW_TAB_FLAGS,
    TAB_HEADROOM,
    PageBookkeeping,
    new_tab_string,
    parse_page_bookkeeping,
    parse_string_array,
    serialize_page_bookkeeping,
    serialize_string_array,
)
from pyopenvba._oforms_records import (
    FORM_SPEC,
    SPECS_BY_CACHE_INDEX,
    TEXT_PROPS_SPEC,
    ParsedRecord,
    Size,
    StoredString,
    parse_record,
    serialize_record,
)
from pyopenvba.cfb import CFB
from pyopenvba.exceptions import FormParseError

# [MS-OFORMS] 2.4.5 FormEmbeddedActiveXControlCached: a ClsidCacheIndex
# below 0x7FFF names a control class directly.  The VBE persists the
# specific morph indices (23/25/26/...) rather than generic MorphData,
# but 15 stays legal from other producers and resolves through
# DisplayStyle in the control's own record.
CONTROL_CLASSES: dict[int, str] = {
    7: "MSForms.Form",
    12: "MSForms.Image",
    14: "MSForms.Frame",
    15: "MSForms.MorphData",
    16: "MSForms.SpinButton",
    17: "MSForms.CommandButton",
    18: "MSForms.TabStrip",
    21: "MSForms.Label",
    23: "MSForms.TextBox",
    24: "MSForms.ListBox",
    25: "MSForms.ComboBox",
    26: "MSForms.CheckBox",
    27: "MSForms.OptionButton",
    28: "MSForms.ToggleButton",
    47: "MSForms.ScrollBar",
    57: "MSForms.MultiPage",
}

# fmDisplayStyle ([MS-OFORMS] 2.5.20.1): what a generic MorphData is.
MORPH_DISPLAY_STYLES: dict[int, str] = {
    1: "MSForms.TextBox",
    2: "MSForms.ListBox",
    3: "MSForms.ComboBox",
    4: "MSForms.CheckBox",
    5: "MSForms.OptionButton",
    6: "MSForms.ToggleButton",
    7: "MSForms.ComboBox",
}

# A ClsidCacheIndex at or above this is an index into the form's class
# table, naming an ActiveX control whose class this cannot resolve.
_CLASS_TABLE_BASE = 0x8000

# Controls that own a child storage rather than a slice of `o`.
_CONTAINER_CLASSES = frozenset({7, 14, 57})

_UNKNOWN_CONTROL = "MSForms.Control"
_ACTIVEX_CONTROL = "ActiveX.Control"

_FORM_CONTROL_VERSION = (0, 4)
_CONTROL_RECORD_VERSION = (0, 2)

# FormPropMask bits whose payload sits in FormStreamData after the
# DataBlock, and so has to be stepped over to reach the sites.
_FORM_MASK_MOUSE_ICON = 1 << 15
_FORM_MASK_FONT = 1 << 20
_FORM_MASK_PICTURE = 1 << 21

# StdFont vs TextProps ([MS-OFORMS] 2.4.6), by the first DWORD of the GUID.
_GUID_STDFONT = 0x0BE35203
_GUID_TEXTPROPS = 0xAFC20920

_VBFRAME_STREAM = "\x03VBFrame"

# Site DataBlock field widths, in mask-bit order.  Bit 8 carries no fixed
# field: reading two bytes for it puts every name two characters late.
_SITE_FIELDS: tuple[tuple[int, str, int], ...] = (
    (0, "NameData", 4),
    (1, "TagData", 4),
    (2, "ID", 4),
    (3, "HelpContextID", 4),
    (4, "BitFlags", 4),
    (5, "ObjectStreamSize", 4),
    (6, "TabIndex", 2),
    (7, "ClsidCacheIndex", 2),
    (9, "GroupID", 2),
    (11, "ControlTipTextData", 4),
    (12, "RuntimeLicKeyData", 4),
    (13, "ControlSourceData", 4),
    (14, "RowSourceData", 4),
)
_SIGNED_SITE_FIELDS = frozenset({"ID", "HelpContextID", "TabIndex"})

# The SiteExtraDataBlock: Name and Tag, then the position bit 8 selects,
# then the rest.  Each string is sized by its own DataBlock length field.
_SITE_STRINGS: tuple[tuple[str, str], ...] = (
    ("NameData", "Name"),
    ("TagData", "Tag"),
)
_SITE_STRINGS_AFTER_POSITION: tuple[tuple[str, str], ...] = (
    ("ControlTipTextData", "ControlTipText"),
    ("RuntimeLicKeyData", "RuntimeLicKey"),
    ("ControlSourceData", "ControlSource"),
    ("RowSourceData", "RowSource"),
)


@dataclass
class FormControl:
    """One control on a form, as the designer streams describe it."""

    name: str
    kind: str
    """``MSForms.TextBox`` and friends; ``ActiveX.Control`` for a control
    whose class lives in the form's class table, and ``MSForms.Control``
    when the ClsidCacheIndex is one this does not know."""
    id: int
    clsid_cache_index: int
    tab_index: int | None
    object_stream_size: int
    """Bytes this control occupies in its parent's ``o`` stream; zero for
    a container, whose own record is the ``f`` of its child storage."""
    record: ParsedRecord | None
    """The parsed property record, or ``None`` for a control whose class
    has no table here."""
    children: tuple[FormControl, ...] = ()

    @property
    def is_container(self) -> bool:
        return bool(self.children) or self.clsid_cache_index in _CONTAINER_CLASSES

    @property
    def properties_set(self) -> int:
        """Raw PropMask: the properties the developer set, as a bit set."""
        return 0 if self.record is None else self.record.mask

    @property
    def property_mask_width(self) -> int:
        """4 or 8 bytes.  MorphData controls carry the wider mask, so a
        bit index is only meaningful together with this."""
        return 8 if self.record is not None and self.record.spec.mask64 else 4

    def properties(self) -> dict[str, object]:
        """Every property this control stores, by name."""
        return {} if self.record is None else self.record.properties()

    def get(self, name: str) -> object:
        """One stored property, or ``None`` when the control does not set it."""
        return self.properties().get(name)

    def set_property(self, name: str, value: object) -> None:
        """Set or clear one property.  ``None`` clears it.

        A string goes to the string table, an integer to the DataBlock,
        and a :class:`~pyopenvba._oforms_records.Size` to the size field.
        Clearing removes the mask bit, which is how the control goes back
        to inheriting the default.
        """
        if self.record is None:
            raise FormParseError(
                f"{self.name!r} is a {self.kind} and has no property table here"
            )
        _set_on_record(self.record, name, value, self.name)


def _set_on_record(
    record: ParsedRecord, name: str, value: object, owner: str
) -> None:
    """Route a property edit to the right table by the value's type."""
    if isinstance(value, Size):
        record.set_size(value.width, value.height)
    elif isinstance(value, str):
        record.set_string(name, value)
    elif value is None:
        if any(f.name == name and f.kind == "str" for f in record.spec.extra):
            record.set_string(name, None)
        else:
            record.set_value(name, None)
    elif isinstance(value, bool):
        record.set_value(name, int(value))
    elif isinstance(value, int):
        record.set_value(name, value)
    else:
        raise FormParseError(
            f"{owner}: cannot store {type(value).__name__} in property {name!r}"
        )


@dataclass
class VBAForm:
    """A UserForm's designer surface."""

    name: str
    designer_source: str
    """The ``\\x03VBFrame`` text, the same block a VBE export writes."""
    controls: tuple[FormControl, ...] = ()
    _levels: list[_Level] = field(default_factory=lambda: [], repr=False)
    _encoding: str = field(default="cp1252", repr=False)
    _removed_storages: list[tuple[list[str], str]] = field(
        default_factory=lambda: [], repr=False
    )

    @property
    def properties_set(self) -> int:
        """The form's own PropMask, read the same way as a control's."""
        return self._levels[0].record.mask if self._levels else 0

    def properties(self) -> dict[str, object]:
        """The form's own stored properties."""
        return self._levels[0].record.properties() if self._levels else {}

    def get(self, name: str) -> object:
        """One of the form's own properties, or ``None`` if it sets none."""
        return self.properties().get(name)

    def set_property(self, name: str, value: object) -> None:
        """Set or clear one of the form's own properties (Caption, Zoom...).

        The form's record heads its ``f`` stream rather than sitting in
        ``o``, but it is edited exactly like a control's.
        """
        if not self._levels:
            raise FormParseError(f"form {self.name!r} has no parsed record")
        _set_on_record(self._levels[0].record, name, value, self.name)

    def walk(self) -> list[FormControl]:
        """Every control, depth-first, containers before their children."""
        out: list[FormControl] = []

        def visit(controls: Sequence[FormControl]) -> None:
            for control in controls:
                out.append(control)
                visit(control.children)

        visit(self.controls)
        return out

    def control(self, name: str) -> FormControl:
        """One control by name, at any depth.  MSForms keeps names unique."""
        for control in self.walk():
            if control.name == name:
                return control
        raise KeyError(f"no control named {name!r} on form {self.name!r}")

    def add_control(
        self,
        kind: str,
        name: str,
        *,
        container: str | None = None,
        left: float = 6,
        top: float = 6,
        width: float | None = None,
        height: float | None = None,
    ) -> FormControl:
        """Add a control to the form, or to a container already on it.

        ``kind`` is the MSForms class without its library prefix --
        ``"CommandButton"``, ``"TextBox"``, ``"Label"`` and so on.
        Geometry is in points, the unit the designer shows.

        The control's id comes from the form's ``NextAvailableID``, which
        is what MSForms uses to name a container's storage, so it is
        advanced here exactly as the VBE advances it.

        Containers (``Frame``, ``MultiPage``, ``Page``) are refused: each
        needs a storage of its own, which is a structural change beyond
        writing these two streams.
        """
        if not self._levels:
            raise FormParseError(f"form {self.name!r} has no parsed record")
        if kind in ("Page", "Form"):
            # A page belongs to a MultiPage, not to a surface; add_page
            # writes the tab and the bookkeeping that come with it.
            raise FormParseError(
                f"cannot add a {kind} here: use add_page(multipage) instead"
            )
        cache_index = _CACHE_INDEX_BY_KIND.get(kind)
        if cache_index is None:
            raise FormParseError(
                f"unknown control kind {kind!r}; expected one of "
                f"{', '.join(sorted(_CACHE_INDEX_BY_KIND))}"
            )
        if any(c.name.casefold() == name.casefold() for c in self.walk()):
            raise FormParseError(f"the form already has a control named {name!r}")

        level = self._level_for(container)
        # NextAvailableID is the highest id already handed out, not the
        # next free one: measured against Excel, whose own Controls.Add on
        # a form whose highest id was 13 produced a control with id 14 and
        # left the field at 14. Using the field as-is collides with the
        # last control, and MSForms then refuses to load the form.
        site_id = self._next_id(level)

        default_w, default_h = _DEFAULT_SIZE_PT.get(kind, _DEFAULT_SIZE_FALLBACK)
        size_w = points_to_himetric(default_w if width is None else width)
        size_h = points_to_himetric(default_h if height is None else height)
        site = _new_site(
            name,
            site_id,
            cache_index,
            len(level.sites),
            points_to_himetric(left),
            points_to_himetric(top),
            encoding=self._encoding,
            container=kind in ("Frame", "MultiPage"),
        )
        if kind in ("Frame", "MultiPage"):
            # A container holds no slice of `o`; its own record heads the
            # `f` of a storage created for it.
            child = _Level(
                path=[*level.path, _storage_name(site_id)],
                f_raw=b"",
                o_raw=b"",
                stream=_new_container_stream(kind, name, size_w, size_h),
                controls=[],
                created=True,
                clsid=_FRAME_CLSID if kind == "Frame" else _MULTIPAGE_CLSID,
                compobj=(
                    _FRAME_COMPOBJ if kind == "Frame" else _MULTIPAGE_COMPOBJ
                ),
            )
            if kind == "MultiPage":
                child.x_raw = b""
                child.pages = PageBookkeeping()
            self._levels.append(child)
            record = child.stream.record
        else:
            record = _new_record(kind, cache_index, name)
            _adopt_font(record, level)
            record.set_size(size_w, size_h)
        level.stream.sites.append(site)
        level.stream.sites_structurally_changed = True
        control = FormControl(
            name=name,
            kind=f"MSForms.{kind}",
            id=site_id,
            clsid_cache_index=cache_index,
            tab_index=site.tab_index,
            object_stream_size=0,
            record=record,
        )
        level.controls.append(control)
        self._reindex()
        if kind == "MultiPage":
            self._seed_multipage(name, site_id)
        return control

    def _seed_multipage(self, name: str, site_id: int) -> None:
        """Give a new MultiPage its TabStrip and the two pages Excel adds."""
        level = next(
            lvl for lvl in self._levels if _storage_id(lvl.path[-1]) == site_id
        )
        tabstrip_id = self._next_id()
        site = _new_site("", tabstrip_id, 18, 0, 0, 0, encoding=self._encoding)
        # The TabStrip has no name of its own; MSForms sites it unnamed.
        site.mask &= ~(1 << 0)
        site.values.pop("NameData", None)
        site.strings.pop("Name", None)
        level.stream.sites.append(site)
        level.stream.sites_structurally_changed = True
        level.controls.append(
            FormControl(
                name="",
                kind="MSForms.TabStrip",
                id=tabstrip_id,
                clsid_cache_index=18,
                tab_index=0,
                object_stream_size=0,
                record=_new_tabstrip_record(),
            )
        )
        # The bookkeeping names the TabStrip's site id, not the MultiPage's.
        assert level.pages is not None
        level.pages.identifier = tabstrip_id
        level.pages.page_props.append(EMPTY_PAGE_PROPERTIES)
        self._reindex()
        self.add_page(name, name="Page1")
        self.add_page(name, name="Page2")

    def remove_control(self, name: str) -> None:
        """Remove a control by name.

        Containers are refused for the same reason they cannot be added:
        their storage would be orphaned, and the next read would refuse
        the form rather than quietly lose their children.
        """
        for level in self._levels:
            for index, control in enumerate(level.controls):
                if control.name != name:
                    continue
                if control.clsid_cache_index == 7:
                    # A page is also a tab of its MultiPage; remove_page
                    # takes the tab and the bookkeeping with it.
                    raise FormParseError(
                        f"cannot remove the page {name!r} here: "
                        "use remove_page(name) instead"
                    )
                child = self._child_level(level, control.id)
                if child is not None:
                    # Its storage goes too, with everything under it: the
                    # next read refuses a form whose child storage no site
                    # claims, which is what makes this mandatory.
                    self._drop_subtree(child)
                del level.controls[index]
                del level.stream.sites[index]
                level.stream.sites_structurally_changed = True
                self._reindex()
                return
        raise KeyError(f"no control named {name!r} on form {self.name!r}")

    def add_page(
        self, multipage: str, name: str | None = None, caption: str | None = None
    ) -> FormControl:
        """Add a page to a MultiPage already on the form.

        A page is a container control *and* a tab, so this writes both:
        a site and a storage of its own, and an entry in each of the
        MultiPage's five TabStrip arrays plus its page bookkeeping.
        """
        parent, control = self._find(multipage)
        if control.clsid_cache_index != 57:
            raise FormParseError(f"{multipage!r} is a {control.kind}, not a MultiPage")
        level = self._child_level(parent, control.id)
        if level is None:
            raise FormParseError(f"{multipage!r} has no storage of its own")
        tabs = self._tabstrip(level)
        book = self._bookkeeping(level)

        page_id = self._next_id(level)
        name = name or f"Page{len(book.page_ids) + 1}"
        # Page names are scoped to their MultiPage, not to the form: Excel
        # gives a second MultiPage its own Page1 and Page2 while the first
        # still has them, and pages do not appear in Designer.Controls at
        # all.  So this checks the siblings, not the whole tree.
        if any(c.name.casefold() == name.casefold() for c in level.controls):
            raise FormParseError(
                f"{multipage!r} already has a page named {name!r}"
            )

        site = _new_site(
            name, page_id, 7, len(level.sites), *_DEFAULT_PAGE_ORIGIN,
            encoding=self._encoding, container=True,
        )
        # Only the selected page carries the active flag, and that is the
        # first one; every page added after it is inactive.
        if book.page_ids:
            site.values["BitFlags"] = _INACTIVE_PAGE_BITFLAGS
        level.stream.sites.append(site)
        level.stream.sites_structurally_changed = True

        child = _Level(
            path=[*level.path, _storage_name(page_id)],
            f_raw=b"",
            o_raw=b"",
            stream=_new_container_stream("Page", None, *_DEFAULT_PAGE_SIZE),
            controls=[],
            created=True,
            clsid=_PAGE_CLSID,
            compobj=_PAGE_COMPOBJ,
        )
        self._levels.append(child)

        self._set_tabs(
            tabs,
            {
                "Items": caption or name,
                "TabNames": _next_tab_name(tabs),
                "TipStrings": "",
                "Tags": "",
                "Accelerators": "",
            },
            add=True,
        )
        book.add(page_id)

        page = FormControl(
            name=name,
            kind="MSForms.Form",
            id=page_id,
            clsid_cache_index=7,
            tab_index=site.tab_index,
            object_stream_size=0,
            record=child.stream.record,
        )
        level.controls.append(page)
        self._reindex()
        return page

    def remove_page(self, name: str, multipage: str | None = None) -> None:
        """Remove one page from its MultiPage, tab and storage included.

        Page names are unique within a MultiPage rather than across the
        form, so ``multipage`` names which one when two of them share a
        page name.
        """
        if multipage is None:
            level, control = self._find(name)
        else:
            parent, owner = self._find(multipage)
            owned = self._child_level(parent, owner.id)
            if owned is None:
                raise FormParseError(f"{multipage!r} has no storage of its own")
            level = owned
            control = next(
                (c for c in level.controls if c.name == name),
                None,
            ) or _missing_page(name, multipage)
        if control.clsid_cache_index != 7:
            raise FormParseError(f"{name!r} is a {control.kind}, not a page")
        index = level.controls.index(control)
        tabs = self._tabstrip(level)
        book = self._bookkeeping(level)
        # The TabStrip is the MultiPage's own first site, so a page's tab
        # sits one earlier in the arrays than its site does in the list.
        self._set_tabs(tabs, {}, add=False, at=index - 1)
        book.remove(control.id)

        child = self._child_level(level, control.id)
        if child is not None:
            self._drop_subtree(child)
        del level.controls[index]
        del level.stream.sites[index]
        level.stream.sites_structurally_changed = True
        self._reindex()

    def _find(self, name: str) -> tuple[_Level, FormControl]:
        """The level and control for a name, at any depth."""
        for level in self._levels:
            for control in level.controls:
                if control.name == name:
                    return level, control
        raise KeyError(f"no control named {name!r} on form {self.name!r}")

    def _next_id(self, level: _Level | None = None) -> int:
        """Allocate a site id, and tell every container that contains it.

        A container's own ``NextAvailableID`` is the highest id anywhere
        beneath it, not just among its direct children -- the fixture's
        MultiPage carries 11, which is a control two levels down on one of
        its pages.  So the id is recorded on every ancestor up to the form.
        """
        root = self._levels[0]
        site_id = max(
            root.record.values.get("NextAvailableID", 0),
            max((c.id for c in self.walk()), default=0),
        ) + 1
        target = level or root
        for ancestor in self._levels:
            if target.path[:len(ancestor.path)] == ancestor.path:
                ancestor.record.set_value("NextAvailableID", site_id)
                if ancestor.record.has("ShapeCookie"):
                    ancestor.record.set_value(
                        "ShapeCookie",
                        ancestor.record.values.get("ShapeCookie", 0) + 1,
                    )
        return site_id

    @staticmethod
    def _tabstrip(level: _Level) -> ParsedRecord:
        """A MultiPage's hidden TabStrip, which owns the tab arrays."""
        for control in level.controls:
            if control.clsid_cache_index == 18 and control.record is not None:
                return control.record
        raise FormParseError(
            f"{'/'.join(level.path)}: the MultiPage has no TabStrip record"
        )

    def _bookkeeping(self, level: _Level) -> PageBookkeeping:
        if level.pages is None:
            if not level.x_raw:
                raise FormParseError(
                    f"{'/'.join(level.path)}: the MultiPage has no 'x' stream"
                )
            level.pages = parse_page_bookkeeping(level.x_raw)
        return level.pages

    def _set_tabs(
        self,
        tabs: ParsedRecord,
        added: dict[str, str],
        *,
        add: bool,
        at: int = -1,
    ) -> None:
        """Add or remove one entry across every TabStrip array at once.

        The five arrays and the flag tail all carry one element per tab;
        letting them drift apart is what a partial edit would do.
        """
        count = 0
        for name, size_field in _TAB_ARRAYS:
            entries = parse_string_array(tabs.arrays.get(name, b""), self._encoding)
            if add:
                entries.append(new_tab_string(added[name]))
            else:
                if not -len(entries) <= at < len(entries):
                    raise FormParseError(f"no tab at index {at} in {name}")
                del entries[at]
            blob = serialize_string_array(entries, self._encoding)
            tabs.arrays[name] = blob
            tabs.set_value(size_field, len(blob))
            if count and len(entries) != count:
                # The five arrays carry one element per tab; if they were
                # not the same length going in, this edit would silently
                # leave them further apart.
                raise FormParseError(
                    f"TabStrip arrays disagree: {name} has {len(entries)} "
                    f"entries where the previous array had {count}"
                )
            count = len(entries)
        flags = [
            int.from_bytes(tabs.tail_raw[i:i + 4], "little")
            for i in range(0, len(tabs.tail_raw), 4)
        ]
        if add:
            flags.append(NEW_TAB_FLAGS)
        elif flags:
            del flags[at]
        tabs.tail_raw = b"".join(f.to_bytes(4, "little") for f in flags)
        tabs.set_value("TabData", count)
        # Capacity runs ahead of the count and Excel never shrinks it.
        allocated = max(tabs.values.get("TabsAllocated", 0), count + TAB_HEADROOM)
        tabs.set_value("TabsAllocated", allocated)

    def _drop_subtree(self, level: _Level) -> None:
        """Forget a container's level, and every level beneath it."""
        depth = len(level.path)
        for other in [
            other for other in self._levels if other.path[:depth] == level.path
        ]:
            self._levels.remove(other)
        if not level.created:
            self._removed_storages.append((level.path[:-1], level.path[-1]))

    def _level_for(self, container: str | None) -> _Level:
        """The level a new control belongs to: the form, or a container."""
        if container is None:
            return self._levels[0]
        for level in self._levels:
            for control in level.controls:
                if control.name != container:
                    continue
                child = self._child_level(level, control.id)
                if child is None:
                    raise FormParseError(
                        f"{container!r} is a {control.kind} and holds no controls"
                    )
                return child
        raise KeyError(f"no container named {container!r} on form {self.name!r}")

    def _child_level(self, parent: _Level, site_id: int) -> _Level | None:
        """The level behind a container site's own storage, if it has one."""
        for level in self._levels:
            if (
                level.path[:-1] == parent.path
                and _storage_id(level.path[-1]) == site_id
            ):
                return level
        return None

    def _reindex(self) -> None:
        """Rebuild the control tree after a structural change."""

        def children_of(level: _Level) -> tuple[FormControl, ...]:
            for control in level.controls:
                child = self._child_level(level, control.id)
                if child is not None:
                    control.children = children_of(child)
            return tuple(level.controls)

        self.controls = children_of(self._levels[0])

    def write_back(self, cfb: CFB) -> bool:
        """Write every edited record back into the project's CFB.

        Only the streams whose bytes actually changed are written, so an
        unedited form leaves the CFB untouched.  Returns whether anything
        was written, which the host uses to decide that the save mutates.
        """
        changed = False
        for parent, name in self._removed_storages:
            cfb.remove_storage_at(parent, name)
            changed = True
        self._removed_storages.clear()
        for level in self._levels:
            if not level.created:
                continue
            # A container's storage is bound by its CLSID and by the
            # CompObj naming what fm20 should treat it as; either one
            # wrong and the container loads without erroring and simply
            # does not appear.
            cfb.add_substorage_at(level.path[:-1], level.path[-1], level.clsid)
            cfb.add_stream_at(level.path, "f", b"")
            cfb.add_stream_at(level.path, "o", b"")
            cfb.add_stream_at(level.path, _COMPOBJ_STREAM, level.compobj)
            if level.pages is not None:
                # A MultiPage also owns the page bookkeeping.
                cfb.add_stream_at(level.path, "x", b"")
            level.created = False
            changed = True
        for level in self._levels:
            f_bytes, o_bytes = level.serialize(self._encoding)
            if o_bytes != level.o_raw:
                cfb.write_stream_at(level.path, "o", o_bytes)
                level.o_raw = o_bytes
                changed = True
            if f_bytes != level.f_raw:
                cfb.write_stream_at(level.path, "f", f_bytes)
                level.f_raw = f_bytes
                changed = True
            if level.pages is not None:
                x_bytes = serialize_page_bookkeeping(level.pages)
                if x_bytes != level.x_raw:
                    cfb.write_stream_at(level.path, "x", x_bytes)
                    level.x_raw = x_bytes
                    changed = True
        return changed


@dataclass
class _Site:
    """One entry of a container's site array, kept losslessly."""

    mask: int
    values: dict[str, int] = field(default_factory=lambda: {})
    strings: dict[str, StoredString] = field(default_factory=lambda: {})
    position: tuple[int, int] | None = None
    """fmPosition (left, top) in HIMETRIC, present when mask bit 8 is set.

    Bit 8 selects no DataBlock field -- reading two bytes for it there puts
    every name two characters late -- but it does select this pair in the
    ExtraDataBlock, after the Name and Tag strings.
    """
    pads: dict[str, bytes] = field(default_factory=lambda: {})

    @property
    def name(self) -> str:
        stored = self.strings.get("Name")
        return "" if stored is None else stored.text

    @property
    def id(self) -> int:
        return self.values.get("ID", 0)

    @property
    def clsid_cache_index(self) -> int:
        return self.values.get("ClsidCacheIndex", 0)

    @property
    def tab_index(self) -> int | None:
        return self.values.get("TabIndex")

    @property
    def object_stream_size(self) -> int:
        return self.values.get("ObjectStreamSize", 0)

    @property
    def is_container(self) -> bool:
        return self.clsid_cache_index in _CONTAINER_CLASSES

    def set_object_stream_size(self, size: int) -> None:
        if "ObjectStreamSize" not in self.values and size:
            self.mask |= 1 << 5
        self.values["ObjectStreamSize"] = size


@dataclass
class _FormStream:
    """A parsed ``f`` stream: the container's record, then its sites."""

    record: ParsedRecord
    sites: list[_Site]
    mouse_icon_raw: bytes = b""
    font_raw: bytes = b""
    picture_raw: bytes = b""
    class_table_present: bool = False
    class_table_raw: bytes = b""
    depths_raw: bytes = b""
    trailing_raw: bytes = b""
    sites_structurally_changed: bool = False
    """Set when a site is added or removed, so SiteDepthsAndTypes is
    recomposed rather than replayed."""

    def serialize(self, encoding: str) -> bytes:
        out = bytearray(serialize_record(self.record, encoding))
        mask = self.record.mask
        for bit, blob, what in (
            (_FORM_MASK_MOUSE_ICON, self.mouse_icon_raw, "MouseIcon"),
            (_FORM_MASK_FONT, self.font_raw, "Font"),
            (_FORM_MASK_PICTURE, self.picture_raw, "Picture"),
        ):
            if not mask & bit:
                continue
            if not blob:
                raise FormParseError(f"masked {what} FormStreamData has no bytes")
            out += blob
        if self.class_table_present:
            out += self.class_table_raw

        sites = [_serialize_site(site, encoding) for site in self.sites]
        depths = (
            _compose_depths(len(self.sites))
            if self.sites_structurally_changed
            else self.depths_raw
        )
        out += struct.pack("<II", len(self.sites), len(depths) + sum(map(len, sites)))
        out += depths
        for site in sites:
            out += site
        out += self.trailing_raw
        return bytes(out)


def _compose_depths(count: int) -> bytes:
    """SiteDepthsAndTypes for ``count`` uniform sites (depth 0, ST_Ole).

    Written in the run-length form Excel itself always uses: one counted
    entry per run of up to 127.  The per-entry form is spec-legal too, but
    matching the only producer MSForms is tested against costs nothing.
    """
    entries = bytearray()
    remaining = count
    while remaining > 0:
        run = min(remaining, 0x7F)
        entries += bytes((0x00, 0x80 | run, 0x01))
        remaining -= run
    over = len(entries) % 4
    return bytes(entries) + bytes(4 - over if over else 0)


def _serialize_site(site: _Site, encoding: str) -> bytes:
    """Rebuild one site record.  An unedited site round-trips exactly."""
    # Refresh the length fields of edited strings before the DataBlock.
    string_bytes: dict[str, bytes] = {}
    for len_field, name in _SITE_STRINGS + _SITE_STRINGS_AFTER_POSITION:
        stored = site.strings.get(name)
        if stored is None:
            continue
        raw = (
            stored.raw
            if not stored.edited
            else stored.text.encode(
                encoding if stored.compressed else "utf-16-le", "replace"
            )
        )
        site.values[len_field] = (len(raw) & 0x7FFFFFFF) | (
            0x80000000 if stored.compressed else 0
        )
        string_bytes[name] = raw

    body = bytearray()
    header = 8  # version(2) + cbSite(2) + mask(4)
    for bit, name, size in _SITE_FIELDS:
        if not site.mask & (1 << bit):
            continue
        body += _align_gap(body, header, size, site.pads.get(f"before:{name}", b""))
        body += _pack_site(site.values.get(name, 0), size, name in _SIGNED_SITE_FIELDS)
    body += _align_gap(body, header, 4, site.pads.get("data:end", b""))

    def write_string(name: str) -> None:
        raw = string_bytes.get(name)
        if raw is None:
            return
        body.extend(raw)
        body.extend(_align_gap(body, header, 4, site.pads.get(f"str:{name}", b"")))

    for _, name in _SITE_STRINGS:
        write_string(name)
    if site.mask & (1 << 8):
        left, top = site.position or (0, 0)
        body += struct.pack("<ii", left, top)
    for _, name in _SITE_STRINGS_AFTER_POSITION:
        write_string(name)

    cb_site = len(body) + 4  # cbSite counts from the mask
    if cb_site > 0xFFFF:
        # The same u16 wrap hazard as a record's cb: refuse, never corrupt.
        raise FormParseError(
            f"site {site.name!r}: too much data for one site ({cb_site} bytes; "
            "the format caps a site at 65535) -- a name, tag, tip or source "
            "is too long"
        )
    return struct.pack("<HHI", 0x0000, cb_site, site.mask & 0xFFFFFFFF) + bytes(body)


def _pack_site(value: int, size: int, signed: bool) -> bytes:
    if not signed:
        value &= (1 << (size * 8)) - 1
    return struct.pack("<" + ({2: "h", 4: "i"} if signed else {2: "H", 4: "I"})[size],
                       value)


def _align_gap(body: bytearray, header: int, size: int, captured: bytes) -> bytes:
    over = (len(body) + header) % size
    if not over:
        return b""
    needed = size - over
    return captured if len(captured) == needed else bytes(needed)


@dataclass
class _Level:
    """One container's streams: the form itself, a Frame, or a Page."""

    path: list[str]
    f_raw: bytes
    o_raw: bytes
    stream: _FormStream
    controls: list[FormControl]
    created: bool = False
    """Set on a level whose storage does not exist yet, so write_back
    creates it before writing its streams."""
    clsid: bytes = b""
    compobj: bytes = b""
    x_raw: bytes = b""
    """A MultiPage's page bookkeeping, verbatim as read."""
    pages: PageBookkeeping | None = None
    """The same bookkeeping parsed, once a page edit needs it."""

    @property
    def record(self) -> ParsedRecord:
        """The container's own property record."""
        return self.stream.record

    @property
    def sites(self) -> list[_Site]:
        return self.stream.sites

    def serialize(self, encoding: str) -> tuple[bytes, bytes]:
        """Rebuild this level's ``f`` and ``o``.

        ``o`` is the sites' records back to back, so each site's
        ObjectStreamSize is refreshed first and ``f`` is then rebuilt
        whole -- nothing inside it is an absolute offset, cbForm and
        cbSites both being lengths.
        """
        chunks: list[bytes] = []
        for site, control in zip(self.sites, self.controls, strict=True):
            if control.record is None or site.is_container:
                # A container keeps its own record in its child storage's
                # `f`, and holds no slice of `o`.
                continue
            record = serialize_record(control.record, encoding)
            chunks.append(record)
            site.set_object_stream_size(len(record))
            control.object_stream_size = len(record)
        return self.stream.serialize(encoding), b"".join(chunks)


# ---------------------------------------------------------------------------
# Bounds-checked reader
# ---------------------------------------------------------------------------

class _Reader:
    """Little-endian reader that refuses to walk off the end."""

    def __init__(self, data: bytes) -> None:
        self._data = data
        self.pos = 0

    def _need(self, count: int) -> None:
        if count < 0 or self.pos + count > len(self._data):
            raise FormParseError("designer stream walked out of bounds")

    def u8(self) -> int:
        self._need(1)
        value = self._data[self.pos]
        self.pos += 1
        return value

    def u16(self) -> int:
        self._need(2)
        value = int(struct.unpack_from("<H", self._data, self.pos)[0])
        self.pos += 2
        return value

    def u32(self) -> int:
        self._need(4)
        value = int(struct.unpack_from("<I", self._data, self.pos)[0])
        self.pos += 4
        return value

    def take(self, count: int) -> bytes:
        self._need(count)
        value = self._data[self.pos:self.pos + count]
        self.pos += count
        return value

    def i16(self) -> int:
        self._need(2)
        value = int(struct.unpack_from("<h", self._data, self.pos)[0])
        self.pos += 2
        return value

    def i32(self) -> int:
        self._need(4)
        value = int(struct.unpack_from("<i", self._data, self.pos)[0])
        self.pos += 4
        return value

    def skip(self, count: int) -> None:
        self._need(count)
        self.pos += count

    def align(self, base: int, size: int) -> bytes:
        """Align to ``size`` relative to ``base`` ([MS-OFORMS] 2.1.1.2.4).

        Returns the bytes stepped over, which the writer replays: the spec
        leaves their value undefined, so zeroing them on write would not
        be byte-identical.
        """
        over = (self.pos - base) % size
        return self.take(size - over) if over else b""


# ---------------------------------------------------------------------------
# The `f` stream
# ---------------------------------------------------------------------------

def _read_picture_blob(reader: _Reader) -> bytes:
    """A GuidAndPicture, captured whole so it can be replayed."""
    start = reader.pos
    reader.skip(16)  # GUID
    reader.u32()     # Preamble
    reader.skip(reader.u32())
    end = reader.pos
    reader.pos = start
    return reader.take(end - start)


def _read_font_blob(reader: _Reader) -> bytes:
    """A StdFont or TextProps blob, captured whole ([MS-OFORMS] 2.4.6)."""
    start = reader.pos
    guid = reader.take(16)
    tag = int(struct.unpack_from("<I", guid, 0)[0])
    if tag == _GUID_STDFONT:
        if reader.u8() != 0x01:
            raise FormParseError("unknown StdFont version in FormStreamData")
        reader.skip(2 + 1 + 2 + 4)  # charset, flags, weight, height
        reader.skip(reader.u8())    # face name
    elif tag == _GUID_TEXTPROPS:
        reader.skip(2)              # minor, major
        reader.skip(reader.u16())
    else:
        raise FormParseError(f"unknown font GUID {tag:#010x} in FormStreamData")
    end = reader.pos
    reader.pos = start
    return reader.take(end - start)


def _read_site(reader: _Reader, encoding: str) -> _Site:
    """One OleSiteConcreteControl ([MS-OFORMS] 2.2.10.12.3).

    ``cbSite`` counts from the mask, not from the start of the record, so
    the next site begins at ``start + 4 + cbSite``.
    """
    start = reader.pos
    if reader.u16() != 0x0000:
        raise FormParseError("unsupported OleSiteConcreteControl version")
    cb_site = reader.u16()
    site = _Site(mask=reader.u32())

    for bit, name, size in _SITE_FIELDS:
        if not site.mask & (1 << bit):
            continue
        gap = reader.align(start, size)
        if gap:
            site.pads[f"before:{name}"] = gap
        signed = name in _SIGNED_SITE_FIELDS
        site.values[name] = (
            reader.i32() if size == 4 and signed
            else reader.u32() if size == 4
            else reader.i16() if signed
            else reader.u16()
        )
    gap = reader.align(start, 4)
    if gap:
        site.pads["data:end"] = gap

    def read_string(len_field: str, name: str) -> None:
        packed = site.values.get(len_field, 0)
        raw = reader.take(packed & 0x7FFFFFFF)
        compressed = bool(packed & 0x80000000)
        site.strings[name] = StoredString(
            raw.decode(encoding if compressed else "utf-16-le", "replace"),
            compressed,
            raw,
        )
        pad = reader.align(start, 4)
        if pad:
            site.pads[f"str:{name}"] = pad

    # The extra block opens with the Name.  It is not padded before the
    # first string: aligning here shifts every name.
    for len_field, name in _SITE_STRINGS:
        if len_field in site.values:
            read_string(len_field, name)
    if site.mask & (1 << 8):
        site.position = (reader.i32(), reader.i32())
    for len_field, name in _SITE_STRINGS_AFTER_POSITION:
        if len_field in site.values:
            read_string(len_field, name)

    end = start + 4 + cb_site
    if reader.pos != end:
        raise FormParseError(
            f"site {site.name!r} has {end - reader.pos} unmodelled bytes"
        )
    return site


def _is_trailing_record(data: bytes, offset: int) -> bool:
    """True when the bytes left after the sites are one trailing record.

    A MultiPage's ``f`` carries a MultiPage record after the FormControl
    ([MS-OFORMS] 2.2.4).  It is version-stamped and length-prefixed, so
    this checks rather than merely tolerating a remainder.
    """
    if offset + 4 > len(data):
        return False
    version = (data[offset], data[offset + 1])
    length = int(struct.unpack_from("<H", data, offset + 2)[0])
    return version == _CONTROL_RECORD_VERSION and offset + 4 + length == len(data)


def _try_site_data(
    data: bytes, start: int, encoding: str, class_table_stored: bool
) -> tuple[list[_Site], bytes, bytes, bytes] | None:
    """Read FormSiteData, or ``None`` when this layout is not the real one.

    Returns the sites plus the three runs replayed verbatim on write: the
    class table, the SiteDepthsAndTypes block, and whatever trails the
    sites.  Whether the class-table count is stored depends on a flag
    inside the DataBlock, so both layouts are tried and the one whose
    counts prove out wins.
    """
    try:
        reader = _Reader(data)
        reader.pos = start
        if class_table_stored:
            for _ in range(reader.u16()):
                if reader.u16() != 0x0000:
                    return None
                reader.skip(reader.u16())
        class_table = data[start:reader.pos]
        count_of_sites = reader.u32()
        count_of_bytes = reader.u32()
        if count_of_sites > 10_000:
            return None
        end = reader.pos + count_of_bytes
        # SiteData closes the stream, so an exact match is the cheap proof
        # that this layout is the real one.
        if end != len(data) and not _is_trailing_record(data, end):
            return None

        # SiteDepthsAndTypes: one entry per site, or one counted entry per
        # run of consecutive sites sharing a depth and type.
        depths_start = reader.pos
        accounted = 0
        while accounted < count_of_sites:
            reader.u8()  # depth
            type_or_count = reader.u8()
            if type_or_count & 0x80:
                accounted += type_or_count & 0x7F
                reader.u8()  # OptionalType
            else:
                accounted += 1
        if accounted != count_of_sites:
            return None
        reader.align(depths_start, 4)
        depths = data[depths_start:reader.pos]
        sites = [_read_site(reader, encoding) for _ in range(count_of_sites)]
        if reader.pos != end:
            return None
        return sites, class_table, depths, data[end:]
    except FormParseError:
        return None


def _read_container(data: bytes, encoding: str) -> tuple[_FormStream, int]:
    """Parse a FormControl stream into its record, its blobs and its sites.

    Used for the form itself and for every container control, whose child
    storage carries a FormControl of its own.  The second value is the
    length of the FormControl block, which a resize has to account for.
    """
    reader = _Reader(data)
    version = (reader.u8(), reader.u8())
    if version != _FORM_CONTROL_VERSION:
        raise FormParseError(
            f"not a FormControl stream (version {version[1]}.{version[0]})"
        )
    cb_form = reader.u16()
    mask = int(struct.unpack_from("<I", data, 4)[0])
    record = parse_record(data, FORM_SPEC, encoding, end=4 + cb_form)
    # The PropMask is counted inside cbForm; skip the rest of the block.
    reader.pos = 4 + cb_form
    # FormStreamData: variable-length blobs sit between the form's own data
    # and its sites, so the site array does not start at 4 + cbForm.
    mouse_icon = _read_picture_blob(reader) if mask & _FORM_MASK_MOUSE_ICON else b""
    font = _read_font_blob(reader) if mask & _FORM_MASK_FONT else b""
    picture = _read_picture_blob(reader) if mask & _FORM_MASK_PICTURE else b""

    stored = True
    parsed = _try_site_data(data, reader.pos, encoding, True)
    if parsed is None:
        stored = False
        parsed = _try_site_data(data, reader.pos, encoding, False)
    if parsed is None:
        raise FormParseError("FormSiteData did not reconcile")
    sites, class_table, depths, trailing = parsed
    return (
        _FormStream(
            record=record,
            sites=sites,
            mouse_icon_raw=mouse_icon,
            font_raw=font,
            picture_raw=picture,
            class_table_present=stored,
            class_table_raw=class_table,
            depths_raw=depths,
            trailing_raw=trailing,
        ),
        4 + cb_form,
    )


# ---------------------------------------------------------------------------
# Typing a control
# ---------------------------------------------------------------------------

def _kind_of(site: _Site, record: ParsedRecord | None) -> str:
    if site.clsid_cache_index >= _CLASS_TABLE_BASE:
        return _ACTIVEX_CONTROL
    named = CONTROL_CLASSES.get(site.clsid_cache_index)
    if named is None:
        # An index this table does not know is a structure this reader
        # does not fully understand; the honest claim is the base surface.
        return _UNKNOWN_CONTROL
    if site.clsid_cache_index == 15 and record is not None:
        # A generic MorphData says what it is through DisplayStyle; not
        # stored means the file-format default, which is Text.
        style = record.values.get("DisplayStyle", 1)
        return MORPH_DISPLAY_STYLES.get(style, _UNKNOWN_CONTROL)
    return named




# ---------------------------------------------------------------------------
# Adding and removing controls
# ---------------------------------------------------------------------------

# Points to HIMETRIC: MSForms stores geometry in hundredths of a
# millimetre, and a point is 1/72 inch.
_HIMETRIC_PER_POINT = 2540 / 72


def points_to_himetric(points: float) -> int:
    """Convert a designer measurement in points to the stored unit."""
    return round(points * _HIMETRIC_PER_POINT)


def himetric_to_points(himetric: int) -> float:
    """Convert a stored measurement back to points."""
    return himetric / _HIMETRIC_PER_POINT


# The site mask a freshly added control carries: Name, ID, ObjectStreamSize,
# TabIndex, ClsidCacheIndex, and the position bit.  This is the set Excel
# itself writes for a plain control, and each one is needed: without ID the
# control cannot own a storage, without ObjectStreamSize its record cannot
# be found, and without the position bit it lands at the origin.
_NEW_SITE_MASK = (1 << 0) | (1 << 2) | (1 << 5) | (1 << 6) | (1 << 7) | (1 << 8)

# A container's site instead carries BitFlags and no ObjectStreamSize.
_NEW_CONTAINER_SITE_MASK = (
    (1 << 0) | (1 << 2) | (1 << 4) | (1 << 6) | (1 << 7) | (1 << 8)
)
# Visible and enabled, as Excel writes for a Frame -- taken from the
# fixture's own containers rather than assembled from flag names.
_CONTAINER_SITE_BITFLAGS = 262179
# A page that is not the selected one, which is the only way a new page's
# BitFlags differ from the first page's.
_INACTIVE_PAGE_BITFLAGS = 262177

# The TabStrip's parallel arrays, each with the DataBlock field holding
# its size.  They stay the same length as each other and as the tab count.
_TAB_ARRAYS: tuple[tuple[str, str], ...] = (
    ("Items", "ItemsSize"),
    ("TipStrings", "TipStringsSize"),
    ("TabNames", "NamesSize"),
    ("Tags", "TagsSize"),
    ("Accelerators", "AcceleratorsSize"),
)


def _next_tab_name(tabs: ParsedRecord) -> str:
    """The internal tab name Excel would use next: Tab3, Tab4, Tab5...

    A counter that continues past the highest existing suffix rather than
    tracking the page count, which is what the fixture's Tab3/Tab4 on a
    two-page MultiPage shows.
    """
    highest = 0
    for entry in parse_string_array(tabs.arrays.get("TabNames", b""), "latin-1"):
        digits = entry.text[3:]
        if entry.text[:3].casefold() == "tab" and digits.isdigit():
            highest = max(highest, int(digits))
    return f"Tab{highest + 1}"

# Default sizes in points, as the VBE uses when you drop a control.
_DEFAULT_SIZE_PT: dict[str, tuple[float, float]] = {
    "CommandButton": (72, 24),
    "Label": (72, 18),
    "TextBox": (72, 18),
    "ComboBox": (72, 18),
    "ListBox": (72, 54),
    "CheckBox": (108, 18),
    "OptionButton": (108, 18),
    "ToggleButton": (72, 24),
    "Image": (72, 72),
    "SpinButton": (13, 24),
    "ScrollBar": (13, 72),
}
_DEFAULT_SIZE_FALLBACK = (72, 24)

# MorphData's PropMask bit 31 is reserved and MUST be 1
# ([MS-OFORMS] 2.2.5.2).
_MORPH_RESERVED_BIT = 31

# What Excel itself sets on a fresh control of each MorphData kind, beyond
# the size and display style every one carries.  A two-state control's
# Value is a string, which is why these are typed rather than all numeric.
_TWO_STATE_DEFAULTS: tuple[tuple[str, int | str], ...] = (
    ("BackColor", 0x8000000F),
    ("ForeColor", 0x80000012),
    ("Value", "0"),
)
_MORPH_DEFAULTS: dict[str, tuple[tuple[str, int | str], ...]] = {
    "TextBox": (("VariousPropertyBits", 0x2C80481B),),
    "ComboBox": (
        ("VariousPropertyBits", 0x2C80481B),
        ("MatchEntry", 1),
        ("ShowDropButtonWhen", 2),
    ),
    "ListBox": (("ScrollBars", 3), ("MatchEntry", 0)),
    "CheckBox": _TWO_STATE_DEFAULTS,
    "OptionButton": _TWO_STATE_DEFAULTS,
    "ToggleButton": _TWO_STATE_DEFAULTS,
}

# Controls whose caption defaults to their name, as the VBE does.
_CAPTIONED = frozenset(
    {"Label", "CommandButton", "ToggleButton", "CheckBox", "OptionButton"}
)

_CACHE_INDEX_BY_KIND = {
    name.split(".", 1)[1]: index for index, name in CONTROL_CLASSES.items()
}


def _new_record(kind: str, cache_index: int, name: str) -> ParsedRecord:
    """The record a newly added control starts with.

    Deliberately close to what Excel itself writes.  MSForms stores only
    what differs from a control's default, so a new control sets little --
    but the few things it must set, it must: a record missing them makes
    Office refuse the whole form rather than fall back to a default.
    """
    spec = SPECS_BY_CACHE_INDEX.get(cache_index)
    if spec is None:
        raise FormParseError(f"no property table for {kind}")
    record = ParsedRecord(spec, 0)
    # fSize is set on every control record Excel writes.
    width, height = _DEFAULT_SIZE_PT.get(kind, _DEFAULT_SIZE_FALLBACK)
    record.set_size(points_to_himetric(width), points_to_himetric(height))
    if kind in _CAPTIONED:
        record.set_string("Caption", name)

    if spec.mask64:
        # MorphData's mask bit 31 is reserved and MUST be 1
        # ([MS-OFORMS] 2.2.5.2).  Measured: an OptionButton written
        # without it makes Excel refuse the form outright, and setting it
        # is the single change that makes the same control load.
        record.mask |= 1 << _MORPH_RESERVED_BIT
        # A MorphData is only typed by its DisplayStyle, so a new one has
        # to say what it is or it reads back as a TextBox.
        style = next(
            (s for s, n in MORPH_DISPLAY_STYLES.items() if n == f"MSForms.{kind}"),
            None,
        )
        if style is None:
            raise FormParseError(f"{kind} has no fmDisplayStyle")
        if style != 1:
            record.set_value("DisplayStyle", style)
        for name_, value in _MORPH_DEFAULTS.get(kind, ()):
            if isinstance(value, str):
                record.set_string(name_, value)
            else:
                record.set_value(name_, value)
    elif kind in ("SpinButton", "ScrollBar"):
        record.set_value("Orientation", -1)

    if spec.text_props:
        # Every record whose class carries TextProps must hold one: the
        # reader expects it right after the StreamData.  Excel writes
        # Tahoma 8.25pt (165 twips) everywhere, and centres button text.
        text_props = ParsedRecord(TEXT_PROPS_SPEC, 0)
        text_props.set_string("FontName", "Tahoma")
        text_props.set_value("FontHeight", 165)
        text_props.set_value("FontCharSet", 0)
        text_props.set_value("FontPitchAndFamily", 2)
        if kind in ("CommandButton", "ToggleButton"):
            text_props.set_value("ParagraphAlign", 3)
        record.text_props = text_props
    return record


def _adopt_font(record: ParsedRecord, level: _Level) -> None:
    """Give a new control the font its siblings already carry.

    A control's font lives in its own TextProps, and MSForms will not
    load a record whose class declares TextProps but stores an empty one.
    Copying a sibling's is also what the control would have inherited, so
    the new control looks like the ones beside it rather than like a
    default this module invented.
    """
    if record.text_props is None:
        return
    for other in level.controls:
        if other.record is not None and other.record.text_props is not None:
            record.text_props = deepcopy(other.record.text_props)
            return


def _new_site(name: str, site_id: int, cache_index: int, tab_index: int,
              left: int, top: int, *, encoding: str,
              container: bool = False) -> _Site:
    # A container's site carries BitFlags and no ObjectStreamSize: its
    # record is the `f` of its own storage, not a slice of the parent's `o`.
    site = _Site(mask=_NEW_CONTAINER_SITE_MASK if container else _NEW_SITE_MASK)
    site.values["ID"] = site_id
    if container:
        site.values["BitFlags"] = _CONTAINER_SITE_BITFLAGS
    else:
        site.values["ObjectStreamSize"] = 0
    site.values["TabIndex"] = tab_index
    site.values["ClsidCacheIndex"] = cache_index
    site.position = (left, top)
    compressed = all(ord(c) <= 0xFF for c in name)
    site.strings["Name"] = StoredString(name, compressed, b"", edited=True)
    site.values["NameData"] = (
        len(name.encode(encoding if compressed else "utf-16-le", "replace"))
        | (0x80000000 if compressed else 0)
    )
    return site



# ---------------------------------------------------------------------------
# Container storages
# ---------------------------------------------------------------------------
#
# A container control keeps its own record and its children in a storage of
# its own, named for its site id.  The storage is bound by its CLSID and by
# the CompObj naming the kind fm20 should treat it as; get either wrong and
# the container loads without erroring and simply does not appear.  Both are
# reproduced verbatim from an Excel-authored fixture.

_FRAME_CLSID = bytes.fromhex("2020186e60f4ce119bcd00aa00608e01")
_PAGE_CLSID = bytes.fromhex("f0692ac6dc16ce119e9800aa00574a4f")
_MULTIPAGE_CLSID = bytes.fromhex("7013e3467a3fce11bed600aa00611080")

_FRAME_COMPOBJ = bytes.fromhex(
    "0100feff030a0000ffffffff2020186e60f4ce119bcd00aa00608e01"
    "1a0000004d6963726f736f667420466f726d7320322e30204672616d"
    "650010000000456d626564646564204f626a656374000e000000466f"
    "726d732e4672616d652e3100f439b271000000000000000000000000"
)
_PAGE_COMPOBJ = bytes.fromhex(
    "0100feff030a0000fffffffff0692ac6dc16ce119e9800aa00574a4f"
    "190000004d6963726f736f667420466f726d7320322e3020466f726d"
    "0010000000456d626564646564204f626a656374000d000000466f72"
    "6d732e466f726d2e3100f439b271000000000000000000000000"
)

_MULTIPAGE_COMPOBJ = bytes.fromhex(
    "0100feff030a0000ffffffff7013e3467a3fce11bed600aa00611080"
    "190000004d6963726f736f667420466f726d7320322e3020466f726d"
    "0010000000456d626564646564204f626a6563740012000000466f72"
    "6d732e4d756c7469506167652e3100f439b271000000000000000000000000"
)

_COMPOBJ_STREAM = "\x01CompObj"

# A MultiPage's `f` ends with a MultiPage record after its sites.  Excel
# writes the same 16 bytes for every one this has seen, and their two
# payload words have no meaning this can state, so they are reproduced
# verbatim rather than invented.
_MULTIPAGE_TRAILER = bytes.fromhex("0002 0c00 19000000 fc8f0000 ff010000")

# BooleanProperties as Excel writes them: 0x8004 for a Frame or a Page,
# 0xC004 for a MultiPage.
_FRAME_BOOLEAN_PROPERTIES = 0x00008004
_MULTIPAGE_BOOLEAN_PROPERTIES = 0x0000C004

# A page's own client area, as Excel sizes the two it creates with a
# MultiPage.  HIMETRIC.
_DEFAULT_PAGE_SIZE = (5080, 3810)
# Where Excel sites a page inside its MultiPage's client area.  HIMETRIC.
_DEFAULT_PAGE_ORIGIN = (53, 556)

# The mask bit a fresh TabStrip carries that the record table does not
# name.  Excel sets it on every TabStrip this has seen; the table stops at
# the fields it can name, so this is set explicitly rather than silently
# folded into one of them.
_TABSTRIP_UNNAMED_BIT = 19

# The mask bits an empty container's own FormControl sets.  Measured from
# the fixture's Frames and Pages: a page authored without BooleanProperties,
# DrawBuffer and LogicalSize was silently not bound.
_CONTAINER_DRAW_BUFFER = 32000


def _new_container_stream(
    kind: str, caption: str | None, width: int, height: int
) -> _FormStream:
    """The minimal FormControl an Excel-authored container carries."""
    record = ParsedRecord(FORM_SPEC, 0)
    record.set_value("NextAvailableID", 1)
    record.set_value(
        "BooleanProperties",
        _MULTIPAGE_BOOLEAN_PROPERTIES
        if kind == "MultiPage"
        else _FRAME_BOOLEAN_PROPERTIES,
    )
    record.set_value("DrawBuffer", _CONTAINER_DRAW_BUFFER)
    record.sizes["DisplayedSize"] = Size(width, height)
    record.mask |= 1 << 10
    record.sizes["LogicalSize"] = Size(0, 0)
    record.mask |= 1 << 11
    if kind == "Frame" and caption is not None:
        record.set_string("Caption", caption)
    return _FormStream(
        record=record,
        sites=[],
        class_table_present=False,
        sites_structurally_changed=True,
        trailing_raw=_MULTIPAGE_TRAILER if kind == "MultiPage" else b"",
    )


def _new_tabstrip_record() -> ParsedRecord:
    """The TabStrip a MultiPage owns, with no tabs yet.

    Not a control the caller ever names: MSForms sites it ahead of the
    pages and it holds the whole of the MultiPage's `o`, so a MultiPage
    without one is not a MultiPage.
    """
    record = ParsedRecord(SPECS_BY_CACHE_INDEX[18], 0)
    record.mask |= 1 << _TABSTRIP_UNNAMED_BIT
    record.set_value("ListIndex", 0)
    record.set_size(*_DEFAULT_PAGE_SIZE)
    for _, size_field in _TAB_ARRAYS:
        record.set_value(size_field, 0)
    record.set_value("TabsAllocated", TAB_HEADROOM)
    record.set_value("TabData", 0)
    for name, _ in _TAB_ARRAYS:
        record.arrays[name] = b""
    text_props = ParsedRecord(TEXT_PROPS_SPEC, 0)
    text_props.set_string("FontName", "Tahoma")
    text_props.set_value("FontHeight", 165)
    text_props.set_value("FontCharSet", 0)
    text_props.set_value("FontPitchAndFamily", 2)
    record.text_props = text_props
    return record


def _storage_name(site_id: int) -> str:
    """The ``iNN`` name a container site's storage takes.

    Zero-padded to two digits, which is what MSForms writes; reading
    matches on the number instead, so only writing depends on this.
    """
    return f"i{site_id:02d}"

def _missing_page(name: str, multipage: str) -> FormControl:
    """Raise for a page a MultiPage does not have; typed to satisfy `or`."""
    raise KeyError(f"{multipage!r} has no page named {name!r}")


def _storage_id(storage: str) -> int | None:
    """The site id a container storage name carries, or ``None``.

    Matching on the id beats rebuilding the name from it: the file says
    which storages exist, and that name's zero padding is only observed.
    """
    digits = storage[1:]
    if storage[:1].casefold() == "i" and digits.isdigit():
        return int(digits)
    return None


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def form_names(cfb: CFB) -> list[str]:
    """Names of the designer storages in a project's CFB.

    A form's storage sits at the root beside ``VBA/`` and holds an ``f``
    stream; that is the structural test, because the dir stream types a
    designer and a class module identically.
    """
    names: list[str] = []
    for storage in cfb.list_storages_at():
        if storage.casefold() == "vba":
            continue
        if "f" in {name.casefold() for name in cfb.list_streams_at([storage])}:
            names.append(storage)
    return names


def read_form(cfb: CFB, name: str, *, code_page: int = 1252) -> VBAForm:
    """Read one form's designer streams.

    Raises :class:`FormParseError` when the streams do not reconcile,
    rather than returning a guessed control list.
    """
    from pyopenvba.vba import encoding_for_codepage

    encoding = encoding_for_codepage(code_page)
    try:
        raw_designer = cfb.get_stream_at([name], _VBFRAME_STREAM)
    except KeyError:
        raw_designer = b""
    levels: list[_Level] = []
    controls = _read_level(cfb, [name], encoding, levels)
    return VBAForm(
        name=name,
        designer_source=raw_designer.decode(encoding, "replace"),
        controls=controls,
        _levels=levels,
        _encoding=encoding,
    )


def read_forms(cfb: CFB, *, code_page: int = 1252) -> list[VBAForm]:
    """Read every form in a project's CFB, in directory order."""
    return [read_form(cfb, name, code_page=code_page) for name in form_names(cfb)]


def _read_level(
    cfb: CFB, path: list[str], encoding: str, levels: list[_Level]
) -> tuple[FormControl, ...]:
    """Parse one container's ``f``/``o`` pair and recurse into its children."""
    where = "/".join(path)
    try:
        f_stream = cfb.get_stream_at(path, "f")
    except KeyError:
        raise FormParseError(f"{where}: no 'f' stream") from None
    try:
        o_stream = cfb.get_stream_at(path, "o")
    except KeyError:
        o_stream = b""

    stream, _ = _read_container(f_stream, encoding)
    sites = stream.sites

    total = sum(site.object_stream_size for site in sites)
    if total != len(o_stream):
        # The strongest check available: the sites describe the whole of
        # `o` or the sites were read wrong.
        raise FormParseError(
            f"{where}: sites account for {total} bytes of 'o' but it holds "
            f"{len(o_stream)}"
        )

    # A container's own record is the `f` of its child storage, named for
    # the site id.  Matching on the id the storage name carries beats
    # rebuilding the name, whose padding this has only observed.
    substorages: dict[int, str] = {}
    for storage in cfb.list_storages_at(path):
        digits = storage[1:]
        if storage[:1].casefold() == "i" and digits.isdigit():
            substorages[int(digits)] = storage

    try:
        x_stream = cfb.get_stream_at(path, "x")
    except KeyError:
        x_stream = b""
    level = _Level(
        path=list(path),
        f_raw=f_stream,
        o_raw=o_stream,
        stream=stream,
        controls=[],
        x_raw=x_stream,
    )
    levels.append(level)

    offset = 0
    for site in sites:
        slice_bytes = o_stream[offset:offset + site.object_stream_size]
        offset += site.object_stream_size
        storage = substorages.pop(site.id, None)
        if storage is not None:
            children = _read_level(cfb, [*path, storage], encoding, levels)
            # The child level parsed the container's own record; find it by
            # path, since recursion appends grandchildren after it.
            record = next(
                lvl.record for lvl in levels if lvl.path == [*path, storage]
            )
        else:
            children = ()
            spec = SPECS_BY_CACHE_INDEX.get(site.clsid_cache_index)
            record = (
                parse_record(slice_bytes, spec, encoding)
                if spec is not None and slice_bytes
                else None
            )
        control = FormControl(
            name=site.name,
            kind=_kind_of(site, record),
            id=site.id,
            clsid_cache_index=site.clsid_cache_index,
            tab_index=site.tab_index,
            object_stream_size=site.object_stream_size,
            record=record,
            children=children,
        )
        level.controls.append(control)
    if substorages:
        # A child storage no site claims means the sites were misread, or
        # this is a layout the reader does not understand.  Either way the
        # control list would be incomplete.
        orphans = ", ".join(sorted(substorages.values()))
        raise FormParseError(f"{where}: child storages claimed by no site: {orphans}")
    return tuple(level.controls)
