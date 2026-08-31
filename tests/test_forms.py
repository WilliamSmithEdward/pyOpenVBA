"""Tests for the UserForm designer reader (GitHub issue #15)."""

from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

import pytest

from pyopenvba._oforms_records import Size
from pyopenvba.cfb import CFB
from pyopenvba.excel import ExcelFile
from pyopenvba.exceptions import FormParseError
from pyopenvba.forms import (
    FormControl,
    VBAForm,
    form_names,
    himetric_to_points,
    read_form,
    read_forms,
)
from pyopenvba.powerpoint import PowerPointFile

_HERE = Path(__file__).parent
_NESTED = _HERE / "live_excel_testing" / "nested_form.xlsm"
_FLAT_XLSM = _HERE / "live_excel_testing" / "test_macro_workbook.xlsm"
_FLAT_PPTM = _HERE / "live_powerpoint_testing" / "Presentation1.pptm"

_VBA_ENTRY = "xl/vbaProject.bin"


def _project_cfb(path: Path, entry: str = _VBA_ENTRY) -> CFB:
    with zipfile.ZipFile(path) as archive:
        return CFB.from_bytes(archive.read(entry))


@pytest.fixture(scope="module")
def nested() -> VBAForm:
    """The nested fixture, parsed once for the tree tests."""
    return read_form(_project_cfb(_NESTED), "FrmNested")


def _by_name(controls: list[FormControl], name: str) -> FormControl:
    for control in controls:
        if control.name == name:
            return control
    raise AssertionError(f"no control named {name!r} in {[c.name for c in controls]}")


# ---------------------------------------------------------------------------
# Discovery
# ---------------------------------------------------------------------------

@pytest.mark.skipif(not _NESTED.exists(), reason="nested form fixture not present")
class TestFormDiscovery:
    def test_form_names_finds_the_designer_storage(self) -> None:
        assert form_names(_project_cfb(_NESTED)) == ["FrmNested"]

    def test_vba_storage_is_not_mistaken_for_a_form(self) -> None:
        """`VBA/` sits at the root beside the forms and must be skipped."""
        assert "VBA" not in form_names(_project_cfb(_NESTED))

    def test_a_project_with_one_empty_form_still_reports_it(self) -> None:
        assert form_names(_project_cfb(_FLAT_XLSM)) == ["UserForm1"]


# ---------------------------------------------------------------------------
# The control tree
# ---------------------------------------------------------------------------

@pytest.mark.skipif(not _NESTED.exists(), reason="nested form fixture not present")
class TestNestedFormTree:
    """The fixture is one Excel-authored form covering every shape the
    reader distinguishes: a leaf, a Frame with children, a MultiPage with
    two Pages, and a leaf after the containers."""

    def test_top_level_controls_in_site_order(self, nested: VBAForm) -> None:
        assert [c.name for c in nested.controls] == [
            "TopLabel", "GroupBox", "Pages", "Picker", "CloseButton",
        ]

    def test_frame_children_live_in_their_own_storage(self, nested: VBAForm) -> None:
        frame = _by_name(list(nested.controls), "GroupBox")
        assert frame.kind == "MSForms.Frame"
        assert [c.name for c in frame.children] == ["OptOne", "OptTwo", "InnerText"]

    def test_multipage_owns_a_hidden_tabstrip_and_its_pages(self, nested: VBAForm) -> None:
        """MSForms sites an unnamed TabStrip beside the Pages; a reader
        that dropped it would misalign every following `o` slice."""
        pages = _by_name(list(nested.controls), "Pages")
        assert pages.kind == "MSForms.MultiPage"
        assert [(c.name, c.kind) for c in pages.children] == [
            ("", "MSForms.TabStrip"),
            ("Page1", "MSForms.Form"),
            ("Page2", "MSForms.Form"),
        ]

    def test_pages_carry_their_own_controls(self, nested: VBAForm) -> None:
        pages = _by_name(list(nested.controls), "Pages")
        page_one = _by_name(list(pages.children), "Page1")
        page_two = _by_name(list(pages.children), "Page2")
        assert [c.name for c in page_one.children] == ["PageOneCheck"]
        assert [c.name for c in page_two.children] == ["PageTwoButton"]

    def test_control_kinds(self, nested: VBAForm) -> None:
        kinds = {c.name: c.kind for c in nested.walk() if c.name}
        assert kinds == {
            "TopLabel": "MSForms.Label",
            "GroupBox": "MSForms.Frame",
            "OptOne": "MSForms.OptionButton",
            "OptTwo": "MSForms.OptionButton",
            "InnerText": "MSForms.TextBox",
            "Pages": "MSForms.MultiPage",
            "Page1": "MSForms.Form",
            "PageOneCheck": "MSForms.CheckBox",
            "Page2": "MSForms.Form",
            "PageTwoButton": "MSForms.CommandButton",
            "Picker": "MSForms.ListBox",
            "CloseButton": "MSForms.CommandButton",
        }

    def test_walk_is_depth_first_containers_before_children(self, nested: VBAForm) -> None:
        walked = [c.name for c in nested.walk()]
        assert walked.index("GroupBox") < walked.index("OptOne")
        assert walked.index("Pages") < walked.index("Page1")
        assert walked.index("Page1") < walked.index("PageOneCheck")
        # And the trailing leaf comes after the containers it follows.
        assert walked.index("PageTwoButton") < walked.index("CloseButton")

    def test_site_ids_are_unique(self, nested: VBAForm) -> None:
        """Ids name the child storages, so a duplicate would mean two
        controls claiming the same one."""
        ids = [c.id for c in nested.walk()]
        assert len(ids) == len(set(ids))

    def test_containers_hold_no_slice_of_o(self, nested: VBAForm) -> None:
        for control in nested.walk():
            if control.children:
                assert control.object_stream_size == 0

    def test_morph_controls_carry_the_wider_mask(self, nested: VBAForm) -> None:
        """A bit index only means something with the width beside it."""
        widths = {c.name: c.property_mask_width for c in nested.walk() if c.name}
        assert widths["InnerText"] == 8      # TextBox is a MorphData
        assert widths["PageOneCheck"] == 8   # so is CheckBox
        assert widths["Picker"] == 8         # and ListBox
        assert widths["TopLabel"] == 4       # Label is not
        assert widths["CloseButton"] == 4

    def test_optional_site_fields_do_not_shift_the_name(
        self, nested: VBAForm
    ) -> None:
        """`Picker` stores ControlSource and RowSource and `CloseButton`
        stores a tip.  Each adds a field ahead of the name, so reading any
        of them at the wrong width renames the control that follows."""
        names = {c.name for c in nested.walk()}
        assert {"Picker", "CloseButton", "InnerText"} <= names

    def test_every_control_reports_a_mask(self, nested: VBAForm) -> None:
        assert all(c.properties_set for c in nested.walk())

    def test_designer_source_is_the_vbframe_text(self, nested: VBAForm) -> None:
        assert nested.designer_source.startswith("VERSION 5.00")
        assert "FrmNested" in nested.designer_source


# ---------------------------------------------------------------------------
# Forms with no controls
# ---------------------------------------------------------------------------

class TestEmptyForms:
    @pytest.mark.skipif(not _FLAT_XLSM.exists(), reason="fixture not present")
    def test_form_with_no_controls_reads_as_empty(self) -> None:
        form = read_form(_project_cfb(_FLAT_XLSM), "UserForm1")
        assert form.controls == ()
        assert form.designer_source.startswith("VERSION 5.00")

    @pytest.mark.skipif(not _FLAT_PPTM.exists(), reason="fixture not present")
    def test_powerpoint_forms_read_the_same_way(self) -> None:
        forms = read_forms(_project_cfb(_FLAT_PPTM, "ppt/vbaProject.bin"))
        assert [f.name for f in forms] == ["UserForm1"]


# ---------------------------------------------------------------------------
# Host accessors
# ---------------------------------------------------------------------------

class TestHostFormsAccessor:
    @pytest.mark.skipif(not _NESTED.exists(), reason="fixture not present")
    def test_excel_forms(self) -> None:
        with ExcelFile(_NESTED) as workbook:
            forms = workbook.forms()
        assert [f.name for f in forms] == ["FrmNested"]
        assert len(forms[0].walk()) == 13

    @pytest.mark.skipif(not _FLAT_PPTM.exists(), reason="fixture not present")
    def test_powerpoint_forms(self) -> None:
        with PowerPointFile(_FLAT_PPTM) as presentation:
            forms = presentation.forms()
        assert [f.name for f in forms] == ["UserForm1"]

    @pytest.mark.skipif(not _NESTED.exists(), reason="fixture not present")
    def test_a_form_is_also_a_module(self) -> None:
        """The code-behind and the design are two views of one component."""
        with ExcelFile(_NESTED) as workbook:
            assert "FrmNested" in workbook.module_names()
            assert [f.name for f in workbook.forms()] == ["FrmNested"]


# ---------------------------------------------------------------------------
# Refusing to guess
# ---------------------------------------------------------------------------

@pytest.mark.skipif(not _NESTED.exists(), reason="fixture not present")
class TestFormParseRefusals:
    """A misread site array yields a plausible-looking control list, so
    each structural check has to fail loudly rather than degrade."""

    def test_o_stream_that_does_not_add_up_is_refused(self) -> None:
        cfb = _project_cfb(_NESTED)
        cfb.write_stream_in_storage(
            "FrmNested", "o", cfb.get_stream_at(["FrmNested"], "o") + b"\x00"
        )
        rebuilt = CFB.from_bytes(cfb.to_bytes())
        with pytest.raises(FormParseError, match="bytes of 'o'"):
            read_form(rebuilt, "FrmNested")

    def test_child_storage_no_site_claims_is_refused(self) -> None:
        cfb = _project_cfb(_NESTED)
        cfb.add_substorage("FrmNested", "i99")
        rebuilt = CFB.from_bytes(cfb.to_bytes())
        with pytest.raises(FormParseError, match="claimed by no site"):
            read_form(rebuilt, "FrmNested")

    def test_f_stream_of_the_wrong_version_is_refused(self) -> None:
        cfb = _project_cfb(_NESTED)
        original = bytearray(cfb.get_stream_at(["FrmNested"], "f"))
        original[1] = 0x09  # major version the reader was not written for
        cfb.write_stream_in_storage("FrmNested", "f", bytes(original))
        rebuilt = CFB.from_bytes(cfb.to_bytes())
        with pytest.raises(FormParseError, match="not a FormControl stream"):
            read_form(rebuilt, "FrmNested")

    def test_truncated_f_stream_is_refused(self) -> None:
        cfb = _project_cfb(_NESTED)
        cfb.write_stream_in_storage(
            "FrmNested", "f", cfb.get_stream_at(["FrmNested"], "f")[:24]
        )
        rebuilt = CFB.from_bytes(cfb.to_bytes())
        with pytest.raises(FormParseError):
            read_form(rebuilt, "FrmNested")

    def test_missing_f_stream_is_refused(self) -> None:
        cfb = _project_cfb(_NESTED)
        cfb.remove_stream_in_storage("FrmNested", "f")
        rebuilt = CFB.from_bytes(cfb.to_bytes())
        with pytest.raises(FormParseError, match="no 'f' stream"):
            read_form(rebuilt, "FrmNested")


@pytest.mark.skipif(not _NESTED.exists(), reason="fixture not present")
class TestSiteAlignment:
    """A site's DataBlock only ends off a 4-byte boundary when GroupID is
    stored, and Excel does not emit that for anything this fixture can
    build -- not even option buttons sharing a GroupName.  The alignment
    before the name is still load-bearing for files that do, so it gets a
    site built by hand and read back through the public API.
    """

    @staticmethod
    def _f_stream_with_group_id(name: bytes) -> bytes:
        import struct

        # Site mask: Name, ID, TabIndex, ClsidCacheIndex, GroupID.  The
        # three 2-byte fields are what leave the block at 2 mod 4.
        mask = (1 << 0) | (1 << 2) | (1 << 6) | (1 << 7) | (1 << 9)
        site_start = 22
        block = struct.pack(
            "<HHIIIHHH",
            0x0000,                      # version
            28,                          # cbSite, counted from the mask
            mask,
            0x80000000 | len(name),      # compressed name, length
            1,                           # ID
            0,                           # TabIndex
            17,                          # ClsidCacheIndex: CommandButton
            3,                           # GroupID
        )
        # Two bytes of padding the reader has to step over.  Reading the
        # name without doing so returns these instead of its first half.
        body = block + b"\xff\xff" + name
        body = body.ljust(4 + 28, b"\x00")   # cbSite says where a next site would start
        head = struct.pack("<BBHI", 0x00, 0x04, 4, 0)   # FormControl, PropMask only
        depths = b"\x00\x01" + b"\x00\x00"              # one site, then aligned to 4
        total = site_start + len(body)
        counts = struct.pack("<HII", 0, 1, total - 18)  # class table, sites, cbSites
        stream = head + counts + depths + body
        assert len(stream) == total, (len(stream), total)
        return stream

    def test_name_is_read_past_the_datablock_padding(self) -> None:
        """The Frame's storage is a leaf, so its streams can be replaced
        without disturbing the rest of the tree."""
        cfb = _project_cfb(_NESTED)
        cfb.write_stream_in_storage("i02", "f", self._f_stream_with_group_id(b"Hello"))
        # The synthetic site stores no ObjectStreamSize, so `o` must be
        # empty for the sites to account for it exactly.
        cfb.write_stream_in_storage("i02", "o", b"")
        rebuilt = CFB.from_bytes(cfb.to_bytes())

        frame = _by_name(list(read_form(rebuilt, "FrmNested").controls), "GroupBox")
        assert [c.name for c in frame.children] == ["Hello"]
        assert frame.children[0].kind == "MSForms.CommandButton"


# ---------------------------------------------------------------------------
# Named properties
# ---------------------------------------------------------------------------

@pytest.mark.skipif(not _NESTED.exists(), reason="fixture not present")
class TestNamedProperties:
    """The mask says *which* properties a control sets; the per-class
    tables in ``_oforms_records`` say which ones those are."""

    def test_a_button_reports_its_caption(self, nested: VBAForm) -> None:
        assert nested.control("CloseButton").get("Caption") == "Close"

    def test_a_label_reports_its_caption_and_size(self, nested: VBAForm) -> None:
        label = nested.control("TopLabel")
        assert label.get("Caption") == "Top level"
        assert label.get("Size") == Size(width=2540, height=635)

    def test_font_properties_come_from_the_nested_textprops(
        self, nested: VBAForm
    ) -> None:
        """A control's font is its own record's TextProps, so it is
        namespaced rather than merged into the control's own fields."""
        assert nested.control("TopLabel").get("Font.FontName") == "Tahoma"

    def test_a_morphdata_reports_the_style_that_types_it(
        self, nested: VBAForm
    ) -> None:
        # fmDisplayStyle 5 is OptionButton, which is also how a generic
        # MorphData site gets typed.
        assert nested.control("OptOne").get("DisplayStyle") == 5
        assert nested.control("OptOne").get("GroupName") == "Choice"

    def test_a_container_reports_its_own_record(self, nested: VBAForm) -> None:
        """A Frame keeps its properties in its child storage's ``f``, not
        in a slice of the parent's ``o``."""
        frame = nested.control("GroupBox")
        assert frame.object_stream_size == 0
        assert frame.get("Caption") == "A frame"

    def test_the_form_reports_its_own_properties(self, nested: VBAForm) -> None:
        assert nested.get("Caption") == "Nested Fixture"

    def test_property_names_agree_with_the_mask(self, nested: VBAForm) -> None:
        """Every named property must correspond to a set bit, so the two
        views of the same record cannot drift apart."""
        for control in nested.walk():
            if control.record is None:
                continue
            for name in control.properties():
                stem = name.split(".", 1)[0] if name.startswith("Font.") else name
                if name.startswith("Font."):
                    continue
                assert control.record.has(stem), f"{control.name}.{name}"


# ---------------------------------------------------------------------------
# Writing
# ---------------------------------------------------------------------------

@pytest.mark.skipif(not _NESTED.exists(), reason="fixture not present")
class TestFormWriteBack:
    """Editing a property rewrites that control's record and patches its
    site's ObjectStreamSize.  Everything the tables do not model is
    replayed verbatim, so an unedited form must not move a byte."""

    @staticmethod
    def _designer_streams(path: Path) -> dict[str, bytes]:
        cfb = _project_cfb(path)
        out: dict[str, bytes] = {}

        def walk(where: list[str]) -> None:
            for name in cfb.list_streams_at(where):
                out["/".join([*where, name])] = cfb.get_stream_at(where, name)
            for storage in cfb.list_storages_at(where):
                walk([*where, storage])

        walk(["FrmNested"])
        return out

    def test_a_save_with_no_form_edits_moves_no_byte(self, tmp_path: Path) -> None:
        out = tmp_path / "noop.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            workbook.forms()          # parse them, change nothing
            workbook.save()
        assert self._designer_streams(out) == self._designer_streams(_NESTED)

    def test_a_longer_caption_round_trips(self, tmp_path: Path) -> None:
        out = tmp_path / "longer.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            workbook.forms()[0].control("CloseButton").set_property(
                "Caption", "Dismiss this form"
            )
            workbook.save()
        with ExcelFile(out) as workbook:
            assert workbook.forms()[0].control("CloseButton").get("Caption") == (
                "Dismiss this form"
            )

    def test_a_shorter_caption_round_trips(self, tmp_path: Path) -> None:
        out = tmp_path / "shorter.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            workbook.forms()[0].control("TopLabel").set_property("Caption", "Hi")
            workbook.save()
        with ExcelFile(out) as workbook:
            assert workbook.forms()[0].control("TopLabel").get("Caption") == "Hi"

    def test_a_resized_record_updates_its_site(self, tmp_path: Path) -> None:
        """ObjectStreamSize is what keeps `o` sliceable; if it were not
        updated the next control's record would be read from the wrong
        offset and the whole level would stop reconciling."""
        out = tmp_path / "resized.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            control = workbook.forms()[0].control("CloseButton")
            before = control.object_stream_size
            control.set_property("Caption", "A considerably longer caption")
            workbook.save()
            assert control.object_stream_size > before
        with ExcelFile(out) as workbook:
            # Reopening re-runs the sum-of-sizes check against len(o).
            assert workbook.forms()[0].control("CloseButton").object_stream_size > (
                before
            )

    def test_a_new_property_appears_in_the_mask(self, tmp_path: Path) -> None:
        out = tmp_path / "newprop.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            label = workbook.forms()[0].control("TopLabel")
            assert "ForeColor" not in label.properties()
            before = label.properties_set
            label.set_property("ForeColor", 0x0000FF)
            assert label.properties_set != before
            workbook.save()
        with ExcelFile(out) as workbook:
            assert workbook.forms()[0].control("TopLabel").get("ForeColor") == 0x0000FF

    def test_clearing_a_property_removes_it(self, tmp_path: Path) -> None:
        out = tmp_path / "cleared.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            control = workbook.forms()[0].control("CloseButton")
            control.set_property("Caption", None)
            workbook.save()
        with ExcelFile(out) as workbook:
            assert "Caption" not in workbook.forms()[0].control(
                "CloseButton"
            ).properties()

    def test_editing_a_control_inside_a_frame(self, tmp_path: Path) -> None:
        out = tmp_path / "inframe.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            workbook.forms()[0].control("OptOne").set_property("Caption", "Ground")
            workbook.save()
        with ExcelFile(out) as workbook:
            assert workbook.forms()[0].control("OptOne").get("Caption") == "Ground"

    def test_editing_a_control_on_a_multipage_page(self, tmp_path: Path) -> None:
        out = tmp_path / "onpage.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            workbook.forms()[0].control("PageTwoButton").set_property(
                "Caption", "Second"
            )
            workbook.save()
        with ExcelFile(out) as workbook:
            assert workbook.forms()[0].control("PageTwoButton").get("Caption") == (
                "Second"
            )

    def test_editing_a_containers_own_record(self, tmp_path: Path) -> None:
        """The Frame's record heads its child storage's `f`, ahead of its
        children's sites, so resizing it shifts them."""
        out = tmp_path / "container.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            workbook.forms()[0].control("GroupBox").set_property(
                "Caption", "Shipping options for this order"
            )
            workbook.save()
        with ExcelFile(out) as workbook:
            form = workbook.forms()[0]
            assert form.control("GroupBox").get("Caption") == (
                "Shipping options for this order"
            )
            # Its children must still be readable at their new offsets.
            assert [c.name for c in form.control("GroupBox").children] == [
                "OptOne", "OptTwo", "InnerText",
            ]

    def test_editing_the_forms_own_caption(self, tmp_path: Path) -> None:
        out = tmp_path / "formcap.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            workbook.forms()[0].set_property("Caption", "Renamed")
            workbook.save()
        with ExcelFile(out) as workbook:
            form = workbook.forms()[0]
            assert form.get("Caption") == "Renamed"
            assert len(form.walk()) == 13

    def test_the_code_behind_survives_a_design_edit(self, tmp_path: Path) -> None:
        out = tmp_path / "codebehind.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            before = workbook.get_module("FrmNested")
            workbook.forms()[0].control("CloseButton").set_property("Caption", "X")
            workbook.save()
        with ExcelFile(out) as workbook:
            assert workbook.get_module("FrmNested") == before

    def test_a_caption_too_long_for_the_format_is_refused(
        self, tmp_path: Path
    ) -> None:
        """cb is a u16; letting it wrap would write a record the reader
        then overruns, so this has to fail before any byte lands."""
        out = tmp_path / "toolong.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            workbook.forms()[0].control("CloseButton").set_property(
                "Caption", "x" * 70_000
            )
            with pytest.raises(FormParseError, match="caps a record at 65535"):
                workbook.save()

    def test_setting_an_unknown_property_is_refused(self) -> None:
        form = read_form(_project_cfb(_NESTED), "FrmNested")
        with pytest.raises(FormParseError, match="no numeric field"):
            form.control("CloseButton").set_property("NoSuchProperty", 1)


# ---------------------------------------------------------------------------
# Adding and removing controls
# ---------------------------------------------------------------------------

@pytest.mark.skipif(not _NESTED.exists(), reason="fixture not present")
class TestAddAndRemoveControls:
    """Structural edits rebuild the site array, so the counts, the depths
    block and the id allocation all have to come out right at once."""

    def test_added_control_survives_a_round_trip(self, tmp_path: Path) -> None:
        out = tmp_path / "added.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            workbook.forms()[0].add_control(
                "CommandButton", "Added", left=12, top=250, width=90, height=24
            )
            workbook.save()
        with ExcelFile(out) as workbook:
            control = workbook.forms()[0].control("Added")
        assert control.kind == "MSForms.CommandButton"
        assert control.get("Caption") == "Added"

    def test_a_new_id_never_collides_with_an_existing_one(
        self, tmp_path: Path
    ) -> None:
        """NextAvailableID is the highest id already handed out, not the
        next free one.  Using it as-is repeats the last control's id, and
        MSForms then refuses the whole form."""
        out = tmp_path / "ids.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            form = workbook.forms()[0]
            highest = max(c.id for c in form.walk())
            added = form.add_control("Label", "Added")
            assert added.id > highest
            workbook.save()
        with ExcelFile(out) as workbook:
            ids = [c.id for c in workbook.forms()[0].walk()]
            assert len(ids) == len(set(ids))

    def test_a_morphdata_control_sets_the_reserved_bit(
        self, tmp_path: Path
    ) -> None:
        """[MS-OFORMS] 2.2.5.2 makes MorphData's mask bit 31 reserved and
        MUST be 1.  Measured: without it Excel refuses the form, and
        setting it is the single change that makes the control load."""
        out = tmp_path / "morph.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            control = workbook.forms()[0].add_control("OptionButton", "Added")
            assert control.properties_set & (1 << 31)
            workbook.save()
        with ExcelFile(out) as workbook:
            reread = workbook.forms()[0].control("Added")
        assert reread.properties_set & (1 << 31)
        assert reread.kind == "MSForms.OptionButton"
        # A two-state control also carries the Value Excel gives it.
        assert reread.get("Value") == "0"

    def test_added_control_goes_into_a_container(self, tmp_path: Path) -> None:
        out = tmp_path / "incontainer.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            workbook.forms()[0].add_control(
                "OptionButton", "Added", container="GroupBox"
            )
            workbook.save()
        with ExcelFile(out) as workbook:
            frame = workbook.forms()[0].control("GroupBox")
        assert [c.name for c in frame.children] == [
            "OptOne", "OptTwo", "InnerText", "Added",
        ]

    def test_geometry_is_given_in_points(self, tmp_path: Path) -> None:
        out = tmp_path / "geometry.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            workbook.forms()[0].add_control(
                "Label", "Added", left=12, top=250, width=90, height=24
            )
            workbook.save()
        with ExcelFile(out) as workbook:
            size = workbook.forms()[0].control("Added").get("Size")
        assert isinstance(size, Size)
        assert round(himetric_to_points(size.width)) == 90
        assert round(himetric_to_points(size.height)) == 24

    def test_removing_a_control_drops_its_site_and_its_record(
        self, tmp_path: Path
    ) -> None:
        out = tmp_path / "removed.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            form = workbook.forms()[0]
            before = len(form.walk())
            form.remove_control("Picker")
            workbook.save()
        with ExcelFile(out) as workbook:
            form = workbook.forms()[0]
            names = [c.name for c in form.walk()]
        assert "Picker" not in names
        assert len(names) == before - 1

    def test_add_then_remove_leaves_the_rest_alone(self, tmp_path: Path) -> None:
        out = tmp_path / "both.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            form = workbook.forms()[0]
            form.add_control("TextBox", "Added", left=12, top=250)
            form.remove_control("Picker")
            workbook.save()
        with ExcelFile(out) as workbook:
            form = workbook.forms()[0]
            assert form.control("CloseButton").get("Caption") == "Close"
            assert [c.name for c in form.control("GroupBox").children] == [
                "OptOne", "OptTwo", "InnerText",
            ]
            assert form.control("Added").kind == "MSForms.TextBox"

    def test_a_designer_edit_invalidates_the_performance_cache(
        self, tmp_path: Path
    ) -> None:
        """Adding a control changes the form class's members, so leaving
        the cache in place makes Office load a member list that no longer
        matches -- measured: Excel refuses the form outright."""
        out = tmp_path / "cache.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            workbook.forms()[0].add_control("Label", "Added")
            workbook.save()
        cfb = _project_cfb(out)
        cache = cfb.get_stream_in_storage("VBA", "_VBA_PROJECT")
        # The 5-byte header is kept; the body is zeroed ([MS-OVBA] 2.3.4.1).
        assert set(cache[5:]) == {0}

    def test_a_duplicate_name_is_refused(self) -> None:
        form = read_form(_project_cfb(_NESTED), "FrmNested")
        with pytest.raises(FormParseError, match="already has a control named"):
            form.add_control("Label", "CloseButton")

    def test_an_unknown_kind_is_refused(self) -> None:
        form = read_form(_project_cfb(_NESTED), "FrmNested")
        with pytest.raises(FormParseError, match="unknown control kind"):
            form.add_control("Sparkline", "Added")

    def test_a_page_is_added_through_its_multipage(self) -> None:
        """A page belongs to a MultiPage, not to a surface: it is also a
        tab, so the tab arrays and page bookkeeping go with it."""
        form = read_form(_project_cfb(_NESTED), "FrmNested")
        with pytest.raises(FormParseError, match="use add_page"):
            form.add_control("Page", "Added")

    def test_a_page_is_removed_through_remove_page(self) -> None:
        form = read_form(_project_cfb(_NESTED), "FrmNested")
        with pytest.raises(FormParseError, match="use remove_page"):
            form.remove_control("Page1")


@pytest.mark.skipif(not _NESTED.exists(), reason="fixture not present")
class TestFrameStorages:
    """A Frame owns a storage of its own, bound by the storage's CLSID and
    by a CompObj naming what fm20 should treat it as.  Creating one means
    creating that storage; removing one means removing it and everything
    under it, because the next read refuses a form whose child storage no
    site claims."""

    def test_a_new_frame_gets_its_own_storage(self, tmp_path: Path) -> None:
        out = tmp_path / "frame.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            added = workbook.forms()[0].add_control(
                "Frame", "NewFrame", left=12, top=250, width=180, height=70
            )
            site_id = added.id
            workbook.save()
        cfb = _project_cfb(out)
        storage = f"i{site_id:02d}"
        assert storage in cfb.list_storages_at(["FrmNested"])
        assert set(cfb.list_streams_at(["FrmNested", storage])) >= {
            "f", "o", "\x01CompObj",
        }

    def test_a_new_frame_reads_back_as_an_empty_container(
        self, tmp_path: Path
    ) -> None:
        out = tmp_path / "frameread.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            workbook.forms()[0].add_control("Frame", "NewFrame")
            workbook.save()
        with ExcelFile(out) as workbook:
            frame = workbook.forms()[0].control("NewFrame")
        assert frame.kind == "MSForms.Frame"
        assert frame.children == ()
        assert frame.object_stream_size == 0
        assert frame.get("Caption") == "NewFrame"

    def test_controls_can_be_added_into_a_new_frame(self, tmp_path: Path) -> None:
        out = tmp_path / "framechild.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            form = workbook.forms()[0]
            form.add_control("Frame", "NewFrame", left=12, top=250)
            form.add_control("CommandButton", "Inside", container="NewFrame")
            workbook.save()
        with ExcelFile(out) as workbook:
            frame = workbook.forms()[0].control("NewFrame")
        assert [c.name for c in frame.children] == ["Inside"]

    def test_removing_a_frame_takes_its_storage_and_children(
        self, tmp_path: Path
    ) -> None:
        out = tmp_path / "framegone.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            workbook.forms()[0].remove_control("GroupBox")
            workbook.save()
        cfb = _project_cfb(out)
        assert cfb.list_storages_at(["FrmNested"]) == ["i06"]
        with ExcelFile(out) as workbook:
            names = [c.name for c in workbook.forms()[0].walk()]
        assert "GroupBox" not in names
        assert "OptOne" not in names          # its children went with it
        assert "PageOneCheck" in names        # the MultiPage did not

    def test_removing_something_that_is_not_there_raises(self) -> None:
        form = read_form(_project_cfb(_NESTED), "FrmNested")
        with pytest.raises(KeyError):
            form.remove_control("NoSuchControl")

    def test_adding_into_a_container_that_is_not_there_raises(self) -> None:
        form = read_form(_project_cfb(_NESTED), "FrmNested")
        with pytest.raises(KeyError):
            form.add_control("Label", "Added", container="NoSuchFrame")


@pytest.mark.skipif(not _NESTED.exists(), reason="fixture not present")
class TestPages:
    """A page is a container control *and* a tab, so adding or removing
    one moves four structures at once: its site, its storage, an entry in
    each of five TabStrip arrays, and the ``x`` page bookkeeping."""

    def test_a_new_page_appears_on_its_multipage(self, tmp_path: Path) -> None:
        out = tmp_path / "addpage.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            workbook.forms()[0].add_page("Pages")
            workbook.save()
        with ExcelFile(out) as workbook:
            pages = workbook.forms()[0].control("Pages")
        # The hidden TabStrip stays first; the pages follow it.
        assert [c.name for c in pages.children] == ["", "Page1", "Page2", "Page3"]

    def test_a_new_page_gets_its_own_storage(self, tmp_path: Path) -> None:
        out = tmp_path / "pagestorage.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            page = workbook.forms()[0].add_page("Pages", name="Extra")
            page_id = page.id
            workbook.save()
        cfb = _project_cfb(out)
        assert f"i{page_id:02d}" in cfb.list_storages_at(["FrmNested", "i06"])

    def test_the_tab_arrays_stay_the_same_length(self, tmp_path: Path) -> None:
        """The five arrays and the flag tail carry one element per tab;
        letting them drift is exactly what a partial edit would do."""
        from pyopenvba._oforms_pages import (
            parse_page_bookkeeping,
            parse_string_array,
        )

        # Named here rather than imported: pinning them in the test is what
        # catches an array quietly dropping out of the set that has to move.
        tab_arrays = ("Items", "TipStrings", "TabNames", "Tags", "Accelerators")
        out = tmp_path / "arrays.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            workbook.forms()[0].add_page("Pages", name="Extra")
            workbook.save()
        cfb = _project_cfb(out)
        record = read_form(cfb, "FrmNested").control("Pages").children[0].record
        assert record is not None
        lengths = {
            name: len(parse_string_array(record.arrays[name], "cp1252"))
            for name in tab_arrays
        }
        assert set(lengths.values()) == {3}
        assert record.values["TabData"] == 3
        assert len(record.tail_raw) == 3 * 4
        book = parse_page_bookkeeping(cfb.get_stream_at(["FrmNested", "i06"], "x"))
        assert len(book.page_ids) == 3
        # One more PageProperties record than there are pages, the first
        # ignored ([MS-OFORMS] 2.1.2.3).
        assert len(book.page_props) == 4

    def test_a_page_can_be_given_a_name_and_a_caption(self, tmp_path: Path) -> None:
        out = tmp_path / "named.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            workbook.forms()[0].add_page("Pages", name="Summary", caption="Totals")
            workbook.save()
        with ExcelFile(out) as workbook:
            pages = workbook.forms()[0].control("Pages")
        assert "Summary" in [c.name for c in pages.children]

    def test_controls_can_be_added_onto_a_new_page(self, tmp_path: Path) -> None:
        out = tmp_path / "onpage.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            form = workbook.forms()[0]
            form.add_page("Pages", name="Extra")
            form.add_control("Label", "OnExtra", container="Extra")
            workbook.save()
        with ExcelFile(out) as workbook:
            extra = workbook.forms()[0].control("Extra")
        assert [c.name for c in extra.children] == ["OnExtra"]

    def test_removing_a_page_takes_its_tab_and_its_storage(
        self, tmp_path: Path
    ) -> None:
        out = tmp_path / "rmpage.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            workbook.forms()[0].remove_page("Page2")
            workbook.save()
        cfb = _project_cfb(out)
        assert cfb.list_storages_at(["FrmNested", "i06"]) == ["i08"]
        form = read_form(cfb, "FrmNested")
        pages = form.control("Pages")
        assert [c.name for c in pages.children] == ["", "Page1"]
        record = pages.children[0].record
        assert record is not None
        assert record.values["TabData"] == 1
        # The page's own control went with it; the other page's did not.
        names = [c.name for c in form.walk()]
        assert "PageTwoButton" not in names
        assert "PageOneCheck" in names

    def test_page_names_are_scoped_to_their_multipage(
        self, tmp_path: Path
    ) -> None:
        """Excel gives a second MultiPage its own Page1 and Page2 while the
        first still has them; pages are not in Designer.Controls at all."""
        out = tmp_path / "scoped.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            form = workbook.forms()[0]
            form.add_control("MultiPage", "Second", left=12, top=250)
            workbook.save()
        with ExcelFile(out) as workbook:
            form = workbook.forms()[0]
            assert [c.name for c in form.control("Second").children] == [
                "", "Page1", "Page2",
            ]
            assert [c.name for c in form.control("Pages").children] == [
                "", "Page1", "Page2",
            ]

    def test_a_duplicate_page_name_on_one_multipage_is_refused(self) -> None:
        form = read_form(_project_cfb(_NESTED), "FrmNested")
        with pytest.raises(FormParseError, match="already has a page named"):
            form.add_page("Pages", name="Page1")

    def test_adding_a_page_to_something_that_is_not_a_multipage(self) -> None:
        form = read_form(_project_cfb(_NESTED), "FrmNested")
        with pytest.raises(FormParseError, match="not a MultiPage"):
            form.add_page("GroupBox")

    def test_removing_something_that_is_not_a_page(self) -> None:
        form = read_form(_project_cfb(_NESTED), "FrmNested")
        with pytest.raises(FormParseError, match="not a page"):
            form.remove_page("CloseButton")


@pytest.mark.skipif(not _NESTED.exists(), reason="fixture not present")
class TestMultiPageFromScratch:
    """A MultiPage is not one structure but five: a container storage, a
    hidden TabStrip that owns the tab arrays, one storage per page, the
    ``x`` bookkeeping, and a trailing record after its sites."""

    def test_a_new_multipage_comes_with_two_pages(self, tmp_path: Path) -> None:
        out = tmp_path / "newmp.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            workbook.forms()[0].add_control(
                "MultiPage", "Fresh", left=12, top=250, width=200, height=90
            )
            workbook.save()
        with ExcelFile(out) as workbook:
            fresh = workbook.forms()[0].control("Fresh")
        assert fresh.kind == "MSForms.MultiPage"
        assert [(c.name, c.kind) for c in fresh.children] == [
            ("", "MSForms.TabStrip"),
            ("Page1", "MSForms.Form"),
            ("Page2", "MSForms.Form"),
        ]

    def test_a_new_multipage_gets_every_stream_it_needs(
        self, tmp_path: Path
    ) -> None:
        out = tmp_path / "mpstreams.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            fresh = workbook.forms()[0].add_control("MultiPage", "Fresh")
            storage = f"i{fresh.id:02d}"
            workbook.save()
        cfb = _project_cfb(out)
        where = ["FrmNested", storage]
        assert set(cfb.list_streams_at(where)) >= {"f", "o", "x", "\x01CompObj"}
        # One storage per page.
        assert len(cfb.list_storages_at(where)) == 2

    def test_a_new_multipage_carries_its_trailing_record(
        self, tmp_path: Path
    ) -> None:
        """A MultiPage's `f` does not end at its sites; without the
        trailer the reader's own CountOfBytes check would reject it."""
        import struct

        out = tmp_path / "mptrailer.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            fresh = workbook.forms()[0].add_control("MultiPage", "Fresh")
            storage = f"i{fresh.id:02d}"
            workbook.save()
        raw = _project_cfb(out).get_stream_at(["FrmNested", storage], "f")
        # The trailer is a version-stamped, length-prefixed record that
        # closes the stream exactly -- checked here on the bytes, since
        # that is what the reader keys off.
        trailer = raw[-16:]
        assert (trailer[0], trailer[1]) == (0x00, 0x02)
        assert 4 + struct.unpack_from("<H", trailer, 2)[0] == len(trailer)

    def test_pages_can_be_added_to_a_new_multipage(self, tmp_path: Path) -> None:
        out = tmp_path / "mpgrow.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            form = workbook.forms()[0]
            form.add_control("MultiPage", "Fresh", left=12, top=250)
            form.add_page("Fresh", name="Third")
            form.add_control("Label", "OnThird", container="Third")
            workbook.save()
        with ExcelFile(out) as workbook:
            form = workbook.forms()[0]
            assert [c.name for c in form.control("Fresh").children] == [
                "", "Page1", "Page2", "Third",
            ]
            assert [c.name for c in form.control("Third").children] == ["OnThird"]

    def test_every_id_stays_unique_across_a_multipage_build(
        self, tmp_path: Path
    ) -> None:
        """Five sites are allocated in one call, and an id that repeats is
        what makes MSForms refuse the form."""
        out = tmp_path / "mpids.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            workbook.forms()[0].add_control("MultiPage", "Fresh")
            workbook.save()
        with ExcelFile(out) as workbook:
            ids = [c.id for c in workbook.forms()[0].walk()]
        assert len(ids) == len(set(ids))
