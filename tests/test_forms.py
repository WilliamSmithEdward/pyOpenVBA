"""Tests for the UserForm designer reader (GitHub issue #15)."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from pyopenvba.cfb import CFB
from pyopenvba.excel import ExcelFile
from pyopenvba.exceptions import FormParseError
from pyopenvba.forms import (
    FormControl,
    VBAForm,
    form_names,
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
