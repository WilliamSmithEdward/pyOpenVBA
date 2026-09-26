"""Where Excel's and Word's designers store a property VBA names, and what setting it changes besides.

tests/fixtures/form_properties.json is what scripts/measure_form_properties.py
saw when each designer set one property per control through VBA and saved:
each control's site, mask and record bytes, which the tests take apart. The
two applications wrote the same records but for the alignment padding, which
holds whatever was in memory. The form's own font was set italic and
underlined first, so every control carries those two effects and the
fAutoColor flag that comes with any effect. docs/userforms.md lists these
as the reference for writing the stored fields.

tests/fixtures/uncoupled_edge.json is what scripts/measure_uncoupled_edge.py
saw Excel and Word do with a TextBox the library gave both a border and its
sunken effect, which the designers never write.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba import ExcelFile
from pyopenvba._oforms_records import SPECS_BY_CACHE_INDEX, ParsedRecord, parse_record

RECORD: dict[str, dict[str, Any]] = json.loads(
    (Path(__file__).parent / "fixtures" / "form_properties.json").read_text(encoding="utf-8"))
EDGE: dict[str, dict[str, Any]] = json.loads(
    (Path(__file__).parent / "fixtures" / "uncoupled_edge.json").read_text(encoding="utf-8"))
HOSTS = list(RECORD)
AUTO_COLOR, DISABLED = 1 << 30, 1 << 13


def _control(host: str, name: str) -> dict[str, Any]:
    return RECORD[host]["controls"][name]


def _record(host: str, name: str) -> ParsedRecord:
    """The designer's record for the control, taken apart by the reader."""
    control = _control(host, name)
    spec = SPECS_BY_CACHE_INDEX[control["site"]["ClsidCacheIndex"]]
    return parse_record(bytes.fromhex(control["record"]), spec, "cp1252")


def _props(host: str, name: str) -> dict[str, object]:
    return _record(host, name).properties()


def _effects(host: str, name: str) -> int:
    effects = _props(host, name).get("Font.FontEffects", 0)
    assert isinstance(effects, int)
    return effects


@pytest.mark.parametrize("host", HOSTS)
def test_text_align_is_the_paragraph_align_with_center_and_right_swapped(host: str) -> None:
    names = ("AlignLeft", "AlignCenter", "AlignRight")
    aligns = {name: _props(host, name).get("Font.ParagraphAlign") for name in names}
    assert aligns == {"AlignLeft": None, "AlignCenter": 3, "AlignRight": 2}


@pytest.mark.parametrize("host", HOSTS)
def test_alignment_left_is_bit_13_of_the_various_property_bits(host: str) -> None:
    left = _props(host, "BoxLeft")["VariousPropertyBits"]
    assert isinstance(left, int) and left & 1 << 13
    assert "VariousPropertyBits" not in _props(host, "BoxRight")


@pytest.mark.parametrize("host", HOSTS)
def test_a_combo_boxs_style_is_its_display_style(host: str) -> None:
    assert (_props(host, "Combo")["DisplayStyle"], _props(host, "DropList")["DisplayStyle"]) == (3, 7)


@pytest.mark.parametrize("host", HOSTS)
def test_triple_state_is_stored_in_multi_select(host: str) -> None:
    assert _props(host, "Triple")["MultiSelect"] == 1


@pytest.mark.parametrize("host", HOSTS)
def test_scroll_bars_carry_keep_scroll_bars_visible(host: str) -> None:
    # ScrollBars = 3 (both) and KeepScrollBarsVisible = 1 (horizontal).
    assert RECORD[host]["form"]["properties"]["ScrollBars"] == 3 | 1 << 2


@pytest.mark.parametrize("host", HOSTS)
def test_a_forms_font_effects_are_stdfont_flags(host: str) -> None:
    # Italic and underline, bits 1 and 2; bold would be the weight.
    font = RECORD[host]["form"]["font"]
    assert (font["flags"], font["weight"]) == (0b0110, 400)


@pytest.mark.parametrize("host", HOSTS)
def test_column_widths_are_column_info_records_last_column_first(host: str) -> None:
    # ColumnWidths "20 pt;30 pt;40 pt": each record is 12 bytes, the width last, in HIMETRIC.
    record = _record(host, "Columns")
    assert record.values["cColumnInfo"] == 3
    tail = record.tail_raw
    widths = [int.from_bytes(tail[at + 8:at + 12], "little") for at in range(0, len(tail), 12)]
    assert widths == [1411, 1058, 705]


@pytest.mark.parametrize("host", HOSTS)
def test_a_label_is_sited_without_tab_stop(host: str) -> None:
    assert {_control(host, name)["site"].get("BitFlags") for name in ("AlignLeft", "Bold")} == {0x32}


@pytest.mark.parametrize("host", HOSTS)
def test_enabled_false_dims_the_text_but_not_a_list_boxs(host: str) -> None:
    assert _effects(host, "Disabled") & DISABLED
    assert not _effects(host, "DisabledList") & DISABLED


@pytest.mark.parametrize("host", HOSTS)
def test_every_font_effect_keeps_the_colour_automatic(host: str) -> None:
    effects = [_effects(host, name) for name in RECORD[host]["controls"] if _effects(host, name)]
    assert effects and all(value & AUTO_COLOR for value in effects)


@pytest.mark.parametrize("host", HOSTS)
def test_bold_is_an_effect_and_a_weight(host: str) -> None:
    assert _effects(host, "Bold") & 1
    assert _props(host, "Bold")["Font.FontWeight"] == 700


@pytest.mark.parametrize("host", HOSTS)
def test_a_new_face_drops_the_charset(host: str) -> None:
    assert "Font.FontCharSet" in _props(host, "BoxRight")
    assert "Font.FontCharSet" not in _props(host, "Face")


@pytest.mark.parametrize("host", HOSTS)
def test_border_style_and_special_effect_clear_each_other(host: str) -> None:
    # A TextBox is sunken, SpecialEffect 2, and an Image single-bordered, BorderStyle 1, by default: those
    # are stored only when they differ.
    def stored(name: str) -> tuple[object, object]:
        props = _props(host, name)
        return props.get("BorderStyle"), props.get("SpecialEffect")

    assert stored("BorderThenEffect") == (None, None)
    assert stored("EffectThenBorder") == (1, 0)
    assert stored("ImageBorderThenEffect") == (0, 2)
    assert stored("ImageEffectThenBorder") == (None, None)


@pytest.mark.parametrize("host", HOSTS)
def test_a_textbox_given_a_border_alone_keeps_both_and_shows_the_effect(host: str) -> None:
    # set_property("BorderStyle", 1) leaves a TextBox sunken. Excel and Word report both at run time, paint
    # the TextBox's corner as they paint a plain sunken one and not as the designers' bordered one, and a
    # designer that rewrites the form saves both back.
    measured = EDGE[host]
    assert measured["written"]["Both"] == {"BorderStyle": 1, "SpecialEffect": None}
    shown = measured["run_time"]
    assert (shown["Both"]["BorderStyle"], shown["Both"]["SpecialEffect"]) == (1, 2)
    assert shown["Both"]["corner"] == shown["Effect"]["corner"] != shown["Border"]["corner"]
    assert measured["resaved"] == measured["written"]


@pytest.mark.parametrize("host", HOSTS)
def test_a_disabled_scroll_bar_disables_its_arrows(host: str) -> None:
    for name in ("BarDisabled", "SpinDisabled"):
        props = _props(host, name)
        assert (props["PrevEnabled"], props["NextEnabled"]) == (0, 0), name


@pytest.mark.parametrize("host", HOSTS)
def test_min_moves_the_position_up_to_it(host: str) -> None:
    props = _props(host, "BarMin")
    assert (props["Min"], props["Position"]) == (50, 50)


@pytest.mark.parametrize("host", HOSTS)
def test_take_focus_on_click_is_a_mask_bit_with_no_data(host: str, tmp_path: Path) -> None:
    # CommandButton PropMask bit 9 is set when TakeFocusOnClick is False; nothing follows it in the record.
    assert _props(host, "NoFocus")["TakeFocusOnClick"] is False
    assert "TakeFocusOnClick" not in _props(host, "Plain")
    assert _control(host, "NoFocus")["mask"] == _control(host, "Plain")["mask"] | 1 << 9
    with ExcelFile.create_new(tmp_path / "focus.xlsm") as workbook:
        button = workbook.add_form("Probe").add_control("CommandButton", "NoFocus")
        button.set_property("Caption", None)
        button.set_property("TakeFocusOnClick", False)
        assert button.get("TakeFocusOnClick") is False
        assert button.properties_set == _control(host, "NoFocus")["mask"]
        button.set_property("TakeFocusOnClick", True)
        assert button.get("TakeFocusOnClick") is None
        assert button.properties_set == _control(host, "Plain")["mask"]
