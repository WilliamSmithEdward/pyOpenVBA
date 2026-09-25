"""Where Excel's and Word's designers store a property VBA names, and what setting it changes besides.

tests/fixtures/form_properties.json is what scripts/measure_form_properties.py
saw when each designer set one property per control through VBA and saved:
each control's record as the library's reader takes it apart. The two
applications wrote the same bytes. The form's own font was set italic and
underlined first, so every control carries those two effects and the
fAutoColor flag that comes with any effect. docs/userforms.md lists these
as the reference for writing the stored fields.
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
