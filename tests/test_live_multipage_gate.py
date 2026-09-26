"""Live gate: a MultiPage the library makes at a size shows every page whole when its form runs (opt-in).

When VBA sizes a MultiPage, the designer lays out only the page it shows, so
a MultiPage sized to 300 x 200 points keeps its second page laid out for the
default size, 141 x 90.75 points (tests/fixtures/multipage_resize.json), and
the library makes one the same way. Each application opens a file the
library made with such a MultiPage: a red Label on the second page, far
outside that layout, and a green one at the same place on the first. The
form runs, shows each page and paints itself (PrintWindow); the red Label
shows as many pixels as the green one, and the second page is as large
inside as the first once it is shown, to within a pixel: a running form
lays out again each page it shows, and rounds the page it switches to a
tenth of a point apart from the one it opened on.

Opt-in per application: RUN_LIVE_EXCEL=1 or RUN_LIVE_WORD=1, on Windows
with it installed. Skipped everywhere else, including CI.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Any

import pytest

from pyopenvba import ExcelFile, WordFile
from pyopenvba._oforms_records import Size

#: Each application: the switch that opts in, the library's file type, the extension and the harness session.
HOSTS: dict[str, tuple[str, Any, str, str]] = {
    "excel": ("RUN_LIVE_EXCEL", ExcelFile, "xlsm", "ExcelSession"),
    "word": ("RUN_LIVE_WORD", WordFile, "docm", "WordSession"),
}
RED, GREEN = 0x0000FF, 0x00FF00
#: Shows the form at one page, has it paint itself into a bitmap and counts every second pixel of each colour.
PROBE = "\r\n".join([
    "Option Explicit",
    "Private Type RECT",
    "    Left As Long",
    "    Top As Long",
    "    Right As Long",
    "    Bottom As Long",
    "End Type",
    'Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal c As String, ByVal t As String) As LongPtr',
    'Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal h As LongPtr, r As RECT) As Long',
    'Private Declare PtrSafe Function GetDC Lib "user32" (ByVal h As LongPtr) As LongPtr',
    'Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal h As LongPtr, ByVal dc As LongPtr) As Long',
    'Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal dc As LongPtr) As LongPtr',
    'Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal dc As LongPtr, '
    'ByVal w As Long, ByVal h As Long) As LongPtr',
    'Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal dc As LongPtr, ByVal o As LongPtr) As LongPtr',
    'Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal o As LongPtr) As Long',
    'Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal dc As LongPtr) As Long',
    'Private Declare PtrSafe Function PrintWindow Lib "user32" (ByVal h As LongPtr, ByVal dc As LongPtr, '
    'ByVal flags As Long) As Long',
    'Private Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal dc As LongPtr, ByVal x As Long, '
    'ByVal y As Long) As Long',
    "",
    "Public Function Measure(ByVal index As Long) As String",
    "    Dim h As LongPtr, r As RECT, w As Long, ht As Long, x As Long, y As Long, px As Long",
    "    Dim sdc As LongPtr, mdc As LongPtr, bmp As LongPtr, old As LongPtr, red As Long, green As Long",
    "    Pages.Show vbModeless",
    "    Pages.MP.Value = index",
    "    Pages.Repaint",
    "    DoEvents",
    '    h = FindWindowA("ThunderDFrame", Pages.Caption)',
    "    GetWindowRect h, r",
    "    w = r.Right - r.Left: ht = r.Bottom - r.Top",
    "    sdc = GetDC(0)",
    "    mdc = CreateCompatibleDC(sdc)",
    "    bmp = CreateCompatibleBitmap(sdc, w, ht)",
    "    old = SelectObject(mdc, bmp)",
    "    PrintWindow h, mdc, 0",
    "    For y = 0 To ht - 1 Step 2",
    "        For x = 0 To w - 1 Step 2",
    "            px = GetPixel(mdc, x, y)",
    f"            If px = {RED} Then red = red + 1",
    f"            If px = {GREEN} Then green = green + 1",
    "        Next",
    "    Next",
    "    SelectObject mdc, old",
    "    DeleteObject bmp",
    "    DeleteDC mdc",
    "    ReleaseDC 0, sdc",
    '    Measure = red & "|" & green & "|" & Pages.MP.Pages(index).InsideWidth & "|" & _',
    "        Pages.MP.Pages(index).InsideHeight",
    "    Unload Pages",
    "End Function",
    "",
])


def _param(host: str) -> Any:
    switch = HOSTS[host][0]
    return pytest.param(host, id=host, marks=pytest.mark.skipif(
        os.environ.get(switch) != "1" or sys.platform != "win32",
        reason=f"live MultiPage gate: set {switch}=1 on Windows with the application installed"))


@pytest.mark.parametrize("host", [_param(host) for host in HOSTS])
def test_a_multipage_made_at_a_size_shows_every_page_whole(host: str, tmp_path: Path) -> None:
    harness = pytest.importorskip("pyvbaharness")
    _, kind, suffix, session_name = HOSTS[host]
    target = tmp_path / f"pages.{suffix}"
    with kind.create_new(target) as document:
        form = document.add_form("Pages", caption="Pages", width=340, height=250)
        form.add_control("MultiPage", "MP", left=6, top=6, width=300, height=200)
        for page, name, colour in (("Page2", "Far", RED), ("Page1", "Near", GREEN)):
            label = form.add_control("Label", name, container=page, left=200, top=110, width=40, height=20)
            label.set_property("Caption", None)
            label.set_property("BackColor", colour)
        document.vba_project().add_module("Probe", PROBE)
        document.save()
    with kind(target) as document:
        second = document.forms()[0].control("Page2")
        assert second.get("DisplayedSize") == Size(4974, 3201)

    with getattr(harness, session_name)(harness.HarnessConfig(lock_wait_s=1200.0)) as office:
        office.open_document(target, read_only=False)
        shown = [office.run_macro("Probe.Measure", index, timeout=180.0) for index in (0, 1)]
        compiled = office.compile_project()
    assert all(result.ok for result in shown), [f"{result.outcome}: {result.message}" for result in shown]
    (red1, green1, *inside1), (red2, green2, *inside2) = (str(result.value).split("|") for result in shown)
    assert (int(red1), int(green2)) == (0, 0)
    assert int(red2) == int(green1) > 0
    pairs = zip(inside1, inside2, strict=True)
    assert all(abs(float(second) - float(first)) < 0.75 for first, second in pairs), (inside1, inside2)
    assert compiled.ok, f"{compiled.outcome}: {compiled.message}"
