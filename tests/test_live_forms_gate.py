"""Live Excel gate for the UserForm designer reader and writer (opt-in).

Every other forms test checks the file against itself: the counts inside
``f`` reconcile, the site sizes account for ``o``.  Those catch a reader
that walks off the rails, but they cannot catch one that walks a
consistent path to the wrong answer -- a site array misread the same way
twice still adds up.

This gate is the outside check the issue asked for first: open the same
workbook in Excel and require that every control Excel reports is one the
file-level reader found, with the same type.  The extras are allowed to
run the other way (a MultiPage's Pages and its hidden TabStrip are real
structure that ``Designer.Controls`` does not enumerate), but nothing
Excel knows about may be missing.

The write half matters more.  Every structural check in the reader would
pass on a form Excel then refuses to load, and two defects were caught
only here: an added control taking an id that collides with the last one,
and a designer edit that left the ``_VBA_PROJECT`` performance cache in
place so Office loaded a member list the form no longer matched.

Opt-in: set ``RUN_LIVE_EXCEL=1`` on a Windows machine with desktop Excel
and ``pyvbaharness`` installed.  Skipped everywhere else, including CI.
``pyvbaharness`` is a test-time oracle only; pyOpenVBA never uses COM.
"""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

import pytest

from pyopenvba import ExcelFile

pytestmark = pytest.mark.skipif(
    os.environ.get("RUN_LIVE_EXCEL") != "1" or sys.platform != "win32",
    reason="live Excel gate: set RUN_LIVE_EXCEL=1 on Windows with Excel installed",
)

_NESTED = Path(__file__).parent / "live_excel_testing" / "nested_form.xlsm"
_FORM = "FrmNested"
_TIMEOUT = 60.0

# Pages are real objects reached through MultiPage.Pages, and the TabStrip
# is MSForms' own; neither appears in Designer.Controls.
_NOT_IN_CONTROLS = {"Page1", "Page2", ""}


def _excel_controls(path: Path) -> list[tuple[str, str]]:
    """Ask Excel for the form's controls: (name, TypeName)."""
    harness = pytest.importorskip("pyvbaharness")

    session = harness.ExcelSession()
    try:
        session.open_document(str(path.resolve()), read_only=False, timeout=120.0)
        designer = f'ActiveWorkbook.VBProject.VBComponents("{_FORM}").Designer'
        count = int(session.eval(f"{designer}.Controls.Count", timeout=_TIMEOUT))
        return [
            (
                str(session.eval(f"{designer}.Controls({i}).Name", timeout=_TIMEOUT)),
                str(session.eval(
                    f"TypeName({designer}.Controls({i}))", timeout=_TIMEOUT
                )),
            )
            for i in range(count)
        ]
    finally:
        try:
            session.close()
        except Exception:
            pass


@pytest.mark.skipif(not _NESTED.exists(), reason="nested form fixture not present")
class TestFormsAgreeWithExcel:
    def test_every_control_excel_reports_is_found_with_the_same_type(self) -> None:
        with ExcelFile(_NESTED) as workbook:
            form = next(f for f in workbook.forms() if f.name == _FORM)
        # Excel's TypeName is the MSForms class without its library prefix.
        mine = {c.name: c.kind.split(".")[-1] for c in form.walk() if c.name}

        theirs = _excel_controls(_NESTED)
        assert theirs, "Excel reported no controls at all"
        missing = [name for name, _ in theirs if name not in mine]
        assert not missing, f"controls Excel reports but the file reader missed: {missing}"
        mismatched = [
            (name, mine[name], kind) for name, kind in theirs if mine[name] != kind
        ]
        assert not mismatched, f"type disagreements (name, ours, Excel): {mismatched}"

    def test_extras_are_only_the_structure_designer_controls_hides(self) -> None:
        with ExcelFile(_NESTED) as workbook:
            form = next(f for f in workbook.forms() if f.name == _FORM)
        theirs = {name for name, _ in _excel_controls(_NESTED)}
        extras = {c.name for c in form.walk()} - theirs
        assert extras <= _NOT_IN_CONTROLS, (
            f"the reader invented controls Excel does not have: "
            f"{sorted(extras - _NOT_IN_CONTROLS)}"
        )


@pytest.mark.skipif(not _NESTED.exists(), reason="nested form fixture not present")
class TestExcelLoadsWhatWeWrite:
    """A form Excel cannot load answers nothing at all, so every probe
    here doubles as proof the file still opens."""

    def test_an_edited_property_is_what_excel_reads(self, tmp_path: Path) -> None:
        out = tmp_path / "edited.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            form = workbook.forms()[0]
            form.control("CloseButton").set_property("Caption", "Dismiss")
            form.control("TopLabel").set_property("ForeColor", 255)
            form.control("OptOne").set_property("Caption", "Ground")
            workbook.save()
        assert _designer_values(
            out,
            [
                ('Controls("CloseButton").Caption', "Dismiss"),
                ('Controls("TopLabel").ForeColor', 255),
                ('Controls("OptOne").Caption', "Ground"),
                ("Controls.Count", 10),
            ],
        )

    def test_added_controls_are_what_excel_reads(self, tmp_path: Path) -> None:
        out = tmp_path / "added.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            form = workbook.forms()[0]
            form.add_control("CommandButton", "AddedButton", left=12, top=250,
                             width=90, height=24)
            form.add_control("TextBox", "AddedBox", left=120, top=250)
            form.add_control("OptionButton", "InFrame", container="GroupBox",
                             left=6, top=60)
            workbook.save()
        assert _designer_values(
            out,
            [
                ('TypeName(Controls("AddedButton"))', "CommandButton"),
                ('Controls("AddedButton").Width', 90),
                ('Controls("AddedButton").Left', 12),
                ('TypeName(Controls("AddedBox"))', "TextBox"),
                ('TypeName(Controls("InFrame"))', "OptionButton"),
                ('Controls("InFrame").Parent.Name', "GroupBox"),
                ("Controls.Count", 13),
            ],
        )

    def test_a_removed_control_is_gone_and_the_rest_survive(
        self, tmp_path: Path
    ) -> None:
        out = tmp_path / "removed.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            workbook.forms()[0].remove_control("Picker")
            workbook.save()
        assert _designer_values(
            out,
            [
                ("Controls.Count", 9),
                ('Controls("CloseButton").Caption', "Close"),
                ('Controls("GroupBox").Controls.Count', 3),
            ],
        )


def _designer_values(path: Path, checks: list[tuple[str, object]]) -> bool:
    """Evaluate each expression against the form's designer in Excel."""
    harness = pytest.importorskip("pyvbaharness")

    session = harness.ExcelSession()
    try:
        session.open_document(str(path.resolve()), read_only=False, timeout=120.0)
        designer = f'ActiveWorkbook.VBProject.VBComponents("{_FORM}").Designer'
        wrong: list[str] = []
        for expression, expected in checks:
            probe = expression.replace("Controls", f"{designer}.Controls", 1)
            got = session.eval(probe, timeout=_TIMEOUT)
            if got != expected:
                wrong.append(f"{expression}: expected {expected!r}, got {got!r}")
        assert not wrong, "; ".join(wrong)
        return True
    finally:
        try:
            session.close()
        except Exception:
            pass


@pytest.mark.skipif(not _NESTED.exists(), reason="nested form fixture not present")
class TestExcelBindsContainersWeCreate:
    """A container's storage is bound by its CLSID and by the CompObj
    naming what fm20 should treat it as.  Get either wrong and the
    container loads without erroring and simply does not appear, which is
    a failure only a live host can show."""

    def test_a_new_frame_appears_and_holds_controls(self, tmp_path: Path) -> None:
        out = tmp_path / "frame.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            form = workbook.forms()[0]
            form.add_control("Frame", "NewFrame", left=12, top=250,
                             width=180, height=70)
            form.add_control("CommandButton", "Inside", container="NewFrame",
                             left=6, top=12)
            workbook.save()
        assert _designer_values(
            out,
            [
                ('TypeName(Controls("NewFrame"))', "Frame"),
                ('Controls("NewFrame").Caption', "NewFrame"),
                ('Controls("NewFrame").Controls.Count', 1),
                ('Controls("Inside").Parent.Name', "NewFrame"),
                ("Controls.Count", 12),
            ],
        )

    def test_removing_a_frame_takes_its_children_and_leaves_the_rest(
        self, tmp_path: Path
    ) -> None:
        out = tmp_path / "framegone.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            workbook.forms()[0].remove_control("GroupBox")
            workbook.save()
        assert _designer_values(
            out,
            [
                ("Controls.Count", 6),
                ('Controls("CloseButton").Caption', "Close"),
                ('Controls("Pages").Pages.Count', 2),
            ],
        )


@pytest.mark.skipif(not _NESTED.exists(), reason="nested form fixture not present")
class TestExcelBindsPagesWeCreate:
    """A page is a container, a tab and an entry in the page bookkeeping.
    Writing three of the four leaves a form that still passes every
    structural check here and that Excel then refuses to load."""

    def test_a_new_page_appears_on_its_multipage(self, tmp_path: Path) -> None:
        out = tmp_path / "addpage.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            workbook.forms()[0].add_page("Pages", name="Extra", caption="Extra tab")
            workbook.save()
        assert _designer_values(
            out,
            [
                ('Controls("Pages").Pages.Count', 3),
                ('Controls("Pages").Pages(2).Name', "Extra"),
                ('Controls("Pages").Pages(2).Caption', "Extra tab"),
                ('Controls("Pages").Pages(0).Name', "Page1"),
            ],
        )

    def test_a_removed_page_takes_its_tab_and_its_controls(
        self, tmp_path: Path
    ) -> None:
        out = tmp_path / "rmpage.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            workbook.forms()[0].remove_page("Page2")
            workbook.save()
        assert _designer_values(
            out,
            [
                ('Controls("Pages").Pages.Count', 1),
                ('Controls("Pages").Pages(0).Name', "Page1"),
                ('TypeName(Controls("PageOneCheck"))', "CheckBox"),
                ("Controls.Count", 9),
            ],
        )

    def test_a_multipage_built_from_nothing_loads(self, tmp_path: Path) -> None:
        out = tmp_path / "newmp.xlsm"
        shutil.copyfile(_NESTED, out)
        with ExcelFile(out) as workbook:
            form = workbook.forms()[0]
            form.add_control("MultiPage", "Fresh", left=12, top=250,
                             width=200, height=90)
            form.add_page("Fresh", name="Third", caption="Third tab")
            form.add_control("Label", "OnThird", container="Third", left=6, top=6)
            workbook.save()
        assert _designer_values(
            out,
            [
                ('TypeName(Controls("Fresh"))', "MultiPage"),
                ('Controls("Fresh").Pages.Count', 3),
                ('Controls("Fresh").Pages(0).Name', "Page1"),
                ('Controls("Fresh").Pages(2).Caption', "Third tab"),
                ('Controls("OnThird").Parent.Name', "Third"),
                # The MultiPage the fixture already had is untouched.
                ('Controls("Pages").Pages.Count', 2),
            ],
        )
