"""Now and Time are whole seconds, as Excel's are: measured live, five samples of each a fraction of a second apart held no fraction, and oleaut32's
SystemTimeToVariantTime and VarDateFromUdate drop milliseconds. Timer keeps its fraction, as VBA's does."""
from __future__ import annotations

import datetime as dt

import pytest

from pyopenvba.apps.excel import ExcelApplication


def _app(when: dt.datetime) -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.interpreter.freeze(when)
    app.add_module('''Function Fraction(ByVal which As String) As Double
    Dim d As Double
    If which = "now" Then d = CDbl(Now) Else d = CDbl(Time)
    Fraction = d * 86400# - Int(d * 86400# + 0.000001)
End Function
Function Shown() As String
    Shown = CStr(Now) & "|" & CStr(Time)
End Function
Function Ticks() As Double
    Ticks = Timer
End Function
''', name='H')
    return app


@pytest.mark.parametrize('microsecond', [0, 1, 499_999, 500_000, 600_000, 999_999])
def test_now_and_time_drop_what_is_below_a_second(microsecond: int) -> None:
    app = _app(dt.datetime(2020, 1, 2, 15, 4, 5, microsecond))
    assert abs(float(str(app.run('H.Fraction', 'now')))) < 1e-4
    assert abs(float(str(app.run('H.Fraction', 'time')))) < 1e-4
    # Shown to the second it is, never rounded up to the next.
    assert app.run('H.Shown') == '1/2/2020 3:04:05 PM|3:04:05 PM'


def test_timer_keeps_its_fraction() -> None:
    app = _app(dt.datetime(2020, 1, 2, 15, 4, 5, 600_000))
    assert abs(float(str(app.run('H.Ticks'))) - (15 * 3600 + 4 * 60 + 5.6)) < 0.01
