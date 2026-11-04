"""A fifteenth sweep of everyday Excel macros: WorksheetFunction, where a function that would show #N/A in a cell stops the macro with 1004. Each
probe reads a list of calls, each with its type and value or its error, on a small sheet of names, numbers, blanks and text.

    python scripts/measure_excel_everyday_15.py

writes tests/fixtures/excel_everyday_15.json, which tests/test_excel_everyday_15.py replays.
"""
from __future__ import annotations

import json
import re
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/excel_everyday_15.json'
PRE = ('Dim ws As Worksheet, f As Object, s As String, v As Variant, a As Variant, r As Range, b As Range, n As Range, t As Range, e As Range, m As Range\n'
       'Set ws = Worksheets.Add\nws.Name = "Sweep"\nSet f = Application.WorksheetFunction\n'
       'ws.Range("A1:C1").Value = Array("Item", "Qty", "Cost")\nws.Range("A2:C2").Value = Array("Pen", 10, 1.5)\nws.Range("A3:C3").Value = Array("Ink", 3, 12.25)\n'
       'ws.Range("A4:C4").Value = Array("Pad", 7, 4)\nws.Range("A5:C5").Value = Array("Cap", 20, 2)\nws.Range("A6:C6").Value = Array("Pen", 5, 1.5)\n'
       'Set r = ws.Range("A2:C6"): Set b = ws.Range("B2:B6"): Set n = ws.Range("A2:A6"): Set t = ws.Range("C2:C6"): Set e = ws.Range("H1:H3")\n'
       'ws.Range("E1").Value = 1: ws.Range("E2").Value = "x": ws.Range("E3").Value = True: ws.Range("E5").Formula = "=1/0": Set m = ws.Range("E1:E5")\n')


def calls(*expressions: str) -> str:
    return 'On Error Resume Next\n' + '\n'.join(
        f'v = Empty\nv = {one}\ns = s & Err.Number & "=" & TypeName(v) & ":" & Show(v) & ";"\nErr.Clear' for one in expressions) + '\nOn Error GoTo 0'


GROUPS = {
    'lookup': ['f.VLookup("Ink", r, 2, False)', 'f.VLookup("Nope", r, 2, False)', 'f.VLookup("Ink", r, 9, False)', 'f.VLookup("Ink", r, 0, False)', 'f.VLookup("I*", r, 3, False)', 'f.VLookup("Ink", r, 2)',
               'f.VLookup("Ink", r, 2, True)', 'f.VLookup(7, ws.Range("B2:C6"), 2, False)', 'f.VLookup("7", ws.Range("B2:C6"), 2, False)', 'f.VLookup(Empty, r, 2, False)', 'f.VLookup(Null, r, 2, False)', 'f.VLookup("Ink", "x", 2, False)', 'f.VLookup("Ink", Array(1), 1, False)',
               'f.HLookup("Qty", ws.Range("A1:C6"), 2, False)', 'f.HLookup("Nope", ws.Range("A1:C6"), 2, False)',
               'f.Match("Pad", n, 0)', 'f.Match("Nope", n, 0)', 'f.Match("P*", n, 0)', 'f.Match(7, b, 0)', 'f.Match(7, b, 1)', 'f.Match(7, b)', 'f.Match(Array(1, 2), b, 0)', 'f.Match("Pad", Array("Pen", "Pad"), 0)', 'f.Match("Pad", ws.Range("A1:C1"), 0)',
               'f.Index(r, 2, 2)', 'f.Index(r, 9, 2)', 'TypeName(f.Index(r, 2, 2))', 'f.Index(r, 2).Address', 'f.Index(b, 2)', 'f.Index(Array(5, 6), 2)', 'f.Index(Array(5, 6), 1, 2)', 'f.Index(r, 0, 1).Address',
               'f.Lookup(7, b, t)', 'f.Lookup(1000, b, t)', 'f.Lookup(-1, b, t)', 'f.Choose(2, "a", "b")', 'f.Choose(9, "a", "b")', 'f.Choose(2, ws.Range("A2"), ws.Range("A3")).Address',
               'f.XLookup("Ink", n, b)', 'f.XLookup("Nope", n, b)', 'f.XLookup("Nope", n, b, "none")', 'f.XMatch("Ink", n)'],
    'aggregate': ['f.Sum(b)', 'f.Sum(b, t)', 'f.Sum(1, 2, 3)', 'f.Sum("1", 2)', 'f.Sum("x", 2)', 'f.Sum(True, 2)', 'f.Sum(Empty)', 'f.Sum(Null)', 'f.Sum(m)', 'f.Sum(n)', 'f.Sum(e)', 'f.Sum(Array(1, 2))', 'f.Sum(Array(1, "2", True))', 'f.Sum()',
                  'f.Average(b)', 'f.Average(n)', 'f.Average(e)', 'f.Average(1, "x")', 'f.AverageIf(n, "Pen", b)', 'f.AverageIf(n, "Nope", b)', 'f.AverageIfs(b, n, "Pen")', 'f.AverageIfs(b, n, "Nope")',
                  'f.Count(b)', 'f.Count(n)', 'f.Count(1, "x", True)', 'f.CountA(n)', 'f.CountA(e)', 'f.CountA(m)', 'f.CountBlank(e)', 'f.CountBlank(m)', 'f.CountIf(n, "Pen")', 'f.CountIf(n, "P*")', 'f.CountIf(b, ">5")', 'f.CountIf(b, 7)', 'f.CountIf(Array(1), 1)', 'f.CountIfs(n, "Pen", b, ">5")', 'f.CountIfs(n, "Pen", ws.Range("B2:B3"), ">5")',
                  'f.SumIf(n, "Pen", b)', 'f.SumIf(b, ">5")', 'f.SumIf(n, "Nope", b)', 'f.SumIfs(b, n, "Pen")', 'f.SumIfs(b, n, "Pen", t, ">1")', 'f.SumIfs(b, ws.Range("A2:A3"), "Pen")', 'f.SumProduct(b, t)', 'f.SumProduct(b, n)', 'f.SumProduct(b, ws.Range("C2:C3"))',
                  'f.Max(b)', 'f.Max(n)', 'f.Max(e)', 'f.Max(1, "2")', 'f.Max(1, "x")', 'f.Min(b)', 'f.Min(e)', 'f.MaxIfs(b, n, "Pen")', 'f.MinIfs(b, n, "Nope")', 'f.Large(b, 2)', 'f.Large(b, 9)', 'f.Large(n, 1)', 'f.Small(b, 1)', 'f.Small(b, 0)', 'f.Rank(7, b)', 'f.Rank(8, b)', 'f.Rank_Eq(7, b, 1)', 'f.Median(b)', 'f.Median(e)', 'f.Mode(ws.Range("C2:C6"))', 'f.Mode(b)', 'f.Mode_Sngl(b)', 'f.StDev(b)', 'f.StDev(1)', 'f.Var(b)', 'f.Var_S(b)', 'f.Percentile(b, 0.5)', 'f.Percentile(b, 2)', 'f.Quartile(b, 1)', 'f.Subtotal(9, b)', 'f.Subtotal(99, b)', 'f.Aggregate(9, 6, m)', 'f.Product(b)', 'f.GeoMean(b)', 'f.Correl(b, t)', 'f.Slope(b, t)', 'f.Intercept(b, t)', 'f.Forecast_Linear(5, b, t)', 'f.Trend(b)', 'f.Frequency(b, ws.Range("C2:C3"))'],
    'math': ['f.Round(2.5, 0)', 'f.Round(2.345, 2)', 'f.Round("2.5", 0)', 'f.Round("x", 0)', 'f.RoundUp(2.1, 0)', 'f.RoundDown(2.9, 0)', 'f.MRound(7, 2)', 'f.MRound(7, -2)', 'f.Ceiling(2.1, 1)', 'f.Ceiling_Math(2.1)', 'f.Ceiling(2.1, -1)', 'f.Floor(2.9, 1)', 'f.Floor_Math(2.9)', 'f.Power(2, 3)', 'f.Power(0, 0)', 'f.Power(-1, 0.5)', 'f.Fact(3)', 'f.Fact(-1)', 'f.Combin(5, 2)', 'f.Combin(2, 5)', 'f.Permut(5, 2)', 'f.Gcd(12, 18)', 'f.Lcm(4, 6)', 'f.Quotient(7, 2)', 'f.Quotient(7, 0)', 'f.Log(8, 2)', 'f.Log(-1, 2)', 'f.Log10(100)', 'f.Ln(1)', 'f.Ln(0)', 'f.Pi()', 'f.Pi', 'f.Radians(180)', 'f.Degrees(3.14)', 'f.Sinh(0)', 'f.Asin(2)', 'f.Acos(1)', 'f.Atan2(1, 1)', 'f.Roman(4)', 'f.Arabic("IV")', 'f.Base(10, 2)', 'f.Decimal("A", 16)', 'f.Dec2Bin(5)', 'f.Hex2Dec("FF")', 'f.RandBetween(1, 1)', 'f.Rand()', 'f.Mod(1, 1)', 'f.Abs(1)', 'f.Int(1)', 'f.Sqrt(4)', 'f.Sign(1)', 'f.Exp(1)', 'f.Now()', 'f.If(True, 1, 2)', 'f.Left("a", 1)', 'f.Len("a")', 'f.Date(2020, 1, 1)', 'f.IsNumber(1)'],
    'text': ['f.Trim("  a  b  ")', 'f.Clean("a" & Chr(7))', 'f.Proper("ab cd")', 'f.Substitute("aXa", "a", "b")', 'f.Substitute("aXa", "a", "b", 2)', 'f.Substitute("aXa", "a", "b", 0)', 'f.Replace("abc", 2, 1, "X")', 'f.Replace("abc", 0, 1, "X")', 'f.Rept("a", 2)', 'f.Rept("a", -1)', 'f.Find("b", "abc")', 'f.Find("z", "abc")', 'f.Find("b", "abc", 9)', 'f.Search("B", "abc")', 'f.Search("z", "abc")', 'f.Text(1.5, "0.00")', 'f.Text(1, "x")', 'f.Text("a", "0")', 'f.Text(#1/2/2020#, "yyyy-mm-dd")', 'f.Dollar(1.5)', 'f.Fixed(1.5, 0)', 'f.Value', 'f.NumberValue("1,5", ",", ".")', 'f.Concat("a", "b")', 'f.Concat(n)', 'f.TextJoin(",", True, n)', 'f.TextJoin(",", False, e)', 'f.Unichar(65)', 'f.Unicode("A")', 'f.Code("A")', 'f.Exact("a", "A")', 'f.Lower("A")', 'f.Upper("a")', 'f.T("a")', 'f.N("a")', 'f.Bahttext(1)', 'f.Encodeurl("a b")'],
    'date': ['f.WeekNum(#1/2/2020#)', 'f.WeekNum(#1/2/2020#, 2)', 'f.WeekNum(-1)', 'f.Weekday(#1/2/2020#)', 'f.IsoWeekNum(#1/2/2020#)', 'f.EDate(#1/31/2020#, 1)', 'TypeName(f.EDate(#1/31/2020#, 1))', 'f.EoMonth(#1/15/2020#, 0)', 'f.EoMonth("x", 0)', 'f.WorkDay(#1/2/2020#, 1)', 'f.WorkDay_Intl(#1/2/2020#, 1, 11)', 'f.NetworkDays(#1/1/2020#, #1/10/2020#)', 'f.NetworkDays_Intl(#1/1/2020#, #1/10/2020#, 11)', 'f.Days(#1/10/2020#, #1/1/2020#)', 'f.Days360(#1/1/2020#, #3/1/2020#)', 'f.YearFrac(#1/1/2020#, #7/1/2020#)', 'f.DateDif', 'f.DateValue("1/2/2020")', 'f.TimeValue', 'f.Year', 'f.Today'],
    'finance': ['f.Pmt(0.01, 12, 100)', 'f.Pmt(0, 0, 100)', 'f.PV(0.01, 12, 10)', 'f.FV(0.01, 12, 10)', 'f.NPer(0.01, 10, 100)', 'f.Rate(12, 10, -100)', 'f.Npv(0.1, b)', 'f.Irr(ws.Range("I1:I3"))', 'f.IPmt(0.01, 1, 12, 100)', 'f.PPmt(0.01, 1, 12, 100)', 'f.Sln(100, 10, 5)', 'f.Syd(100, 10, 5, 1)', 'f.Db(100, 10, 5, 1)', 'f.Ddb(100, 10, 5, 1)', 'f.Effect(0.1, 12)', 'f.Nominal(0.1, 12)'],
    'arrays': ['UBound(f.Transpose(b), 1)', 'UBound(f.Transpose(Array(1, 2)), 1)', 'f.Transpose(b)(1, 1)', 'TypeName(f.Transpose(b))', 'UBound(f.Transpose(n))', 'f.Transpose(Array(1, 2))(2, 1)', 'f.MMult(ws.Range("B2:B2"), ws.Range("B2:B2"))', 'f.MMult(ws.Range("B2:C3"), ws.Range("B2:C3"))(2, 2)', 'f.MDeterm(ws.Range("B2:C3"))', 'f.MInverse(ws.Range("B2:C3"))(1, 1)', 'f.Unique(n)', 'UBound(f.Unique(n), 1)', 'UBound(f.Sort(b), 1)', 'f.Sort(b)(1, 1)', 'f.SortBy(n, b)(1, 1)', 'UBound(f.Filter(n, ws.Range("B2:B6")), 1)', 'f.Sequence(2)(2, 1)', 'f.RandArray', 'f.Let', 'f.Lambda'],
    'info': ['f.IsError(1)', 'f.IsError(ws.Range("E5"))', 'f.IsNA(ws.Range("E5"))', 'f.IsErr(ws.Range("E5"))', 'f.IsText(ws.Range("A1"))', 'f.IsText("a")', 'f.IsNumber(ws.Range("B2"))', 'f.IsLogical(True)', 'f.IsNonText(1)', 'f.IsBlank(ws.Range("H1"))', 'f.IsRef(ws.Range("A1"))', 'f.IsEven(2)', 'f.IsOdd(2)', 'f.IsFormula(ws.Range("E5"))', 'f.NA()', 'f.And(True, False)', 'f.Or(True, False)', 'f.Xor(True, False)', 'f.Not(True)', 'f.IfError(1, 2)', 'f.IfNa(1, 2)', 'f.Rows(r)', 'f.Columns(r)', 'f.Row(r)', 'f.Address(1, 1)', 'f.Indirect("A1")', 'f.Offset(r, 1, 1)', 'f.Hyperlink("x")', 'f.Cell("row", r)', 'f.Info("osversion")', 'f.Sheet', 'f.Sheets', 'f.Type(1)', 'f.Error_Type', 'f.Delta(1, 1)', 'f.GeStep(1, 1)', 'f.Bin2Dec("101")', 'f.Convert(1, "m", "cm")', 'f.Convert(1, "m", "kg")', 'f.Erf(1)', 'f.Norm_S_Dist(0, True)', 'f.Norm_Dist(0, 0, 1, True)', 'f.Norm_Inv(0.5, 0, 1)', 'f.Binom_Dist(1, 2, 0.5, False)', 'f.Poisson_Dist(1, 1, False)', 'f.T_Dist(1, 2, True)', 'f.Chisq_Dist(1, 2, True)', 'f.Confidence_Norm(0.05, 1, 10)', 'f.Fisher(0.5)', 'f.Gamma(5)', 'f.GammaLn(5)', 'f.Beta_Dist(0.5, 1, 1, True)', 'f.Expon_Dist(1, 1, True)', 'f.Weibull_Dist(1, 1, 1, True)', 'f.Hypgeom_Dist(1, 2, 2, 4, False)', 'f.LogNorm_Dist(1, 0, 1, True)', 'f.F_Dist(1, 1, 1, True)', 'f.DSum(ws.Range("A1:C6"), "Qty", ws.Range("K1:K2"))', 'f.DCount(ws.Range("A1:C6"), 2, ws.Range("A1:A2"))', 'f.DGet(ws.Range("A1:C6"), "Qty", ws.Range("K1:K2"))'],
}
CASES = {name: calls(*expressions) for name, expressions in GROUPS.items()}
HELPERS = """
Private Function Show(ByVal v As Variant) As String
    If IsObject(v) Then
        Show = "object"
    ElseIf IsArray(v) Then
        Show = "array"
    ElseIf IsError(v) Then
        Show = "error"
    ElseIf IsNull(v) Then
        Show = "null"
    Else
        Show = CStr(v)
    End If
End Function
"""


def macro(name: str) -> str:
    cleanup = '\nOn Error Resume Next\nApplication.DisplayAlerts = False\nws.Delete\nApplication.DisplayAlerts = True\nOn Error GoTo 0'
    return HELPERS + '\nFunction Probe() As String\n' + PRE + CASES[name] + cleanup + '\nProbe = s\nEnd Function\n'


def main() -> None:
    probes: list[dict[str, object]] = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for name in CASES:
            result = excel.run_vba(macro(name), 'Probe', timeout=120.0, line_numbers=False)
            assert result.ok and not result.dialogs, (name, result.error, result.message, result.dialogs)
            probes.append({'name': name, 'calls': GROUPS[name], 'macro': macro(name), 'result': result.value})
            print(name, str(result.value)[:120], flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'probes': probes}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
