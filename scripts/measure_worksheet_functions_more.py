"""More worksheet functions called from VBA, as Excel answers them: one or two calls of each function the first
probe left out, on the same small sheet, recorded as scripts/measure_worksheet_functions.py records its own.

    python scripts/measure_worksheet_functions_more.py

writes tests/fixtures/worksheet_functions_more.json, which tests/test_excel_worksheet_functions_more.py replays.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_worksheet_functions import BATCH, HELPER, SETUP

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "worksheet_functions_more.json"

A, B, D, F = 'ws.Range("A1:A5")', 'ws.Range("B1:B5")', 'ws.Range("D1:E3")', 'ws.Range("F1:F3")'
X, Y = "Array(1, 2, 3, 4, 5)", "Array(2, 4, 5, 4, 5)"
DB, CRIT = 'ws.Range("G1:J2")', 'ws.Range("G1:G2")'
S, M = "#1/1/2020#", "#1/1/2025#"

#: (function, arguments as VBA); each is called through WorksheetFunction and through Application.
CASES: list[tuple[str, str]] = [
    ("AccrInt", f"#1/1/2020#, #7/1/2020#, #3/1/2020#, 0.05, 1000, 2"), ("AccrIntM", f"{S}, #6/1/2020#, 0.05, 1000"),
    ("Acos", "0.5"), ("Acos", "2"), ("Acosh", "2"), ("Acot", "1"), ("Acoth", "2"),
    ("Aggregate", f"9, 0, {A}"), ("Aggregate", f"14, 6, {A}, 2"),
    ("AmorDegrc", f"2400, #8/19/2008#, #12/31/2008#, 300, 1, 0.15, 1"),
    ("AmorLinc", f"2400, #8/19/2008#, #12/31/2008#, 300, 1, 0.15, 1"),
    ("Arabic", '"MCMXCIV"'), ("ArrayToText", "Array(1, \"a\", True)"), ("ArrayToText", f"{D}, 1"),
    ("Asc", '"abc"'), ("Asin", "0.5"), ("Asinh", "1"), ("Atan2", "1, 1"), ("Atanh", "0.5"),
    ("AveDev", X), ("AverageIfs", f'{A}, {A}, ">1", {A}, "<5"'),
    ("BahtText", "1234.5"), ("Base", "255, 16"), ("Base", "7, 2, 8"),
    ("BesselI", "1.5, 1"), ("BesselJ", "1.9, 2"), ("BesselK", "1.5, 1"), ("BesselY", "2.5, 1"),
    ("Beta_Dist", "2, 8, 10, True, 1, 3"), ("Beta_Inv", "0.685470581, 8, 10, 1, 3"),
    ("BetaDist", "2, 8, 10, 1, 3"), ("BetaInv", "0.685470581, 8, 10, 1, 3"),
    ("Bin2Dec", '"1100100"'), ("Bin2Hex", '"11111011", 4'), ("Bin2Oct", '"1001"'),
    ("Binom_Dist", "6, 10, 0.5, False"), ("Binom_Dist_Range", "60, 0.75, 48"), ("Binom_Inv", "6, 0.5, 0.75"),
    ("BinomDist", "6, 10, 0.5, True"),
    ("Bitand", "13, 25"), ("Bitlshift", "4, 2"), ("Bitor", "23, 10"), ("Bitrshift", "13, 2"), ("Bitxor", "5, 3"),
    ("Ceiling_Precise", "-4.3"), ("Ceiling_Precise", "4.3, 2"),
    ("ChiDist", "18.307, 10"), ("ChiInv", "0.05, 10"), ("ChiSq_Dist", "0.5, 1, True"), ("ChiSq_Dist_RT", "18.307, 10"),
    ("ChiSq_Inv", "0.93, 1"), ("ChiSq_Inv_RT", "0.05, 10"),
    ("ChiSq_Test", f'ws.Range("D1:D3"), ws.Range("E1:E3")'), ("ChiTest", f'ws.Range("D1:D3"), ws.Range("E1:E3")'),
    ("Clean", '"a" & Chr(7) & "b"'), ("Combin", "8, 2"), ("Combina", "4, 3"),
    ("Complex", "3, 4"), ("Complex", '3, 4, "j"'),
    ("Confidence", "0.05, 2.5, 50"), ("Confidence_Norm", "0.05, 2.5, 50"), ("Confidence_T", "0.05, 1, 50"),
    ("Convert", '68, "F", "C"'), ("Convert", '1, "lbm", "kg"'), ("Correl", f"{X}, {Y}"),
    ("Cosh", "4"), ("Cot", "30"), ("Coth", "2"),
    ("CoupDayBs", f"#1/25/2011#, #11/15/2011#, 2, 1"), ("CoupDays", f"#1/25/2011#, #11/15/2011#, 2, 1"),
    ("CoupDaysNc", f"#1/25/2011#, #11/15/2011#, 2, 1"), ("CoupNcd", f"#1/25/2011#, #11/15/2011#, 2, 1"),
    ("CoupNum", f"#1/25/2007#, #11/15/2008#, 2, 1"), ("CoupPcd", f"#1/25/2011#, #11/15/2011#, 2, 1"),
    ("Covar", f"{X}, {Y}"), ("Covariance_P", f"{X}, {Y}"), ("Covariance_S", f"{X}, {Y}"),
    ("CritBinom", "6, 0.5, 0.75"), ("Csc", "15"), ("Csch", "1.5"),
    ("CumIPmt", "0.0075, 360, 125000, 13, 24, 0"), ("CumPrinc", "0.0075, 360, 125000, 13, 24, 0"),
    ("DAverage", f'{DB}, "q", {CRIT}'), ("Days360", f"{S}, #12/31/2020#"), ("Days360", f"{S}, #12/31/2020#, True"),
    ("Db", "1000000, 100000, 6, 1, 7"), ("Dbcs", '"abc"'), ("DCount", f'{DB}, "q", {CRIT}'),
    ("DCountA", f'{DB}, "q", {CRIT}'), ("Ddb", "2400, 300, 10, 1"),
    ("Dec2Bin", "9, 4"), ("Dec2Bin", "-100"), ("Dec2Hex", "100, 4"), ("Dec2Oct", "58, 3"),
    ("Decimal", '"FF", 16'), ("Degrees", "3.14159265358979"), ("Delta", "5, 4"), ("Delta", "5"),
    ("DGet", f'{DB}, "r", {CRIT}'), ("Disc", f"#1/25/2018#, #6/15/2018#, 97.975, 100, 1"),
    ("DMax", f'{DB}, "s", {CRIT}'), ("DMin", f'{DB}, "s", {CRIT}'),
    ("Dollar", "1234.567, 2"), ("Dollar", "-1234.567, -2"), ("DollarDe", "1.02, 16"), ("DollarFr", "1.125, 16"),
    ("DProduct", f'{DB}, "q", {CRIT}'), ("DStDev", f'{DB}, "q", {CRIT}'), ("DStDevP", f'{DB}, "q", {CRIT}'),
    ("DSum", f'{DB}, "q", {CRIT}'),
    ("Duration", "#7/1/2018#, #1/1/2048#, 0.08, 0.09, 2, 1"),
    ("DVar", f'{DB}, "q", {CRIT}'), ("DVarP", f'{DB}, "q", {CRIT}'),
    ("Effect", "0.0525, 4"), ("EncodeURL", '"a b&c"'),
    ("Erf", "0.745"), ("Erf", "0, 1.5"), ("Erf_Precise", "1"), ("ErfC", "1"), ("ErfC_Precise", "1"),
    ("Even", "1.5"), ("Even", "-1"), ("Expon_Dist", "0.2, 10, True"), ("ExponDist", "0.2, 10, False"),
    ("F_Dist", "15.2069, 6, 4, True"), ("F_Dist_RT", "15.2069, 6, 4"), ("F_Inv", "0.01, 6, 4"),
    ("F_Inv_RT", "0.01, 6, 4"), ("F_Test", f"{X}, {Y}"),
    ("Fact", "5"), ("Fact", "1.9"), ("FactDouble", "6"), ("FDist", "15.2069, 6, 4"),
    ("Filter", f'{A}, ws.Range("A1:A5") > 2'), ("FindB", '"b", "abc"'), ("FInv", "0.01, 6, 4"),
    ("Fisher", "0.75"), ("FisherInv", "0.972955"), ("Fixed", "1234.567, 1"), ("Fixed", "1234.567, -1, True"),
    ("Floor_Precise", "-3.2, -1"), ("Forecast_Linear", f"30, {Y}, {X}"),
    ("Frequency", f'{A}, Array(2, 4)'), ("FTest", f"{X}, {Y}"),
    ("Fv", "0.06/12, 10, -200, -500, 1"), ("FVSchedule", "1, Array(0.09, 0.11, 0.1)"),
    ("Gamma", "2.5"), ("Gamma_Dist", "10.00001131, 9, 2, False"), ("Gamma_Inv", "0.068094, 9, 2"),
    ("GammaDist", "10.00001131, 9, 2, True"), ("GammaInv", "0.068094, 9, 2"), ("GammaLn", "4"),
    ("GammaLn_Precise", "4"), ("Gauss", "2"), ("Gcd", "24, 36"), ("Gcd", "5, 0"),
    ("GeoMean", X), ("GeStep", "5, 4"), ("Growth", f"{Y}, {X}"), ("HarMean", X),
    ("Hex2Bin", '"F", 8'), ("Hex2Dec", '"A5"'), ("Hex2Dec", '"FFFFFFFF5B"'), ("Hex2Oct", '"F", 3'),
    ("HypGeom_Dist", "1, 4, 8, 20, True"), ("HypGeomDist", "1, 4, 8, 20"),
    ("ImAbs", '"5+12i"'), ("Imaginary", '"3+4i"'), ("ImArgument", '"3+4i"'), ("ImConjugate", '"3+4i"'),
    ("ImCos", '"1+i"'), ("ImCosh", '"4+3i"'), ("ImCot", '"4+3i"'), ("ImCsc", '"4+3i"'), ("ImCsch", '"4+3i"'),
    ("ImDiv", '"-238+240i", "10+24i"'), ("ImExp", '"1+i"'), ("ImLn", '"3+4i"'), ("ImLog10", '"3+4i"'),
    ("ImLog2", '"3+4i"'), ("ImPower", '"2+3i", 3'), ("ImProduct", '"3+4i", "5-3i"'), ("ImReal", '"6-9i"'),
    ("ImSec", '"4+3i"'), ("ImSech", '"4+3i"'), ("ImSin", '"4+3i"'), ("ImSinh", '"4+3i"'), ("ImSqrt", '"1+i"'),
    ("ImSub", '"13+4i", "5+3i"'), ("ImSum", '"3+4i", "5-3i"'), ("ImTan", '"4+3i"'),
    ("Intercept", f"{Y}, {X}"), ("IntRate", "#2/15/2008#, #5/15/2008#, 1000000, 1014420, 2"),
    ("Ipmt", "0.1/12, 1, 36, 8000"), ("Irr", "Array(-70000, 12000, 15000, 18000, 21000, 26000)"),
    ("IsEven", "-1"), ("IsEven", "2.5"), ("IsFormula", 'ws.Range("A1")'), ("ISO_Ceiling", "-4.3, 2"),
    ("IsOdd", "5"), ("IsoWeekNum", "#3/9/2012#"), ("Ispmt", "0.1/12, 1, 36, 8000000"),
    ("Kurt", "Array(3, 4, 5, 2, 3, 4, 5, 6, 4, 7)"), ("Lcm", "5, 2"), ("Lcm", "24, 36"),
    ("LinEst", f"{Y}, {X}"), ("LinEst", f"{Y}, {X}, True, True"), ("LogEst", f"{Y}, {X}"),
    ("LogInv", "0.039084, 3.5, 1.2"), ("LogNorm_Dist", "4, 3.5, 1.2, True"), ("LogNorm_Inv", "0.039084, 3.5, 1.2"),
    ("LogNormDist", "4, 3.5, 1.2"), ("Lookup", f"4.5, {A}, {B}"), ("Lookup", f'"c", {B}'),
    ("MaxIfs", f'{A}, {A}, "<4"'), ("MDeterm", "Array(Array(1, 3, 8), Array(4, 2, 1), Array(6, 4, 7))"),
    ("MDuration", "#1/1/2008#, #1/1/2016#, 0.08, 0.09, 2, 1"), ("MinIfs", f'{A}, {A}, ">2"'),
    ("MInverse", "Array(Array(4, -1), Array(2, 0))"), ("MIrr", "Array(-120000, 39000, 30000, 21000, 37000, 46000), 0.1, 0.12"),
    ("MMult", "Array(Array(1, 2), Array(3, 4)), Array(Array(5, 6), Array(7, 8))"),
    ("Mode", "Array(5.6, 4, 4, 3, 2, 4)"), ("Mode_Mult", "Array(1, 2, 3, 4, 3, 2, 1, 2, 3)"),
    ("Mode_Sngl", "Array(5.6, 4, 4, 3, 2, 4)"), ("MultiNomial", "2, 3, 4"), ("Munit", "2"),
    ("NegBinom_Dist", "10, 5, 0.25, True"), ("NegBinomDist", "10, 5, 0.25"),
    ("NetworkDays", "#10/1/2012#, #3/1/2013#"), ("NetworkDays", "#10/1/2012#, #3/1/2013#, Array(#11/22/2012#)"),
    ("NetworkDays_Intl", "#1/1/2006#, #1/31/2006#, 7"), ("NetworkDays_Intl", '#1/1/2006#, #1/31/2006#, "0010001"'),
    ("Nominal", "0.053543, 4"), ("Norm_Dist", "42, 40, 1.5, True"), ("Norm_Inv", "0.908789, 40, 1.5"),
    ("Norm_S_Dist", "1.333333, True"), ("Norm_S_Inv", "0.908789"), ("NormDist", "42, 40, 1.5, False"),
    ("NormInv", "0.908789, 40, 1.5"), ("NormSDist", "1.333333"), ("NormSInv", "0.908789"),
    ("NPer", "0.12/12, -100, -1000, 10000, 1"), ("Npv", "0.1, -10000, 3000, 4200, 6800"),
    ("NumberValue", '"2.500,27", ",", "."'),
    ("Oct2Bin", '"3", 3'), ("Oct2Dec", '"54"'), ("Oct2Hex", '"100", 4'), ("Odd", "1.5"), ("Odd", "-2"),
    ("OddFPrice", "#11/11/2008#, #3/1/2021#, #10/15/2008#, #3/1/2009#, 0.0785, 0.0625, 100, 2, 1"),
    ("OddFYield", "#11/11/2008#, #3/1/2021#, #10/15/2008#, #3/1/2009#, 0.0575, 84.5, 100, 2, 0"),
    ("OddLPrice", "#2/7/2008#, #6/15/2008#, #10/15/2007#, 0.0375, 0.0405, 100, 2, 0"),
    ("OddLYield", "#4/20/2008#, #6/15/2008#, #12/24/2007#, 0.0375, 99.875, 100, 2, 0"),
    ("PDuration", "0.025, 2000, 2200"), ("Pearson", f"{X}, {Y}"),
    ("Percentile", f"{A}, 0.3"), ("Percentile_Exc", f"{A}, 0.3"), ("Percentile_Inc", f"{A}, 0.3"),
    ("PercentRank", f"{A}, 2"), ("PercentRank_Exc", f"{A}, 2, 4"), ("PercentRank_Inc", f"{A}, 2.5"),
    ("Permut", "100, 3"), ("Permutationa", "3, 2"), ("Phi", "0.75"),
    ("Pmt", "0.08/12, 10, 10000"), ("Pmt", "0.08/12, 10, 10000, 0, 1"),
    ("Poisson", "2, 5, True"), ("Poisson_Dist", "2, 5, False"), ("Ppmt", "0.1/12, 1, 24, 2000"),
    ("Price", "#2/15/2008#, #11/15/2017#, 0.0575, 95.04287, 100, 2, 0"),
    ("PriceDisc", "#2/16/2008#, #3/1/2008#, 0.0525, 100, 2"),
    ("PriceMat", "#2/15/2008#, #4/13/2008#, #11/11/2007#, 0.061, 0.061, 0"),
    ("Prob", "Array(0, 1, 2, 3), Array(0.2, 0.3, 0.1, 0.4), 2"),
    ("Pv", "0.08/12, 12*20, 500, , 0"), ("Quartile", f"{A}, 1"), ("Quartile_Exc", "Array(6, 7, 15, 36, 39, 40, 41, 42, 43, 47, 49), 1"),
    ("Quartile_Inc", f"{A}, 3"), ("Quotient", "5, 2"), ("Quotient", "-10, 3"), ("Radians", "270"),
    ("RandArray", "1, 1, 5, 5, True"), ("RandBetween", "3, 3"),
    ("Rank", f"3, {A}"), ("Rank", f"3, {A}, 1"), ("Rank_Avg", f'2, ws.Range("F1:F3")'), ("Rank_Eq", f"3, {A}"),
    ("Rate", "4*12, -200, 8000"), ("Received", "#2/15/2008#, #5/15/2008#, 1000000, 0.0575, 2"),
    ("ReplaceB", '"abcdef", 2, 3, "X"'), ("Roman", "499"), ("Roman", "499, 4"), ("Rri", "96, 10000, 11000"),
    ("RSq", f"{Y}, {X}"), ("SearchB", '"C", "abc"'), ("Sec", "45"), ("Sech", "45"),
    ("Sequence", "3"), ("Sequence", "2, 3, 10, -1"), ("SeriesSum", "0.785398163, 0, 2, Array(1, -0.5, 0.041666667)"),
    ("Single", f"{A}"), ("Sinh", "1"), ("Skew", "Array(3, 4, 5, 2, 3, 4, 5, 6, 4, 7)"),
    ("Skew_p", "Array(3, 4, 5, 2, 3, 4, 5, 6, 4, 7)"), ("Sln", "30000, 7500, 10"), ("Slope", f"{Y}, {X}"),
    ("Sort", f"{F}"), ("Sort", f"{F}, 1, -1"), ("SortBy", f'{B}, {F}, 1'),
    ("SqrtPi", "2"), ("Standardize", "42, 40, 1.5"), ("StEyx", f"{Y}, {X}"),
    ("SumX2MY2", f"{X}, {Y}"), ("SumX2PY2", f"{X}, {Y}"), ("SumXMY2", f"{X}, {Y}"), ("Syd", "30000, 7500, 10, 1"),
    ("T_Dist", "60, 1, True"), ("T_Dist_2T", "1.959999998, 60"), ("T_Dist_RT", "1.959999998, 60"),
    ("T_Inv", "0.75, 2"), ("T_Inv_2T", "0.546449, 60"), ("T_Test", f"{X}, {Y}, 2, 1"), ("Tanh", "0.5"),
    ("TBillEq", "#3/31/2008#, #6/1/2008#, 0.0914"), ("TBillPrice", "#3/31/2008#, #6/1/2008#, 0.09"),
    ("TBillYield", "#3/31/2008#, #6/1/2008#, 98.45"), ("TDist", "1.959999998, 60, 2"), ("TInv", "0.05464, 60"),
    ("Trend", f"{Y}, {X}"), ("TrimMean", "Array(4, 5, 6, 7, 2, 3, 4, 5, 1, 2, 3), 0.2"),
    ("TTest", f"{X}, {Y}, 2, 2"), ("Unichar", "66"), ("Unicode", '"B"'), ("Unique", "Array(1, 2, 1, 3)"),
    ("Unique", f'ws.Range("F1:F3")'), ("USDollar", "1234.567, 2"), ("ValueToText", '"a"'), ("ValueToText", '"a", 1'),
    ("Vdb", "2400, 300, 10*365, 0, 1"), ("WeekNum", "#3/9/2012#"), ("WeekNum", "#3/9/2012#, 2"),
    ("Weibull", "105, 20, 100, True"), ("Weibull_Dist", "105, 20, 100, False"),
    ("WorkDay", "#10/1/2008#, 151"), ("WorkDay_Intl", "#1/1/2012#, 30, 0"), ("WorkDay_Intl", "#1/1/2012#, 30, 11"),
    ("Xirr", "Array(-10000, 2750, 4250, 3250, 2750), Array(#1/1/2008#, #3/1/2008#, #10/30/2008#, #2/15/2009#, #4/1/2009#)"),
    ("XMatch", f'"c", {B}'), ("XMatch", f"3.5, {A}, 1"),
    # Excel's type library gives WorksheetFunction.Xnpv two arguments, the rate not among them.
    ("Xnpv", "Array(-10000, 2750, 4250, 3250, 2750), Array(#1/1/2008#, #3/1/2008#, #10/30/2008#, #2/15/2009#, #4/1/2009#)"),
    ("YearFrac", f"{S}, #7/30/2020#"), ("YearFrac", f"{S}, #7/30/2020#, 1"),
    ("YieldDisc", "#2/16/2008#, #3/1/2008#, 99.795, 100, 2"),
    ("YieldMat", "#3/15/2008#, #11/3/2008#, #11/8/2007#, 0.0625, 100.0123, 0"),
    ("Z_Test", "Array(3, 6, 7, 8, 6, 5, 4, 2, 1, 9), 4"), ("ZTest", "Array(3, 6, 7, 8, 6, 5, 4, 2, 1, 9), 4, 1"),
]


def module() -> str:
    batches: list[str] = []
    for start in range(0, len(CASES), BATCH):
        lines = [f"Private Function Batch{start // BATCH}(ws As Object) As String", "Dim out As String, v As Variant",
                 "On Error Resume Next"]
        for name, args in CASES[start:start + BATCH]:
            call = f"({args})" if args else ""
            for target in (f"Application.WorksheetFunction.{name}{call}", f"Application.{name}{call}"):
                lines += ["Err.Clear", "v = Empty", f"v = {target}",
                          'If Err.Number <> 0 Then out = out & "E" & Err.Number & "^" Else out = out & Show(v) & "^"']
            lines.append('out = out & "|"')
        lines += ["On Error GoTo 0", f"Batch{start // BATCH} = out", "End Function"]
        batches.append("\n".join(lines))
    probe = ["Public Function Probe() As String", "Dim wb As Object, ws As Object, out As String",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)",
             "Set ws = wb.Worksheets(1)", "Fill ws",
             *[f"out = out & Batch{index}(ws)" for index in range(len(batches))],
             "wb.Close False", "Probe = out", "End Function"]
    return "\n".join([HELPER, SETUP, *batches, *probe]) + "\n"


def main() -> None:
    code = module()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(code, "Probe", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error} {result.dialogs}"
    answers = [part.split("^")[:2] for part in str(result.value).split("|")[: len(CASES)]]
    record = {"helper": HELPER, "setup": SETUP, "cases": [
        {"function": name, "arguments": args, "worksheet_function": wf, "application": app}
        for (name, args), (wf, app) in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for (name, args), (wf, app) in zip(CASES, answers, strict=True):
        print(f"{name}({args}) => {wf} | {app}")


if __name__ == "__main__":
    main()
