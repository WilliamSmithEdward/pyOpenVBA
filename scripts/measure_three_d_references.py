"""What formulas give for a 3D reference, one that spans sheets: ThreeA:ThreeB!A1.

Each probe adds a sheet for the formula and, after it, two sheets ThreeA
and ThreeB holding numbers, text and a logical in A1:B2 and a number in C3.
It writes one formula to C3 of the first sheet and records the value
(CStr, so an error shows as "Error 2023") and the formula Excel reads back,
then deletes the sheets.

    python scripts/measure_three_d_references.py [references|functions]

writes tests/fixtures/three_d_references.json, the reference standing alone,
with operators, in brackets and by name, and tests/fixtures/three_d_functions.json,
the reference given to each function; tests/test_excel_three_d_references.py
replays both. With an argument it measures only that one.
"""
from __future__ import annotations

import sys
from pathlib import Path

from vba_probe_fixture import record

FIXTURES = Path(__file__).resolve().parents[1] / 'tests/fixtures'
#: The reference itself, written with Range.Formula.
REFERENCES = {
    'alone': '=ThreeA:ThreeB!A1',
    'alone_area': '=ThreeA:ThreeB!A1:B2',
    'alone_column': '=ThreeA:ThreeB!C:C',
    'alone_one_sheet': '=ThreeA:ThreeA!A1',
    'alone_reversed': '=ThreeB:ThreeA!A1',
    'plus_one': '=ThreeA:ThreeB!A1+1',
    'negated': '=-ThreeA:ThreeB!A1',
    'if_returns_it': '=IF(TRUE,ThreeA:ThreeB!A1)',
    'choose_returns_it': '=CHOOSE(1,ThreeA:ThreeB!A1)',
    'iferror': '=IFERROR(ThreeA:ThreeB!A1,"caught")',
    'let': '=LET(x,ThreeA:ThreeB!A1,SUM(x))',
    'name_summed': '=SUM(BothSheets)',
    'name_alone': '=BothSheets',
    'quoted': "=SUM('ThreeA:ThreeB'!A1)",
    'sum_reversed': '=SUM(ThreeB:ThreeA!A1:B2)',
    'sum_one_sheet': '=SUM(ThreeA:ThreeA!A1:B2)',
    'sum_with_others': '=SUM(ThreeA:ThreeB!A1,5,ThreeA!B1)',
    'sum_times_one': '=SUM(ThreeA:ThreeB!A1*1)',
    'union': '=SUM((ThreeA:ThreeB!A1,ThreeA!B1))',
    'intersection': '=SUM(ThreeA:ThreeB!A1:B2 ThreeA:ThreeB!A1)',
    'range_to_a_cell': '=SUM(ThreeA:ThreeB!A1:ThreeB!B1)',
    'range_on_from_an_area': '=SUM(ThreeA:ThreeB!A1:B1:C3)',
    'bracketed': '=(ThreeA:ThreeB!A1)',
    'bracketed_summed': '=SUM((ThreeA:ThreeB!A1))',
    'bracketed_negated': '=-(ThreeA:ThreeB!A1)',
    'bracketed_after_operator': '=1+(ThreeA:ThreeB!A1)',
    'bracketed_twice': '=SUM(((ThreeA:ThreeB!A1)))',
    'bracketed_area': '=SUM((ThreeA:ThreeB!A1:B2))',
    'bracketed_beside_another': '=SUM(ThreeA:ThreeB!A1,(ThreeA:ThreeB!A1))',
    'bracketed_in_if': '=IF(TRUE,(ThreeA:ThreeB!A1))',
    'bracketed_one_sheet': '=(ThreeA!A1)',
    'quoted_bracketed': "=SUM(('ThreeA:ThreeB'!A1))",
    'bracketed_after_one_plus': '=(1+ThreeA:ThreeB!A1)',
    'bracketed_before_plus_one': '=(ThreeA:ThreeB!A1+1)',
    'bracketed_union_second': '=SUM((ThreeA!B1,ThreeA:ThreeB!A1))',
    'bracketed_around_sum': '=(SUM(ThreeA:ThreeB!A1))',
    'bracketed_let': '=LET(x,(ThreeA:ThreeB!A1),SUM(x))',
    'sum_of_if': '=SUM(IF(TRUE,ThreeA:ThreeB!A1))',
    'sum_of_choose': '=SUM(CHOOSE(1,ThreeA:ThreeB!A1))',
    'spilled': '=SUM(ThreeA:ThreeB!A1#)',
    'indirect': '=SUM(INDIRECT("ThreeA:ThreeB!A1"))',
    'indirect_one_sheet': '=SUM(INDIRECT("ThreeA:ThreeA!A1"))',
    'indirect_quoted': '=SUM(INDIRECT("\'ThreeA:ThreeB\'!A1"))',
}
#: The reference given to a function, written with Range.Formula.
FUNCTIONS = {
    'sum': '=SUM(ThreeA:ThreeB!A1)',
    'sum_area': '=SUM(ThreeA:ThreeB!A1:B2)',
    'average': '=AVERAGE(ThreeA:ThreeB!A1:B2)',
    'averagea': '=AVERAGEA(ThreeA:ThreeB!A1:B2)',
    'count': '=COUNT(ThreeA:ThreeB!A1:B2)',
    'counta': '=COUNTA(ThreeA:ThreeB!A1:B2)',
    'max': '=MAX(ThreeA:ThreeB!A1:B2)',
    'maxa': '=MAXA(ThreeA:ThreeB!A1:B2)',
    'min': '=MIN(ThreeA:ThreeB!A1:B2)',
    'mina': '=MINA(ThreeA:ThreeB!A1:B2)',
    'product': '=PRODUCT(ThreeA:ThreeB!A1:B1)',
    'stdev': '=STDEV(ThreeA:ThreeB!A1:B1)',
    'stdev_s': '=STDEV.S(ThreeA:ThreeB!A1:B1)',
    'stdev_p': '=STDEV.P(ThreeA:ThreeB!A1:B1)',
    'stdeva': '=STDEVA(ThreeA:ThreeB!A1:B2)',
    'stdevpa': '=STDEVPA(ThreeA:ThreeB!A1:B2)',
    'stdevp': '=STDEVP(ThreeA:ThreeB!A1:B1)',
    'var': '=VAR(ThreeA:ThreeB!A1:B1)',
    'var_s': '=VAR.S(ThreeA:ThreeB!A1:B1)',
    'var_p': '=VAR.P(ThreeA:ThreeB!A1:B1)',
    'vara': '=VARA(ThreeA:ThreeB!A1:B2)',
    'varpa': '=VARPA(ThreeA:ThreeB!A1:B2)',
    'varp': '=VARP(ThreeA:ThreeB!A1:B1)',
    'sumsq': '=SUMSQ(ThreeA:ThreeB!A1)',
    'sumsq_area': '=SUMSQ(ThreeA:ThreeB!A1:B1)',
    'large': '=LARGE(ThreeA:ThreeB!A1,1)',
    'small': '=SMALL(ThreeA:ThreeB!A1,1)',
    'median': '=MEDIAN(ThreeA:ThreeB!A1)',
    'rank': '=RANK(1,ThreeA:ThreeB!A1)',
    'rank_eq': '=RANK.EQ(2,ThreeA:ThreeB!A1:B1)',
    'rank_avg': '=RANK.AVG(2,ThreeA:ThreeB!A1:B1)',
    'avedev': '=AVEDEV(ThreeA:ThreeB!A1:B1)',
    'devsq': '=DEVSQ(ThreeA:ThreeB!A1:B1)',
    'geomean': '=GEOMEAN(ThreeA:ThreeB!A1:B1)',
    'harmean': '=HARMEAN(ThreeA:ThreeB!A1:B1)',
    'kurt': '=KURT(ThreeA:ThreeB!A1:B1)',
    'skew': '=SKEW(ThreeA:ThreeB!A1:B1)',
    'skew_p': '=SKEW.P(ThreeA:ThreeB!A1:B1)',
    'percentile': '=PERCENTILE(ThreeA:ThreeB!A1:B1,0.5)',
    'percentile_inc': '=PERCENTILE.INC(ThreeA:ThreeB!A1:B1,0.5)',
    'percentile_exc': '=PERCENTILE.EXC(ThreeA:ThreeB!A1:B1,0.5)',
    'quartile': '=QUARTILE(ThreeA:ThreeB!A1:B1,1)',
    'quartile_inc': '=QUARTILE.INC(ThreeA:ThreeB!A1:B1,1)',
    'quartile_exc': '=QUARTILE.EXC(ThreeA:ThreeB!A1:B1,1)',
    'trimmean': '=TRIMMEAN(ThreeA:ThreeB!A1:B1,0.5)',
    'percentrank': '=PERCENTRANK(ThreeA:ThreeB!A1:B1,2)',
    'percentrank_inc': '=PERCENTRANK.INC(ThreeA:ThreeB!A1:B1,2)',
    'percentrank_exc': '=PERCENTRANK.EXC(ThreeA:ThreeB!A1:B1,2)',
    'frequency': '=SUM(FREQUENCY(ThreeA:ThreeB!A1:B1,5))',
    'npv': '=NPV(0.1,ThreeA:ThreeB!A1:B1)',
    'and': '=AND(ThreeA:ThreeB!A1)',
    'or': '=OR(ThreeA:ThreeB!A1)',
    'xor': '=XOR(ThreeA:ThreeB!A1)',
    'concat': '=CONCAT(ThreeA:ThreeB!A1)',
    'textjoin': '=TEXTJOIN(",",TRUE,ThreeA:ThreeB!A1)',
    'sheet': '=SHEET(ThreeA:ThreeB!A1)',
    'sheets': '=SHEETS(ThreeA:ThreeB!A1)',
    'sumproduct': '=SUMPRODUCT(ThreeA:ThreeB!A1)',
    'mode': '=MODE(ThreeA:ThreeB!A1:B1,1)',
    'mode_sngl': '=MODE.SNGL(ThreeA:ThreeB!A1:B1,1)',
    'sumx2my2': '=SUMX2MY2(ThreeA:ThreeB!A1:B1,ThreeA:ThreeB!A1:B1)',
    'sumx2py2': '=SUMX2PY2(ThreeA:ThreeB!A1:B1,ThreeA:ThreeB!A1:B1)',
    'sumxmy2': '=SUMXMY2(ThreeA:ThreeB!A1:B1,ThreeA:ThreeB!A1:B1)',
    'correl': '=CORREL(ThreeA:ThreeB!A1:B1,ThreeA:ThreeB!A1:B1)',
    'pearson': '=PEARSON(ThreeA:ThreeB!A1:B1,ThreeA:ThreeB!A1:B1)',
    'rsq': '=RSQ(ThreeA:ThreeB!A1:B1,ThreeA:ThreeB!A1:B1)',
    'slope': '=SLOPE(ThreeA:ThreeB!A1:B1,ThreeA:ThreeB!A1:B1)',
    'intercept': '=INTERCEPT(ThreeA:ThreeB!A1:B1,ThreeA:ThreeB!A1:B1)',
    'covar': '=COVAR(ThreeA:ThreeB!A1:B1,ThreeA:ThreeB!A1:B1)',
    'covariance_s': '=COVARIANCE.S(ThreeA:ThreeB!A1:B1,ThreeA:ThreeB!A1:B1)',
    'forecast': '=FORECAST(1,ThreeA:ThreeB!A1:B1,ThreeA:ThreeB!A1:B1)',
    'steyx': '=STEYX(ThreeA:ThreeB!A1:B1,ThreeA:ThreeB!A1:B1)',
    't_test': '=T.TEST(ThreeA:ThreeB!A1:B1,ThreeA:ThreeB!A1:B1,2,1)',
    'mmult': '=SUM(MMULT(ThreeA:ThreeB!A1,1))',
    'mdeterm': '=MDETERM(ThreeA:ThreeB!A1)',
    'minverse': '=SUM(MINVERSE(ThreeA:ThreeB!A1))',
    'transpose': '=SUM(TRANSPOSE(ThreeA:ThreeB!A1:B1))',
    'trend': '=SUM(TREND(ThreeA:ThreeB!A1:B1))',
    'linest': '=SUM(LINEST(ThreeA:ThreeB!A1:B1))',
    'arraytotext': '=ARRAYTOTEXT(ThreeA:ThreeB!A1:B1)',
    'countblank': '=COUNTBLANK(ThreeA:ThreeB!A1:C3)',
    'countblank_one_sheet': '=COUNTBLANK(ThreeA:ThreeA!A1:C3)',
    'countif': '=COUNTIF(ThreeA:ThreeB!A1,">0")',
    'countifs': '=COUNTIFS(ThreeA:ThreeB!A1,">0")',
    'sumif': '=SUMIF(ThreeA:ThreeB!A1,">0")',
    'sumif_sum_range': '=SUMIF(ThreeA!A1:B1,">0",ThreeA:ThreeB!A1:B1)',
    'sumifs': '=SUMIFS(ThreeA:ThreeB!A1,ThreeA:ThreeB!A1,">0")',
    'averageif': '=AVERAGEIF(ThreeA:ThreeB!A1,">0")',
    'maxifs': '=MAXIFS(ThreeA:ThreeB!A1,ThreeA:ThreeB!A1,">0")',
    'aggregate': '=AGGREGATE(9,0,ThreeA:ThreeB!A1)',
    'subtotal': '=SUBTOTAL(9,ThreeA:ThreeB!A1)',
    'index': '=INDEX(ThreeA:ThreeB!A1:B2,1,1)',
    'index_area': '=INDEX(ThreeA:ThreeB!A1:B2,1,1,2)',
    'vlookup': '=VLOOKUP(1,ThreeA:ThreeB!A1:B2,2,FALSE)',
    'hlookup': '=HLOOKUP(1,ThreeA:ThreeB!A1:B2,2,FALSE)',
    'match': '=MATCH(1,ThreeA:ThreeB!A1,0)',
    'lookup': '=LOOKUP(1,ThreeA:ThreeB!A1:B1)',
    'xlookup': '=XLOOKUP(1,ThreeA:ThreeB!A1:B1,ThreeA:ThreeB!A1:B1)',
    'xmatch': '=XMATCH(1,ThreeA:ThreeB!A1:B1)',
    'offset': '=SUM(OFFSET(ThreeA:ThreeB!A1,0,0))',
    'rows': '=ROWS(ThreeA:ThreeB!A1:B2)',
    'columns': '=COLUMNS(ThreeA:ThreeB!A1:B2)',
    'areas': '=AREAS(ThreeA:ThreeB!A1)',
    'row': '=ROW(ThreeA:ThreeB!A1)',
    'column': '=COLUMN(ThreeA:ThreeB!A1)',
    'cell_address': '=CELL("address",ThreeA:ThreeB!A1)',
    'formulatext': '=FORMULATEXT(ThreeA:ThreeB!A1)',
    'isformula': '=ISFORMULA(ThreeA:ThreeB!A1)',
    'gcd': '=GCD(ThreeA:ThreeB!A1:B1)',
    'lcm': '=LCM(ThreeA:ThreeB!A1:B1)',
    'irr': '=IRR(ThreeA:ThreeB!A1:B1)',
    'mirr': '=MIRR(ThreeA:ThreeB!A1:B1,0.1,0.1)',
    'xnpv': '=XNPV(0.1,ThreeA:ThreeB!A1:B1,ThreeA:ThreeB!A1:B1)',
    'fvschedule': '=FVSCHEDULE(1,ThreeA:ThreeB!A1:B1)',
    'multinomial': '=MULTINOMIAL(ThreeA:ThreeB!A1:B1)',
    'imsum': '=IMSUM(ThreeA:ThreeB!A1:B1)',
    'seriessum': '=SERIESSUM(1,1,1,ThreeA:ThreeB!A1:B1)',
    'workday': '=WORKDAY(1,1,ThreeA:ThreeB!A1)',
    'networkdays': '=NETWORKDAYS(1,10,ThreeA:ThreeB!A1)',
    'isref': '=ISREF(ThreeA:ThreeB!A1)',
    'iserror': '=ISERROR(ThreeA:ThreeB!A1)',
    'iserr': '=ISERR(ThreeA:ThreeB!A1)',
    'isblank': '=ISBLANK(ThreeA:ThreeB!A1)',
    'isnumber': '=ISNUMBER(ThreeA:ThreeB!A1)',
    'type': '=TYPE(ThreeA:ThreeB!A1)',
    'error_type': '=ERROR.TYPE(ThreeA:ThreeB!A1)',
    'n': '=N(ThreeA:ThreeB!A1)',
    't': '=T(ThreeA:ThreeB!A1)',
    'abs': '=ABS(ThreeA:ThreeB!A1)',
    'text': '=TEXT(ThreeA:ThreeB!A1,"0")',
    'concatenate': '=CONCATENATE(ThreeA:ThreeB!A1)',
    'averageifs': '=AVERAGEIFS(ThreeA:ThreeB!A1,ThreeA:ThreeB!A1,">0")',
    'minifs': '=MINIFS(ThreeA:ThreeB!A1,ThreeA:ThreeB!A1,">0")',
    'chisq_test': '=CHISQ.TEST(ThreeA:ThreeB!A1:B1,ThreeA:ThreeB!A1:B1)',
    'chitest': '=CHITEST(ThreeA:ThreeB!A1:B1,ThreeA:ThreeB!A1:B1)',
    'covariance_p': '=COVARIANCE.P(ThreeA:ThreeB!A1:B1,ThreeA:ThreeB!A1:B1)',
    'f_test': '=F.TEST(ThreeA:ThreeB!A1:B1,ThreeA:ThreeB!A1:B1)',
    'ftest': '=FTEST(ThreeA:ThreeB!A1:B1,ThreeA:ThreeB!A1:B1)',
    'forecast_linear': '=FORECAST.LINEAR(1,ThreeA:ThreeB!A1:B1,ThreeA:ThreeB!A1:B1)',
    'ttest': '=TTEST(ThreeA:ThreeB!A1:B1,ThreeA:ThreeB!A1:B1,2,1)',
    'z_test': '=Z.TEST(ThreeA:ThreeB!A1:B1,1)',
    'ztest': '=ZTEST(ThreeA:ThreeB!A1:B1,1)',
    'prob': '=PROB(ThreeA:ThreeB!A1:B1,ThreeA:ThreeB!A1:B1,1)',
    'growth': '=SUM(GROWTH(ThreeA:ThreeB!A1:B1))',
    'logest': '=SUM(LOGEST(ThreeA:ThreeB!A1:B1))',
    'improduct': '=IMPRODUCT(ThreeA:ThreeB!A1:B1)',
    'xirr': '=XIRR(ThreeA:ThreeB!A1:B1,ThreeA:ThreeB!A1:B1)',
    'workday_intl': '=WORKDAY.INTL(1,1,1,ThreeA:ThreeB!A1)',
    'networkdays_intl': '=NETWORKDAYS.INTL(1,10,1,ThreeA:ThreeB!A1)',
    'getpivotdata': '=GETPIVOTDATA("x",ThreeA:ThreeB!A1)',
    'daverage': '=DAVERAGE(ThreeA:ThreeB!A1:B2,1,ThreeA!A1:A2)',
    'dcount': '=DCOUNT(ThreeA:ThreeB!A1:B2,1,ThreeA!A1:A2)',
    'dcounta': '=DCOUNTA(ThreeA:ThreeB!A1:B2,1,ThreeA!A1:A2)',
    'dget': '=DGET(ThreeA:ThreeB!A1:B2,1,ThreeA!A1:A2)',
    'dmax': '=DMAX(ThreeA:ThreeB!A1:B2,1,ThreeA!A1:A2)',
    'dmin': '=DMIN(ThreeA:ThreeB!A1:B2,1,ThreeA!A1:A2)',
    'dproduct': '=DPRODUCT(ThreeA:ThreeB!A1:B2,1,ThreeA!A1:A2)',
    'dstdev': '=DSTDEV(ThreeA:ThreeB!A1:B2,1,ThreeA!A1:A2)',
    'dstdevp': '=DSTDEVP(ThreeA:ThreeB!A1:B2,1,ThreeA!A1:A2)',
    'dsum': '=DSUM(ThreeA:ThreeB!A1:B2,1,ThreeA!A1:A2)',
    'dvar': '=DVAR(ThreeA:ThreeB!A1:B2,1,ThreeA!A1:A2)',
    'dvarp': '=DVARP(ThreeA:ThreeB!A1:B2,1,ThreeA!A1:A2)',
    'dsum_criteria': '=DSUM(ThreeA!A1:B2,1,ThreeA:ThreeB!A1:A2)',
    'concat_order': '=CONCAT(ThreeA:ThreeB!A1:B2)',
    'textjoin_order': '=TEXTJOIN(",",FALSE,ThreeA:ThreeB!A1:C3)',
    'and_with_text': '=AND(ThreeA:ThreeB!A1:B2)',
    'concat_union': '=CONCAT((ThreeA!A1,ThreeA!B1))',
    'textjoin_union': '=TEXTJOIN(",",TRUE,(ThreeA!A1,ThreeA!B1))',
    'index_union': '=INDEX((ThreeA!A1:B2,ThreeA!A2:B2),1,1,2)',
    'areas_union': '=AREAS((ThreeA!A1,ThreeA!B1))',
    'rows_union': '=ROWS((ThreeA!A1,ThreeA!B1))',
    'sumproduct_union': '=SUMPRODUCT((ThreeA!A1,ThreeA!B1))',
    'isref_union': '=ISREF((ThreeA!A1,ThreeA!B1))',
}
#: Written with Range.Formula2, as dynamic-array formulas.
FORMULAS2 = {
    'alone_formula2': (REFERENCES, '=ThreeA:ThreeB!A1'),
    'alone_area_formula2': (REFERENCES, '=ThreeA:ThreeB!A1:B2'),
    'sum_times_one_formula2': (REFERENCES, '=SUM(ThreeA:ThreeB!A1*1)'),
    'vstack_formula2': (FUNCTIONS, '=SUM(VSTACK(ThreeA:ThreeB!A1:B1))'),
    'vstack_rows_formula2': (FUNCTIONS, '=ROWS(VSTACK(ThreeA:ThreeB!A1:B2))'),
    'vstack_mixed_formula2': (FUNCTIONS, '=ROWS(VSTACK(ThreeA:ThreeB!A1:B1,ThreeA!A1:B1))'),
    'vstack_union_formula2': (FUNCTIONS, '=ROWS(VSTACK((ThreeA!A1:B1,ThreeA!A2:B2)))'),
    'hstack_formula2': (FUNCTIONS, '=COLUMNS(HSTACK(ThreeA:ThreeB!A1:B1))'),
    'tocol_formula2': (FUNCTIONS, '=SUM(TOCOL(ThreeA:ThreeB!A1:B1))'),
    'torow_formula2': (FUNCTIONS, '=SUM(TOROW(ThreeA:ThreeB!A1:B1))'),
    'filter_formula2': (FUNCTIONS, '=SUM(FILTER(ThreeA:ThreeB!A1:B1,TRUE))'),
    'sort_formula2': (FUNCTIONS, '=SUM(SORT(ThreeA:ThreeB!A1:B1))'),
    'sortby_formula2': (FUNCTIONS, '=SUM(SORTBY(ThreeA:ThreeB!A1:B1,ThreeA:ThreeB!A1:B1))'),
    'unique_formula2': (FUNCTIONS, '=SUM(UNIQUE(ThreeA:ThreeB!A1:B1))'),
    'take_formula2': (FUNCTIONS, '=SUM(TAKE(ThreeA:ThreeB!A1:B1,1))'),
    'drop_formula2': (FUNCTIONS, '=SUM(DROP(ThreeA:ThreeB!A1:B1,1))'),
    'expand_formula2': (FUNCTIONS, '=SUM(EXPAND(ThreeA:ThreeB!A1:B1,1))'),
    'choosecols_formula2': (FUNCTIONS, '=SUM(CHOOSECOLS(ThreeA:ThreeB!A1:B1,1))'),
    'chooserows_formula2': (FUNCTIONS, '=SUM(CHOOSEROWS(ThreeA:ThreeB!A1:B1,1))'),
    'wraprows_formula2': (FUNCTIONS, '=SUM(WRAPROWS(ThreeA:ThreeB!A1:B1,1))'),
    'map_formula2': (FUNCTIONS, '=SUM(MAP(ThreeA:ThreeB!A1:B1,LAMBDA(x,x)))'),
    'byrow_formula2': (FUNCTIONS, '=SUM(BYROW(ThreeA:ThreeB!A1:B1,LAMBDA(r,SUM(r))))'),
    'reduce_formula2': (FUNCTIONS, '=REDUCE(0,ThreeA:ThreeB!A1:B1,LAMBDA(a,x,a+x))'),
    'groupby_formula2': (FUNCTIONS, '=ROWS(GROUPBY(ThreeA:ThreeB!A1,ThreeA:ThreeB!A1,SUM))'),
    'mode_mult_formula2': (FUNCTIONS, '=SUM(MODE.MULT(ThreeA:ThreeB!A1:B1,1))'),
    'wrapcols_formula2': (FUNCTIONS, '=SUM(WRAPCOLS(ThreeA:ThreeB!A1:B1,1))'),
    'bycol_formula2': (FUNCTIONS, '=SUM(BYCOL(ThreeA:ThreeB!A1:B1,LAMBDA(c,SUM(c))))'),
    'scan_formula2': (FUNCTIONS, '=SUM(SCAN(0,ThreeA:ThreeB!A1:B1,LAMBDA(a,x,a+x)))'),
    'pivotby_formula2': (FUNCTIONS, '=ROWS(PIVOTBY(ThreeA:ThreeB!A1,ThreeA:ThreeB!A1,ThreeA:ThreeB!A1,SUM))'),
    'percentof_formula2': (FUNCTIONS, '=PERCENTOF(ThreeA:ThreeB!A1,ThreeA:ThreeB!A1:B1)'),
    'trimrange_formula2': (FUNCTIONS, '=SUM(TRIMRANGE(ThreeA:ThreeB!A1:B2))'),
    'anchorarray_formula2': (FUNCTIONS, '=SUM(ANCHORARRAY(ThreeA:ThreeB!A1))'),
    'at_sign_formula2': (FUNCTIONS, '=@ThreeA:ThreeB!A1'),
    'textafter_formula2': (FUNCTIONS, '=TEXTAFTER("1x2",ThreeA:ThreeB!B2)'),
    'textbefore_formula2': (FUNCTIONS, '=TEXTBEFORE("1x2",ThreeA:ThreeB!B2)'),
    'textsplit_formula2': (FUNCTIONS, '=ROWS(TEXTSPLIT("1x2",ThreeA:ThreeB!B2))'),
    'vstack_order_formula2': (FUNCTIONS, '=TEXTJOIN(",",TRUE,VSTACK(ThreeA:ThreeB!A1:B2))'),
    'hstack_order_formula2': (FUNCTIONS, '=TEXTJOIN(",",TRUE,HSTACK(ThreeA:ThreeB!A1:B2))'),
    'hstack_union_formula2': (FUNCTIONS, '=COLUMNS(HSTACK((ThreeA!A1:B1,ThreeA!A2:B2)))'),
    'tocol_order_formula2': (FUNCTIONS, '=TEXTJOIN(",",TRUE,TOCOL(ThreeA:ThreeB!A1:B2))'),
    'tocol_by_column_formula2': (FUNCTIONS, '=TEXTJOIN(",",TRUE,TOCOL(ThreeA:ThreeB!A1:B2,0,TRUE))'),
    'torow_order_formula2': (FUNCTIONS, '=TEXTJOIN(",",TRUE,TOROW(ThreeA:ThreeB!A1:B2))'),
    'torow_by_column_formula2': (FUNCTIONS, '=TEXTJOIN(",",TRUE,TOROW(ThreeA:ThreeB!A1:B2,0,TRUE))'),
    'tocol_union_formula2': (FUNCTIONS, '=ROWS(TOCOL((ThreeA!A1:B1,ThreeA!A2:B2)))'),
    'tocol_skipping_formula2': (FUNCTIONS, '=ROWS(TOCOL(ThreeA:ThreeB!A1:C3,1))'),
    'n_formula2': (FUNCTIONS, '=N(ThreeA:ThreeB!A1)'),
    'isref_formula2': (FUNCTIONS, '=ISREF(ThreeA:ThreeB!A1)'),
}
#: With a name ThreeA, for ThreeB!B1, beside the sheet of that name: in brackets of its own, ThreeA:ThreeB!A1 reads
#: the name.
NAMED = {
    'bracketed_beside_a_name': '=SUM((ThreeA:ThreeB!A1))',
    'beside_a_name': '=SUM(ThreeA:ThreeB!A1)',
}
#: With sheets named like cells, Q1 and Q4, and like names of other kinds.
NAMED_SHEETS = {
    'cell_like_sheets': (('Q1', 'Q4'), '=SUM(Q1:Q4!A1)'),
    'cell_like_sheets_bracketed': (('Q1', 'Q4'), '=SUM((Q1:Q4!A1))'),
    'cell_like_sheets_alone': (('Q1', 'Q4'), '=Q1:Q4!A1'),
    'cell_like_sheets_quoted': (('Q1', 'Q4'), "=SUM('Q1:Q4'!A1)"),
    'column_like_sheets': (('AB', 'CD'), '=SUM(AB:CD!A1)'),
    'column_like_sheets_bracketed': (('AB', 'CD'), '=SUM((AB:CD!A1))'),
    'cell_like_last_sheet': (('ThreeA', 'Q4'), '=SUM(ThreeA:Q4!A1)'),
    'reversed_bracketed': (('ThreeA', 'ThreeB'), '=SUM((ThreeB:ThreeA!A1))'),
    'spaced_sheets': (('Three A', 'Three B'), "=SUM('Three A:Three B'!A1)"),
    'spaced_sheets_reversed': (('Three A', 'Three B'), "=SUM('Three B:Three A'!A1)"),
    'one_end_spaced': (('ThreeA', 'Three B'), "=SUM('ThreeA:Three B'!A1)"),
    'spaced_sheets_bracketed': (('Three A', 'Three B'), "=SUM(('Three A:Three B'!A1))"),
}


def macro(formula: str, member: str = 'Formula', name: bool = False,
          sheets: tuple[str, str] = ('ThreeA', 'ThreeB')) -> str:
    quoted = formula.replace('"', '""')
    add = 'ThisWorkbook.Names.Add Name:="ThreeA", RefersTo:="=ThreeB!$B$1"\n' if name else ''
    delete = 'ThisWorkbook.Names("ThreeA").Delete\n' if name else ''
    first, last = sheets
    return f'''Function Probe() As String
Dim ws As Worksheet, a As Worksheet, b As Worksheet, s As String
Set ws = Worksheets.Add
Set a = Worksheets.Add(After:=ws)
a.Name = "{first}"
a.Range("A1").Value = 1: a.Range("B1").Value = 2: a.Range("A2").Value = 3: a.Range("B2").Value = "x"
a.Range("C3").Value = 5
Set b = Worksheets.Add(After:=a)
b.Name = "{last}"
b.Range("A1").Value = 10: b.Range("B1").Value = 20: b.Range("A2").Value = 30: b.Range("B2").Value = True
b.Range("C3").Value = 50
ThisWorkbook.Names.Add Name:="BothSheets", RefersTo:="={first}:{last}!$A$1"
{add}ws.Range("C3").{member} = "{quoted}"
s = CStr(ws.Range("C3").Value) & "|" & ws.Range("C3").Formula
ThisWorkbook.Names("BothSheets").Delete
{delete}Application.DisplayAlerts = False
ws.Delete
a.Delete
b.Delete
Application.DisplayAlerts = True
Probe = s
End Function
'''


def probes(group: dict[str, str]) -> dict[str, str]:
    found = {name: macro(formula) for name, formula in group.items()}
    found.update({name: macro(formula, 'Formula2') for name, (owner, formula) in FORMULAS2.items() if owner is group})
    if group is REFERENCES:
        found.update({name: macro(formula, name=True) for name, formula in NAMED.items()})
        found.update({name: macro(formula, sheets=sheets) for name, (sheets, formula) in NAMED_SHEETS.items()})
    return found


if __name__ == '__main__':
    wanted = sys.argv[1:] or ['references', 'functions']
    if 'references' in wanted:
        record(probes(REFERENCES), FIXTURES / 'three_d_references.json')
    if 'functions' in wanted:
        record(probes(FUNCTIONS), FIXTURES / 'three_d_functions.json')
