"""A disconnected ADODB.Recordset, as a macro keeps a table in memory, and Range.CopyFromRecordset. No connection and no provider:
each recordset is made by appending fields, so nothing outside the process is reached.

    python scripts/measure_adodb_recordset.py

writes tests/fixtures/adodb_recordset.json, which tests/test_adodb_recordset.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/adodb_recordset.json'
CASES = [
    # Before it is open.
    'Set r = New0()\ns = r.State & "|" & r.Fields.Count & "|" & r.CursorType & "|" & r.LockType & "|" & r.CursorLocation',
    'Set r = New0()\ns = r.EOF',
    'Set r = New0()\ns = r.RecordCount',
    'Set r = New0()\nr.AddNew',
    'Set r = New0()\nr.MoveFirst',
    'Set r = New0()\nr.Close',
    'Set r = New0()\nr.Open',
    'Set r = New0()\ns = r.Fields(0).Name',
    'Set r = New0()\nr.Fields.Append "N", adVarChar, 10\ns = r.Fields.Count & r.Fields(0).Name & r.Fields(0).Type & r.Fields(0).DefinedSize & "|" & r.State',
    'Set r = New0()\nr.Fields.Append "N", adVarChar\ns = r.Fields(0).DefinedSize',
    'Set r = New0()\nr.Fields.Append "N", adVarChar, 10\nr.Fields.Append "n", adInteger',
    # Empty and open.
    'Set r = Made()\ns = r.State & "|" & r.EOF & r.BOF & r.RecordCount & "|" & r.CursorType & r.LockType & r.CursorLocation',
    'Set r = Made()\nr.MoveFirst',
    'Set r = Made()\nr.MoveNext',
    'Set r = Made()\ns = r("N")',
    'Set r = Made()\nr.Fields.Append "X", adInteger',
    'Set r = Made()\nr.Open',
    'Set r = Made()\nr.Close\ns = r.State & r.Fields.Count',
    'Set r = Made()\nr.Close\nr.Open\ns = r.State & r.RecordCount',
    # Adding.
    'Set r = Made()\nr.AddNew\ns = r.RecordCount & "|" & r.EOF & "|" & IsNull(r("N")) & "|" & r.EditMode\nr("N") = "a"\ns = s & r("N") & r.EditMode\nr.Update\ns = s & r.EditMode & r.RecordCount',
    'Set r = Made()\nr.AddNew\nr("N") = "a"\nr.AddNew\nr("N") = "b"\nr.Update\ns = r.RecordCount & r("N") & r.AbsolutePosition',
    'Set r = Made()\nr.AddNew Array("N", "Q"), Array("a", 1)\ns = r.RecordCount & r("N") & r("Q") & r.EditMode',
    'Set r = Made()\nr.AddNew "N", "a"\ns = r.RecordCount & r("N") & r.EditMode',
    'Set r = Made()\nr.AddNew\nr.CancelUpdate\ns = r.RecordCount & r.EOF & r.BOF',
    'Set r = Rows3()\ns = r.RecordCount & "|" & r.AbsolutePosition & "|" & r("N") & r.EOF & r.BOF',
    'Set r = Rows3()\nr.Update\ns = r.EditMode',
    'Set r = Rows3()\nr.CancelUpdate\ns = r.EditMode',
    # Moving.
    'Set r = Rows3()\nr.MoveFirst\ns = r("N") & r.AbsolutePosition\nr.MoveNext\ns = s & r("N") & r.AbsolutePosition\nr.MoveLast\ns = s & r("N") & r.AbsolutePosition\nr.MovePrevious\ns = s & r("N")',
    'Set r = Rows3()\nr.MoveFirst\nr.MovePrevious\ns = r.BOF & r.EOF & r.AbsolutePosition\nr.MovePrevious',
    'Set r = Rows3()\nr.MoveLast\nr.MoveNext\ns = r.EOF & r.BOF & r.AbsolutePosition\nr.MoveNext',
    'Set r = Rows3()\nr.MoveLast\nr.MoveNext\ns = r("N")',
    'Set r = Rows3()\nr.MoveFirst\nr.Move 2\ns = r("N")\nr.Move -1\ns = s & r("N")\nr.Move 5\ns = s & r.EOF',
    'Set r = Rows3()\nr.MoveFirst\nr.Move 1, adBookmarkLast\ns = r.EOF',
    'Set r = Rows3()\nr.AbsolutePosition = 2\ns = r("N")\nr.AbsolutePosition = 9',
    'Set r = Rows3()\nr.MoveFirst\nDo Until r.EOF\ns = s & r("N") & r("Q") & ","\nr.MoveNext\nLoop',
    'Set r = Rows3()\nr.MoveFirst\nv = r.Bookmark\nr.MoveLast\nr.Bookmark = v\ns = r("N") & TypeName(v) & v',
    # Reading fields.
    'Set r = Rows3()\nr.MoveFirst\ns = r!N & r.Fields("Q").Value & r.Fields(0) & r(1) & r.Fields.Item("N").Name & r.Collect("N")',
    'Set r = Rows3()\nr.MoveFirst\ns = r("n") & TypeName(r("N").Value) & TypeName(r("Q").Value) & TypeName(r("N"))',
    'Set r = Rows3()\nr.MoveFirst\ns = r("Nope")',
    'Set r = Rows3()\nr.MoveFirst\ns = r.Fields(9).Name',
    'Set r = Rows3()\nFor Each o In r.Fields\ns = s & o.Name & o.Type & ","\nNext',
    'Set r = Rows3()\ns = r.Fields(1).Name & r.Fields(1).Type & r.Fields(1).DefinedSize & r.Fields(1).ActualSize & r.Fields(0).ActualSize',
    # Types and values.
    'Set r = Typed()\nr.AddNew\nr("I") = "5"\nr("D") = "1.5"\nr("B") = 1\nr("T") = "1/2/2020"\nr("S") = 12\nr("C") = 2.5\nr.Update\ns = TypeName(r("I").Value) & r("I") & TypeName(r("D").Value) & r("D") & TypeName(r("B").Value) & r("B") & TypeName(r("T").Value) & r("T") & TypeName(r("S").Value) & r("S") & TypeName(r("C").Value) & r("C")',
    'Set r = Typed()\nr.AddNew\nr("I") = "x"',
    'Set r = Typed()\nr.AddNew\nr("I") = 99999999999#',
    'Set r = Typed()\nr.AddNew\nr("S") = "abcdefghijk"\ns = Len(r("S"))',
    'Set r = Typed()\nr.AddNew\nr("I") = Null\nr("S") = Null\ns = IsNull(r("I")) & IsNull(r("S"))',
    'Set r = Typed()\nr.AddNew\nr("S") = ""\ns = Len(r("S"))',
    # Editing and deleting.
    'Set r = Rows3()\nr.MoveFirst\nr("N") = "z"\ns = r.EditMode & r("N")\nr.MoveNext\ns = s & r.EditMode\nr.MoveFirst\ns = s & r("N")',
    'Set r = Rows3()\nr.MoveFirst\nr("N") = "z"\nr.CancelUpdate\ns = r("N") & r.EditMode',
    'Set r = Rows3()\nr.MoveFirst\nr.Update "N", "w"\ns = r("N") & r.EditMode',
    'Set r = Rows3()\nr.MoveFirst\nr.Delete\ns = r.RecordCount & r.EOF\ns = s & r("N")',
    'Set r = Rows3()\nr.MoveFirst\nr.Delete\nr.MoveNext\ns = r("N") & r.RecordCount & r.AbsolutePosition',
    'Set r = Rows3()\nr.MoveLast\nr.Delete\nr.MoveNext\ns = r.EOF & r.RecordCount',
    # Sort, Filter, Find.
    'Set r = Rows3()\nr.Sort = "Q DESC"\ns = r.Sort & r("N") & r.AbsolutePosition\nr.MoveNext\ns = s & r("N")',
    'Set r = Rows3()\nr.Sort = "N"\nr.MoveFirst\ns = r("N")\nr.Sort = ""\nr.MoveFirst\ns = s & r("N")',
    'Set r = Rows3()\nr.Sort = "Nope"',
    'Set r = Rows3()\nr.Filter = "Q > 1"\ns = r.RecordCount & r("N") & r.AbsolutePosition\nr.Filter = adFilterNone\ns = s & r.RecordCount & TypeName(r.Filter) & r.Filter',
    'Set r = Rows3()\nr.Filter = "N = \'b\'"\ns = r.RecordCount & r("N") & r.Filter',
    'Set r = Rows3()\nr.Filter = "N LIKE \'b*\'"\ns = r.RecordCount',
    'Set r = Rows3()\nr.Filter = "Q > 9"\ns = r.RecordCount & r.EOF & r.BOF',
    'Set r = Rows3()\nr.Filter = "Q > 1 AND N = \'c\'"\ns = r.RecordCount & r("N")',
    'Set r = Rows3()\nr.Filter = "Nope = 1"',
    'Set r = Rows3()\nr.MoveFirst\nr.Find "N = \'b\'"\ns = r("N") & r.AbsolutePosition\nr.Find "N = \'zz\'"\ns = s & r.EOF',
    'Set r = Rows3()\nr.MoveLast\nr.Find "N = \'a\'", , adSearchBackward\ns = r("N")',
    'Set r = Rows3()\nr.MoveFirst\nr.Find "Q > 1", 1\ns = r("N")',
    # Whole rows.
    'Set r = Rows3()\nr.MoveFirst\nv = r.GetRows\ns = LBound(v, 1) & UBound(v, 1) & LBound(v, 2) & UBound(v, 2) & v(0, 2) & v(1, 2) & r.EOF',
    'Set r = Rows3()\nr.MoveFirst\nv = r.GetRows(2)\ns = UBound(v, 2) & r("N")',
    'Set r = Rows3()\nr.MoveFirst\nv = r.GetRows(, , "Q")\ns = UBound(v, 1) & v(0, 1)',
    'Set r = Made()\nv = r.GetRows',
    'Set r = Rows3()\nr.MoveFirst\ns = Replace(Replace(r.GetString, vbTab, "T"), vbCr, "R") & r.EOF',
    'Set r = Rows3()\nr.MoveFirst\ns = r.GetString(adClipString, 2, ",", ";", "-")',
    'Set r = Clone3()\nr.MoveFirst\ns = r("N")',
    'Set r = Rows3()\ns = r.Supports(adAddNew) & r.Supports(adMovePrevious) & r.Supports(adBookmark) & r.Supports(adDelete) & r.Supports(adFind)',
    # Copying to the sheet.
    'Set r = Rows3()\nr.MoveFirst\nv = ws.Range("A1").CopyFromRecordset(r)\ns = TypeName(v) & v & "|" & Cells3(ws) & "|" & r.EOF',
    'Set r = Rows3()\nr.MoveFirst\nr.MoveNext\nv = ws.Range("A1").CopyFromRecordset(r)\ns = v & "|" & Cells3(ws)',
    'Set r = Rows3()\nr.MoveFirst\nv = ws.Range("A1").CopyFromRecordset(r, 2)\ns = v & "|" & Cells3(ws) & "|" & r("N")',
    'Set r = Rows3()\nr.MoveFirst\nv = ws.Range("A1").CopyFromRecordset(r, , 1)\ns = v & "|" & Cells3(ws)',
    'Set r = Rows3()\nr.MoveFirst\nv = ws.Range("A1:A1").CopyFromRecordset(r, 0)\ns = v & "|" & Cells3(ws)',
    'Set r = Rows3()\nr.MoveFirst\nv = ws.Range("C2").CopyFromRecordset(r)\ns = v & ws.Range("C2").Value & ws.Range("D4").Value',
    'Set r = Rows3()\nr.MoveFirst\nv = ws.Range("B2:C2").CopyFromRecordset(r)\ns = v & ws.Range("B2").Value & ws.Range("C4").Value',
    'Set r = Rows3()\nr.MoveFirst\nv = ws.Range("A1,C1").CopyFromRecordset(r)\ns = v & ws.Range("A1").Value & ws.Range("C1").Value',
    'Set r = Rows3()\nr.MoveLast\nr.MoveNext\nv = ws.Range("A1").CopyFromRecordset(r)\ns = v & "|" & ws.Range("A1").Value',
    'Set r = Made()\nv = ws.Range("A1").CopyFromRecordset(r)\ns = TypeName(v) & v',
    'Set r = New0()\nv = ws.Range("A1").CopyFromRecordset(r)',
    'Set r = Rows3()\nr.Close\nv = ws.Range("A1").CopyFromRecordset(r)',
    'v = ws.Range("A1").CopyFromRecordset(5)',
    'Set r = Nothing\nv = ws.Range("A1").CopyFromRecordset(r)',
    'Set r = Typed()\nr.AddNew\nr("I") = 5\nr("D") = 1.5\nr("B") = True\nr("T") = #1/2/2020#\nr("S") = "x"\nr("C") = 2.5\nr.Update\nr.MoveFirst\nv = ws.Range("A1").CopyFromRecordset(r)\nFor i = 1 To 6\ns = s & TypeName(ws.Cells(1, i).Value) & "=" & ws.Cells(1, i).Value & "~" & ws.Cells(1, i).NumberFormat & ","\nNext',
    'Set r = Typed()\nr.AddNew\nr("I") = Null\nr("S") = "=1+1"\nr.Update\nr.MoveFirst\nv = ws.Range("A1").CopyFromRecordset(r)\ns = IsEmpty(ws.Range("A1").Value) & ws.Range("E1").Value & ws.Range("E1").HasFormula',
    'Set r = Rows3()\nr.MoveFirst\nws.Range("A1:C1").Value = 9\nv = ws.Range("A1").CopyFromRecordset(r)\ns = ws.Range("C1").Value & ws.Range("A1").Value',
    'Set r = Rows3()\nr.Filter = "Q > 1"\nr.MoveFirst\nv = ws.Range("A1").CopyFromRecordset(r)\ns = v & Cells3(ws)',
    'Set r = Rows3()\nr.MoveFirst\nws.Protect\nOn Error Resume Next\nv = ws.Range("A1").CopyFromRecordset(r)\ns = Err.Number & "|" & ws.Range("A1").Value\nws.Unprotect',
    # Odds.
    'Set r = Rows3()\nr.Requery',
    'Set r = Rows3()\ns = TypeName(r) & TypeName(r.Fields) & TypeName(r.Fields(0)) & TypeName(r.ActiveConnection) & r.Source & "|" & r.PageSize & r.PageCount & r.MaxRecords & r.CacheSize',
    'Set r = Rows3()\nr.Close\nSet r = Nothing\ns = r Is Nothing',
    'Set r = Rows3()\nr.Close\ns = r.RecordCount',
    'Set r = Rows3()\nr.Close\ns = r.Fields.Count & r.Fields(0).Name',
    'Set r = Rows3()\nr.Close\ns = r("N")',
    # Where Sort and Filter leave the current row.
    'Set r = Rows3()\nr.MoveFirst\nr.MoveNext\nr.Sort = "Q DESC"\ns = r("N") & r.AbsolutePosition',
    'Set r = Rows3()\nr.MoveFirst\nr.Sort = "Q DESC"\ns = r("N") & r.AbsolutePosition',
    'Set r = Rows3()\nr.MoveFirst\nr.MoveNext\nr.Filter = "Q > 0"\ns = r("N") & r.AbsolutePosition',
    'Set r = Rows3()\nr.MoveFirst\nr.MoveNext\nr.Filter = "Q > 1"\nr.Filter = 0\ns = r("N") & r.AbsolutePosition',
    'Set r = Rows3()\nr.Sort = "N DESC, Q"\nr.MoveFirst\ns = r("N")\nr.Sort = "N ASC"\nr.MoveFirst\ns = s & r("N")',
    'Set r = Made()\nr.AddNew Array("N", "Q"), Array("b", 1)\nr.AddNew Array("N", "Q"), Array("a", 2)\nr.AddNew Array("Q"), Array(3)\nr.Update\nr.Sort = "N"\nr.MoveFirst\ns = r("Q")\nr.MoveNext\ns = s & r("Q")\nr.MoveNext\ns = s & r("Q")',
    # Filter's forms.
    *(f'Set r = Rows3()\nr.Filter = "{text}"\nr.MoveFirst\ns = r.RecordCount & r("N")'
      for text in ("Q >= 2 OR N = 'a'", "N <> 'a'", "N = 'B'", "(Q = 1)", "Q IN (1, 2)", "N LIKE '*b'", "N LIKE '%b%'", "N LIKE 'b%'", "Q = '1'", "Q = 1 AND (N = 'a' OR N = 'b')", "Q > 1 OR Q = 1 AND N = 'z'", "N = b", "N = \"\"b\"\"")),
    'Set r = Rows3()\nr.Filter = ""\ns = r.RecordCount & TypeName(r.Filter) & r.Filter',
    'Set r = Rows3()\nr.Filter = "Q > 1"\nr.MoveLast\nr.Delete\nr.MoveFirst\ns = r.RecordCount\nr.Filter = 0\ns = s & r.RecordCount',
    'Set r = Rows3()\nr.Filter = "Q > 1"\nr.AddNew\nr("Q") = 0\nr("N") = "n"\nr.Update\ns = r.RecordCount & r("N")\nr.MoveFirst\ns = s & r.RecordCount',
    # Find's forms.
    'Set r = Rows3()\nr.MoveFirst\nr.Find "N LIKE \'b*\'"\ns = r("N")',
    'Set r = Rows3()\nr.MoveFirst\nr.Find "Q > 1 AND N = \'c\'"\ns = r("N")',
    'Set r = Made()\nr.Find "Q > 1"\ns = r.EOF',
    'Set r = Rows3()\nr.MoveFirst\nr.MoveNext\nr.Find "N = \'a\'"\ns = r.EOF',
    'Set r = Rows3()\nr.MoveLast\nr.Find "N = \'zz\'", , adSearchBackward\ns = r.BOF & r.EOF',
    # Small things around the current row.
    'Set r = Made()\ns = r.EditMode',
    'Set r = Made()\nr.Delete',
    'Set r = Made()\nr.Update',
    'Set r = Made()\nr.CancelUpdate',
    'Set r = Rows3()\nr.MoveFirst\nr.Move 0\ns = r("N")',
    'Set r = Rows3()\nr.MoveFirst\nr.Move -5\ns = r.BOF & r.EOF',
    'Set r = Rows3()\nr.MoveFirst\nr.Delete\nr.MoveFirst\ns = r("N") & r.EditMode & r.RecordCount',
    'Set r = Rows3()\nr.MoveFirst\nr.Delete\ns = r.EditMode',
    'Set r = Rows3()\nr.MoveFirst\nr.Delete\nr.Delete',
    'Set r = Rows3()\nr.MoveFirst\nr.Delete\nr("N") = "x"',
    'Set r = Rows3()\nr.MoveFirst\nr.Delete\nr.MovePrevious\ns = r.BOF',
    'Set r = Rows3()\nr.MoveFirst\nr.CancelUpdate\nr.MoveFirst\ns = r("N") & r.RecordCount',
    'Set r = Rows3()\nr.MoveLast\nr.CancelUpdate\ns = r.RecordCount & r.EOF & r.BOF',
    'Set r = Rows3()\nr.UpdateBatch\nr.MoveFirst\ns = r.EditMode\nr("N") = "z"\ns = s & r.EditMode\nr.Update\ns = s & r.EditMode & r("N")\nr("N") = "y"\nr.CancelUpdate\ns = s & r("N") & r.EditMode & r.RecordCount',
    'Set r = Rows3()\nr.UpdateBatch\nr.MoveFirst\nr.Delete\ns = r.RecordCount\nr.CancelBatch\ns = s & r.RecordCount',
    'Set r = Rows3()\nr.MoveFirst\nr.CancelBatch\ns = r.RecordCount & r.EOF',
    'Set r = Rows3()\nr.MoveFirst\nr.Fields("N").Value = "x"\nr.Fields(1) = 7\ns = r("N") & r("Q") & r("N").OriginalValue & "|" & r("N").UnderlyingValue',
    'Set r = Rows3()\nr.UpdateBatch\nr.MoveFirst\nr("N") = "x"\ns = r("N") & r("N").OriginalValue & r("N").UnderlyingValue',
    'Set r = Rows3()\ns = r.Fields(0).Attributes & "|" & r.Fields(1).Attributes & "|" & r.Fields(1).Precision & r.Fields(1).NumericScale & "|" & r.Status',
    'Set r = Typed()\nr.AddNew\ns = r.Fields(0).Attributes',
    'Set r = Rows3()\nr.MoveFirst\nr("N") = Null\ns = IsNull(r("N"))',
    'Set r = Rows3()\nr.Fields.Delete "N"',
    'Set r = New0()\nr.Fields.Append "N", adVarChar, 10\nr.Fields.Delete "N"\ns = r.Fields.Count',
    'Set r = New0()\nr.Fields.Append "N", adVarChar, 10\nr.Fields.Append "Q", adInteger\nr.Open\ns = r.Fields.Refresh',
    'Set r = Rows3()\nr.MoveFirst\ns = Replace(r.GetString(, , , , "-"), vbCr, "R")',
    'Set r = Made()\ns = r.GetString',
    'Set r = Made()\nr.AddNew\nr("Q") = 1\nr.Update\nr.MoveFirst\ns = "[" & Replace(Replace(r.GetString, vbTab, "T"), vbCr, "R") & "]"',
    'Set r = Made()\nr.AddNew\nr("Q") = 1\nr.Update\nr.MoveFirst\ns = "[" & Replace(Replace(r.GetString(, , , , "-"), vbTab, "T"), vbCr, "R") & "]"',
    'Set r = Rows3()\nr.MoveFirst\nv = r.GetRows(, 2)\ns = UBound(v, 2)',
    'Set r = Rows3()\nr.MoveFirst\nv = r.GetRows(, , Array("Q", "N"))\ns = UBound(v, 1) & v(0, 0) & v(1, 0)',
    'Set r = Rows3()\nr.MoveFirst\nv = r.GetRows(, , 1)\ns = UBound(v, 1) & v(0, 0)',
    'Set r = Rows3()\nr.MoveFirst\nv = r.GetRows(, , "Nope")',
    'Set r = Rows3()\nr.MoveFirst\nv = r.GetRows(9)\ns = UBound(v, 2)',
    'Set r = Rows3()\nr.MoveFirst\nv = r.GetRows(0)\ns = TypeName(v)',
    'Set r = Rows3()\nSet o = r.Clone\nr.MoveFirst\no.MoveLast\ns = r("N") & o("N")\nr("N") = "z"\no.MoveFirst\ns = s & o("N") & o.RecordCount\no.AddNew\no("Q") = 9\no.Update\ns = s & r.RecordCount',
    'Set r = Rows3()\nSet o = r.Clone\nr.Close\ns = o.State & o.RecordCount',
    'Set r = Rows3()\ns = r.Supports(adUpdate) & r.Supports(adUpdateBatch) & r.Supports(adResync) & r.Supports(adHoldRecords) & r.Supports(adApproxPosition) & r.Supports(adIndex) & r.Supports(adSeek) & r.Supports(adNotify)',
    # The other field types.
    *(f'Set r = New0()\nr.Fields.Append "F", {code}, 10, adFldIsNullable\nr.Open\nr.AddNew\ns = r("F").Type & "~" & r("F").DefinedSize & "~"\nr("F") = {value}\ns = s & TypeName(r("F").Value) & "~" & r("F").Value'
      for code, value in ((202, '"5"'), (130, '"5"'), (201, '"5"'), (203, '"5"'), (129, '"5"'), (8, '"5"'), (2, '5.4'), (4, '5.5'), (16, '5'), (17, '5'), (18, '5'), (19, '5'), (20, '5'), (21, '5'), (14, '5.5'), (131, '5.5'), (133, '#1/2/2020#'), (134, '#1/2/2020 3:04:05 AM#'), (135, '#1/2/2020#'), (72, '"{00000000-0000-0000-0000-000000000001}"'), (128, '"5"'), (204, '"5"'), (205, '"5"'), (12, '5'), (64, '#1/2/2020#'))),
    'Set r = New0()\nr.Fields.Append "F", adInteger, 99\ns = r.Fields(0).DefinedSize',
    'Set r = New0()\nr.Fields.Append "F", 9999, 10',
    # What CopyFromRecordset writes.
    'Set r = Typed()\nr.AddNew\nr("S") = "12"\nr.Update\nr.MoveFirst\nv = ws.Range("A1").CopyFromRecordset(r)\ns = TypeName(ws.Range("E1").Value) & ws.Range("E1").Value & ws.Range("E1").NumberFormat & ws.Range("E1").PrefixCharacter',
    'Set r = Typed()\nr.AddNew\nr("S") = "1/2"\nr.Update\nr.MoveFirst\nv = ws.Range("A1").CopyFromRecordset(r)\ns = TypeName(ws.Range("E1").Value) & ws.Range("E1").Value & ws.Range("E1").NumberFormat',
    'Set r = Typed()\nr.AddNew\nr("S") = "\'x"\nr.Update\nr.MoveFirst\nv = ws.Range("A1").CopyFromRecordset(r)\ns = ws.Range("E1").Value & ws.Range("E1").PrefixCharacter & ws.Range("E1").HasFormula',
    'Set r = Typed()\nr.AddNew\nr("C") = 2.5\nr.Update\nr.MoveFirst\nws.Range("F1").NumberFormat = "0.000"\nws.Range("D1").NumberFormat = "0.000"\nv = ws.Range("A1").CopyFromRecordset(r)\ns = ws.Range("F1").NumberFormat & ws.Range("D1").NumberFormat',
    'Set r = Typed()\nr.AddNew\nr("T") = #1/2/2020#\nr.Update\nr.MoveFirst\nws.Range("D1").NumberFormat = "0.000"\nv = ws.Range("A1").CopyFromRecordset(r)\ns = ws.Range("D1").NumberFormat & ws.Range("D1").Value2',
    'Set r = Typed()\nr.AddNew\nr("T") = #1/2/2020 3:04:05 AM#\nr.Update\nr.MoveFirst\nv = ws.Range("A1").CopyFromRecordset(r)\ns = ws.Range("D1").NumberFormat & ws.Range("D1").Text',
    'Set r = Rows3()\nr.MoveFirst\nws.Range("A1").Value = "x"\nws.Range("A1:B2").Merge\nv = ws.Range("A1").CopyFromRecordset(r)\ns = v & ws.Range("A1").Value & ws.Range("A2").Value & ws.Range("B3").Value',
    'Set r = Rows3()\nr.MoveFirst\nv = ws.Range("A1048575").CopyFromRecordset(r)\ns = v & ws.Range("A1048576").Value',
    'Set r = Rows3()\nr.MoveFirst\nv = ws.Range("XFD1").CopyFromRecordset(r)\ns = v & ws.Range("XFD1").Value',
    'Set r = Rows3()\nr.MoveFirst\nv = ws.Rows(2).CopyFromRecordset(r)\ns = v & ws.Range("A2").Value',
    'Set r = Rows3()\nr.MoveFirst\nws.Range("A1").CopyFromRecordset r\ns = ws.Range("A1").Value',
    'Set r = Rows3()\nr.MoveFirst\nv = ws.Range("A1").CopyFromRecordset(r, -1)',
    'Set r = Rows3()\nr.MoveFirst\nv = ws.Range("A1").CopyFromRecordset(r, , 0)\ns = v & ws.Range("A1").Value',
    'Set r = Rows3()\nr.MoveFirst\nv = ws.Range("A1").CopyFromRecordset(r, , 9)\ns = v & ws.Range("B1").Value',
    'Set r = Rows3()\nr.Filter = "Q > 9"\nv = ws.Range("A1").CopyFromRecordset(r)\ns = v',
    'Set r = Rows3()\nr.MoveFirst\nr.Delete\nv = ws.Range("A1").CopyFromRecordset(r)\ns = v & ws.Range("A1").Value',
    'Set r = Rows3()\nr.MoveFirst\nr.MovePrevious\nv = ws.Range("A1").CopyFromRecordset(r)\ns = v & ws.Range("A1").Value',
    'Set r = Rows3()\nr.MoveFirst\nws.Range("A1").Validation.Add xlValidateWholeNumber, , xlBetween, "1", "5"\nv = ws.Range("A1").CopyFromRecordset(r)\ns = v & ws.Range("A1").Value',
    'Set r = Rows3()\nr.MoveFirst\nws.ListObjects.Add xlSrcRange, ws.Range("A1:B2"), , xlYes\nv = ws.Range("A1").CopyFromRecordset(r)\ns = v & ws.Range("A1").Value & ws.ListObjects(1).Range.Address',
    'Set r = Rows3()\nr.MoveFirst\nv = ws.Range("A3").CopyFromRecordset(r)\ns = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row & ws.UsedRange.Address',
]
PROBE = """Private Function New0() As Object
    Set New0 = CreateObject("ADODB.Recordset")
End Function

Private Function Made() As Object
    Set Made = New0()
    Made.Fields.Append "N", adVarChar, 10
    Made.Fields.Append "Q", adInteger
    Made.Open
End Function

Private Function Rows3() As Object
    Dim i As Long, t As Object
    Set t = Made()
    For i = 1 To 3
        t.AddNew
        t("N") = Chr(96 + i)
        t("Q") = i
        t.Update
    Next
    Set Rows3 = t
End Function

Private Function Clone3() As Object
    Set Clone3 = Rows3().Clone
End Function

Private Function Typed() As Object
    Set Typed = New0()
    Typed.Fields.Append "I", adInteger, , adFldIsNullable
    Typed.Fields.Append "D", adDouble, , adFldIsNullable
    Typed.Fields.Append "B", adBoolean, , adFldIsNullable
    Typed.Fields.Append "T", adDate, , adFldIsNullable
    Typed.Fields.Append "S", adVarChar, 5, adFldIsNullable
    Typed.Fields.Append "C", adCurrency, , adFldIsNullable
    Typed.Open
End Function

Private Function Cells3(ByVal ws As Object) As String
    Dim i As Long
    For i = 1 To 4
        Cells3 = Cells3 & ws.Cells(i, 1).Value & ws.Cells(i, 2).Value & ";"
    Next
End Function

Public Function Probe(ByVal which As Long) As String
    Dim ws As Object, r As Object, o As Object, s As String, v As Variant, i As Long
    Application.DisplayAlerts = False
    Set ws = Workbooks.Add(xlWBATWorksheet).Worksheets(1)
    On Error GoTo Failed
    Select Case which
@@CASES@@
    End Select
    s = "ok " & s
Answered:
    On Error Resume Next
    ws.Unprotect
    ws.Parent.Close False
    On Error GoTo 0
    Application.DisplayAlerts = True
    Probe = s
    Exit Function
Failed:
    s = "Err " & Err.Number & " " & s
    Resume Answered
End Function
"""
#: The ADO constants, since the probe has no reference to the library; each stays a plain number for the replay too.
CONSTANTS = {'adVarChar': 200, 'adInteger': 3, 'adDouble': 5, 'adBoolean': 11, 'adDate': 7, 'adCurrency': 6, 'adFldIsNullable': 32, 'adBookmarkLast': 2,
             'adFilterNone': 0, 'adSearchBackward': -1, 'adAddNew': 16778240, 'adMovePrevious': 512, 'adBookmark': 8192, 'adDelete': 16779264, 'adFind': 524288,
             'adClipString': 2, 'adUpdate': 16809984, 'adUpdateBatch': 65536, 'adResync': 131072, 'adHoldRecords': 256, 'adApproxPosition': 16384, 'adIndex': 1048576,
             'adSeek': 2097152, 'adNotify': 262144}


def source() -> str:
    head = '\n'.join(f'Private Const {name} As Long = {value}' for name, value in CONSTANTS.items())
    cases = '\n'.join(f'    Case {index}\n        {case}' for index, case in enumerate(CASES))
    return head + '\n' + PROBE.replace('@@CASES@@', cases)


def main() -> None:
    text = source()
    cases: list[list[str]] = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for index, case in enumerate(CASES):
            result = excel.run_vba(text, 'Probe', (index,), timeout=60.0)
            assert result.ok and not result.dialogs, (case, result)
            cases.append([case, str(result.value)])
            print(index, case.replace('\n', '; ')[-64:], '=>', str(result.value)[:90], flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': text, 'cases': cases}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
