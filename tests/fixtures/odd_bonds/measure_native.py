from __future__ import annotations
from pathlib import Path
import argparse,datetime as dt,hashlib,json
from zipfile import ZipFile
ROOT=Path(__file__).resolve().parent
OUT=ROOT/'odd-bonds-native-v2.json'
CASES=[]
def add(name,formula):CASES.append({'name':name,'formula':formula})
for basis in range(5):
 for frequency in (1,2,4):
  for length,settle,issue,first in [('short','DATE(2024,6,1)','DATE(2024,5,15)','DATE(2024,6,30)'),('long','DATE(2023,2,17)','DATE(2022,1,15)','DATE(2024,6,30)')]:
   for function,value in [('ODDFPRICE','0.06'),('ODDFYIELD','98')]:add(f'{function}-{length}-f{frequency}-b{basis}',f'={function}({settle},DATE(2028,6,30),{issue},{first},0.05,{value},100,{frequency},{basis})')
  for length,maturity in [('short','DATE(2023,11,15)'),('long','DATE(2025,5,30)')]:
   for function,value in [('ODDLPRICE','0.06'),('ODDLYIELD','98')]:add(f'{function}-{length}-f{frequency}-b{basis}',f'={function}(DATE(2023,10,15),{maturity},DATE(2023,8,31),0.05,{value},100,{frequency},{basis})')
for index,formula in enumerate([
 '=ODDLPRICE(DATE(2021,9,10),DATE(2022,5,30),DATE(2021,8,31),0.1,0.1,100,4,1)',
 '=ODDLPRICE(DATE(2021,9,10),DATE(2022,5,27),DATE(2021,8,31),0.1,0.1,100,4,1)',
 '=ODDLPRICE(DATE(2018,4,2),DATE(2018,5,14),DATE(2018,2,28),0.1,0.1,100,2,0)',
 '=ODDLYIELD(DATE(2021,5,30),DATE(2021,5,31),DATE(2021,1,1),0.05,98,100,2,0)',
 '=ODDFPRICE(DATE(2021,11,3),DATE(2027,4,30),DATE(2021,10,29),DATE(2023,4,30),0.1,0.1,100,1,0)',
 '=ODDFYIELD(DATE(2022,4,12),DATE(2027,6,15),DATE(2021,10,29),DATE(2023,6,15),0.1,90,100,1,0)',
 '=ODDFPRICE(DATE(2021,2,1),DATE(2023,6,30),DATE(2020,12,31),DATE(2021,6,30),0.05,0.06,100,2,0)',
 '=ODDFPRICE(DATE(2021,2,3),DATE(2023,6,30),DATE(2021,1,1),DATE(2021,6,29),0.05,0.06,100,2,0)',
 '=ODDFPRICE(59,1000,1,181,0.05,0.06,100,2,1)',
 '=ODDLPRICE(60,181,1,0.05,0.06,100,2,1)',
]):add('calendar-edge-'+str(index),formula)
for function in ['ODDFPRICE','ODDFYIELD','ODDLPRICE','ODDLYIELD']:
 first=function.startswith('ODDF');price=function.endswith('PRICE')
 values=['DATE(2024,6,1)','DATE(2028,6,30)']+(['DATE(2024,5,15)','DATE(2024,6,30)']if first else ['DATE(2024,5,15)'])+['0.05','0.06'if price else'98','100','2','0']
 rate=4 if first else 3;yield_price=rate+1;redemption=rate+2;freq=rate+3;basis=rate+4
 for label,index,value in [('bad-settlement',0,'"bad"'),('negative-settlement',0,'-1'),('settlement-maturity',0,values[1]),('rate-negative',rate,'-0.01'),('redemption-zero',redemption,'0'),('frequency-invalid',freq,'3'),('frequency-truncated',freq,'2.9'),('basis-invalid',basis,'5'),('basis-truncated',basis,'1.9'),('value-negative',yield_price,'-1'),('value-zero',yield_price,'0'),('rate-zero',rate,'0'),('numeric-text',rate,'"0.05"'),('propagated-error',rate,'NA()')]:
  changed=values.copy();changed[index]=value;add(function+'-'+label,'='+function+'('+','.join(changed)+')')
 add(function+'-omitted-basis','='+function+'('+','.join(values[:-1])+')')
 add(function+'-fractional-dates','='+function+'('+','.join(v+'+0.75'if i < (4 if first else 3)else v for i,v in enumerate(values))+')')
HELPERS="""Private Type NativeDouble
    value As Double
End Type
Private Type NativeWords
    low As Long
    high As Long
End Type
Function DoubleBits(ByVal value As Double) As String
Dim d As NativeDouble,w As NativeWords
d.value=value
LSet w=d
DoubleBits=Right$("00000000" & Hex$(w.high),8) & Right$("00000000" & Hex$(w.low),8)
End Function
Sub Measure(ByRef output As String, ByVal i As Long, ByVal expression As String)
Dim cell As Range,v As Variant,item As String
Set cell=ActiveSheet.Cells(i+1,1)
On Error GoTo Failed
cell.Formula2=expression
v=cell.Value2
If IsError(v) Then
item=CStr(VarType(v)) & "|" & CStr(v) & "|" & cell.Text
Else
item=CStr(VarType(v)) & "|" & DoubleBits(CDbl(v))
End If
output=output & CStr(i) & "|" & item & vbLf
Exit Sub
Failed:
output=output & CStr(i) & "|VBA-ERROR:" & Err.Number & ":" & Err.Description & vbLf
End Sub
"""
def source():
 body='Function Probe() As String\nDim output As String\nApplication.EnableEvents=False\nApplication.Calculation=-4105\nActiveWorkbook.Date1904=False\n'
 body+='\n'.join('Measure output,'+str(i)+',"'+c['formula'].replace('"','""')+'"'for i,c in enumerate(CASES))
 return HELPERS+body+'\nProbe=output\nEnd Function\n'

def execute():
 from pyvbaharness import ExcelSession
 assert not OUT.exists()
 s=source();record={'started_utc':dt.datetime.now(dt.timezone.utc).isoformat(),'source':s,'source_sha256':hashlib.sha256(s.encode()).hexdigest(),'cases':CASES,'instrumentation':'One session freshbook Date1904False eventsdisabled Automatic; exact raw typed2Darray Double/errors; independent native input allformula cells; Save51 ifbatchsuccess'}
 def flush():OUT.write_text(json.dumps(record,indent=2,default=str)+'\n',encoding='utf8')
 flush()
 try:
  with ExcelSession()as excel:
   record.update(session_id=excel.session_id,excel_pid=excel.excel_pid);flush();excel.new_document()
   diagnostic='Function Diagnostic() As Variant\nDim data(0 To 1,0 To 1) As Variant\ndata(0,0)=1.25\ndata(0,1)="x"\ndata(1,0)=2.5\ndata(1,1)="y"\nDiagnostic=data\nEnd Function'
   diag=excel.run_vba(diagnostic,'Diagnostic',module_name='OddDiagnostic',line_numbers=False,timeout=None,idle_timeout=None)
   record['diagnostic']={'source':diagnostic,'repr':repr(diag),'message':diag.message,'ok':diag.ok,'value':diag.value};excel.save_trace(ROOT/'odd-bonds-native-v2-trace.jsonl');flush()
   if not diag.ok or diag.dialogs or diag.value!=[[1.25,'x'],[2.5,'y']]:return

   result=excel.run_vba(s,'Probe',module_name='OddBonds',line_numbers=False,timeout=None,idle_timeout=None)
   record.update(status={'ok':result.ok,'outcome':str(result.outcome),'dialogs':str(result.dialogs),'error':str(result.error),'message':result.message,'output':result.output,'repr':repr(result)},raw_result=result.value);excel.save_trace(ROOT/'odd-bonds-native-v2-trace.jsonl');flush()
   if not result.ok or result.dialogs:return
   path=ROOT/'odd-bonds-native-v2.xlsx';save='Function SaveIt() As String\nApplication.DisplayAlerts=False\nActiveWorkbook.SaveAs Filename:="'+str(path).replace('"','""')+'",FileFormat:=51\nSaveIt="saved"\nEnd Function'
   saved=excel.run_vba(save,'SaveIt',module_name='OddSave',line_numbers=False,timeout=None,idle_timeout=None)
   record.update(save_source=save,save_status={'ok':saved.ok,'dialogs':str(saved.dialogs)},save_result=saved.value);flush()
   if saved.ok and saved.value=='saved':
    record.update(file=str(path),file_sha256=hashlib.sha256(path.read_bytes()).hexdigest())
    with ZipFile(path)as z:record['parts']={n:hashlib.sha256(z.read(n)).hexdigest()for n in z.namelist()};assert not any('vbaproject'in n.lower()for n in z.namelist())
   record['completed']=True
 finally:
  record['closed_utc']=dt.datetime.now(dt.timezone.utc).isoformat();flush();print('SESSION CLOSED '+record['closed_utc'],flush=True)
if __name__=='__main__':
 parser=argparse.ArgumentParser();parser.add_argument('--execute',action='store_true');args=parser.parse_args()
 if args.execute:execute()
 else:
  from pyopenvba.interpreter._parse import parse_module
  parse_module(source(),name='OddBonds');print(json.dumps({'cases':len(CASES),'source_chars':len(source()),'native_executed':False}))
