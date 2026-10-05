#!/usr/bin/env python3
"""Build editable worksheet examples for every public native interface."""
import argparse,json
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
def generate():
 specs=json.loads((ROOT/'api/contracts.json').read_text())['functions']; sheets=[]
 for family in dict.fromkeys(s['family'] for s in specs):
  sheet=dict(name=family,title=family+' API examples',description='Edit the input cells in column B. Detailed results spill into D:F. Indices match docs/API-REFERENCE.md; options are an optional two-column key/value range. Zero event dates mean unavailable contacts.',headers=['Input / function','Editable value','Input units','Output field','Value','Output units'],widths=[39,29,62,32,42,80],rows=[],cells=[],formulas=[],examples=[])
  r=6
  for s in (s for s in specs if s['family']==family):
   sheet['cells'].append(dict(cell=f'A{r}',value=s['worksheetName']));start=r;r+=2;args=[]
   for p in s['parameters']:
    if p['direction']=='out':continue
    value=p['example'];first=r
    for i,v in enumerate(value if isinstance(value,list) else [value]):
     sheet['cells']+= [dict(cell=f'A{r}',value=p['name']+(f'[{i}]' if isinstance(value,list) else '')),dict(cell=f'B{r}',value=v),dict(cell=f'C{r}',value=p['units'])]
     if v=='@DATA@':sheet['formulas'].append(dict(cell=f'B{r}',formula='=SW_DATA_PATH()'));sheet['cells'][-2]['value']=''
     r+=1
    args.append(f'B{first}:B{r-1}' if isinstance(value,list) else f'B{first}')
   if s['command']:
    sheet['cells'].append(dict(cell=f'D{start}',value='VBA command: '+s['worksheetName']+'('+', '.join(p['name'] for p in s['parameters'] if p['direction']!='out')+')'))
    sheet['cells'].append(dict(cell=f'D{start+1}',value='Run from VBA; setters do not configure worksheet formulas. Use explicit options there.'))
   else:
    formula='='+s['worksheetName']+'('+','.join(args+['','TRUE'])+')'
    sheet['formulas'].append(dict(cell=f'D{start}',formula=formula))
    sheet['examples'].append(dict(name=s['name'],cell=f'D{start}',formula=formula))
   # Reserve enough rows for all native output slots, metadata, and inputs.
   count=sum((12 if p['name'] in ('cusps','cusp_speed') else p['visible']) for p in s['parameters'] if p['direction'] in ('out','inout') and p['name']!='serr')
   r=max(r+4,start+count+100)
  sheets.append(sheet)
 return json.dumps({'sheets':sheets},indent=2)+'\n'
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--check',action='store_true');a=p.parse_args();out=ROOT/'workbook/api-examples.json';value=generate()
 if a.check:assert out.read_text()==value,'Examples are stale'
 else:out.write_text(value)
 print('95 worksheet examples and 11 VBA command recipes agree with contracts.')
