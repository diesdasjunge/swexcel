#!/usr/bin/env python3
"""Execute every pinned native entry against independent C output, with guards.

Build the macOS reference shared library from the pinned C source first. Each
function runs in an isolated child so a native failure cannot erase other cases.
These fixtures establish ABI/output parity, not astronomical model correctness.
"""
import argparse,ctypes as C,json,sys,subprocess,time
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'tools'))
from api_contracts import describe
TYPES={'Double':C.c_double,'Long':C.c_int32,'Byte':C.c_ubyte,'LongPtr':C.c_void_p}

def run_one(name,library,data):
    catalog=json.loads((ROOT/'api/catalog.json').read_text())
    f=next(x for x in catalog['functions'] if x['name']==name);s=describe(f)
    lib=C.CDLL(str(library));lib.swe_set_ephe_path(str(data).encode());lib.swe_set_jpl_file(b'de431.eph')
    lib.swe_set_astro_models(b'0,0,0,0,0,0,0,0',258)
    lib.swe_set_delta_t_userdef.argtypes=[C.c_double];lib.swe_set_delta_t_userdef(-1e-10)
    lib.swe_set_tid_acc.argtypes=[C.c_double];lib.swe_set_tid_acc(999999.)
    lib.swe_set_lapse_rate.argtypes=[C.c_double];lib.swe_set_lapse_rate(.0065)
    lib.swe_set_interpolate_nut(0)
    lib.swe_set_sid_mode.argtypes=[C.c_int32,C.c_double,C.c_double];lib.swe_set_sid_mode(0,0,0)
    lib.swe_set_topo.argtypes=[C.c_double]*3;lib.swe_set_topo(0,0,0)
    if name=='swe_get_current_file_data':
        lib.swe_calc_ut.argtypes=[C.c_double,C.c_int32,C.c_int32,C.POINTER(C.c_double),C.c_void_p]
        lib.swe_calc_ut(2451545.,0,258,(C.c_double*6)(),C.create_string_buffer(256))
    func=getattr(lib,name);func.restype=None if f['returns']['kind']=='Sub' else TYPES[f['returns']['vbaType']]
    args=[];argtypes=[];storage={};inputs=[]
    for p in s['parameters']:
        value=p['example'];kind=TYPES[p['vbaType']]
        if value=='@DATA@':value=str(data)
        if p['direction']!='out':inputs.append(value)
        if p['pointer']:
            n=p['capacity'];array=(kind*(n+1))();guard=0xA7 if kind==C.c_ubyte else 123456789 if kind==C.c_int32 else 9.87654321e199
            array[n]=guard
            if p['direction']!='out':
                vals=list(value.encode())+[0] if p['string'] else value
                if len(vals)>n:raise ValueError('input overflow')
                for i,v in enumerate(vals):array[i]=v
            args.append(array);argtypes.append(C.POINTER(kind));storage[p['name']]=(array,guard,p)
        else:
            if p['name'] in ('hsys','c','pchar','mchar'):value=ord(value)
            args.append(value);argtypes.append(kind)
    func.argtypes=argtypes
    begin=time.monotonic();returned=func(*args);elapsed=time.monotonic()-begin
    if f['returns']['vbaType']=='LongPtr':returned=C.string_at(returned).decode(errors='replace') if returned else None
    fields=[];warning=''
    for name2,(array,guard,p) in storage.items():
        assert array[p['capacity']]==guard, f'Native buffer overrun: {name}.{name2}'
        if p['direction'] in ('out','inout'):
            if p['string']:
                raw=bytes(array[:p['capacity']]);assert b'\0' in raw
                value=raw.split(b'\0',1)[0].decode(errors='replace')
                if name2=='serr':warning=value
                else:fields.append({'label':name2,'value':value})
            else:
                indices=range(p['visible'])
                if name2 in ('cusps','cusp_speed'):indices=range(1,13)
                for i in indices:fields.append({'label':f'{name2}[{i}]','value':array[i]})
    if f['returns']['vbaType']=='LongPtr' and not any(p['string'] and p['direction']=='out' for p in s['parameters']):fields.insert(0,{'label':'Name','value':returned})
    if fields and f['returns']['vbaType']!='LongPtr' and f['returns']['kind']=='Function' and s['resultKind'] in ('value','crossing'):fields.insert(0,{'label':'Result','value':returned})
    if not fields and f['returns']['kind']=='Function':fields=[{'label':'Result','value':returned}]
    return {'name':name,'worksheetName':s['worksheetName'],'command':s['command'],'inputs':inputs,'nativeReturn':0 if returned is None and f['returns']['kind']=='Sub' else returned,'warning':warning,'fields':fields,'seconds':elapsed,'bufferGuards':'passed'}

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--library',type=Path,default=ROOT/'build/macos/libswexcel-reference.dylib');p.add_argument('--data',type=Path,default=ROOT/('dist/SWExcel-'+(ROOT/'VERSION').read_text().strip()+'/runtime/ephe'));p.add_argument('--one');p.add_argument('--output',type=Path,default=ROOT/'build/native-api-reference.json');a=p.parse_args()
    if a.one:
        print(json.dumps(run_one(a.one,a.library.resolve(),a.data.resolve())));sys.exit()
    cases=[]
    for f in json.loads((ROOT/'api/catalog.json').read_text())['functions']:
        try:
            completed=subprocess.run([sys.executable,__file__,'--library',str(a.library),'--data',str(a.data),'--one',f['name']],capture_output=True,text=True,timeout=90)
            if completed.returncode:raise RuntimeError(completed.stderr[-1000:])
            case=json.loads(completed.stdout)
        except Exception as error:case={'name':f['name'],'error':str(error)}
        cases.append(case);a.output.write_text(json.dumps({'cases':cases},indent=2)+'\n')
        print(case['name'], 'ERROR '+case['error'] if 'error' in case else f"return={case['nativeReturn']} warning={case['warning'][:100]}",flush=True)
    if any('error' in c for c in cases):sys.exit(1)
