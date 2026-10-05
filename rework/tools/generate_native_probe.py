#!/usr/bin/env python3
"""Independent typed C caller for every Windows DLL export, with buffer guards."""
import json
from pathlib import Path
from api_contracts import describe
ROOT=Path(__file__).resolve().parents[1]
def q(s):return json.dumps(s,ensure_ascii=True)
def generate():
 funcs=json.loads((ROOT/'api/catalog.json').read_text())['functions']
 code=['#include <windows.h>','#include <stdio.h>','#include <stdlib.h>','#include <string.h>','#include <math.h>','#include "swephexp.h"','static FILE *output; static const char *data_path; static int field_count; static UINT text_codepage=CP_ACP;']
 for f in funcs:
  ret=f['returns']['cType'];args=', '.join(p['cType'] for p in f['parameters']) or 'void'
  code.extend([f'typedef {ret} (*FN_{f["name"]})({args});',f'static FN_{f["name"]} fn_{f["name"]};'])
 code.extend([r'''
static void text_json(const char *s) {
  wchar_t wide[32768]; int i;
  if (!s) { fputs("null", output); return; }
  if (!MultiByteToWideChar(text_codepage,0,s,-1,wide,32768)) exit(3);
  fputc('"',output);
  for(i=0;wide[i];i++) {
    unsigned int c=wide[i];
    if(c=='"'||c=='\\') {fputc('\\',output);fputc(c,output);}
    else if(c<32||c>126) fprintf(output,"\\u%04x",c);
    else fputc(c,output);
  }
  fputc('"',output);
}
static void number_json(double n) {if(_finite(n)) fprintf(output,"%.17g",n); else fputs("null",output);}
static void field_prefix(const char *name) {if(field_count++)fputc(',',output);fputs("{\"label\":",output);text_json(name);fputs(",\"value\":",output);}
static void number_field(const char *name,double value){field_prefix(name);number_json(value);fputc('}',output);}
static void text_field(const char *name,const char *value){field_prefix(name);text_json(value);fputc('}',output);}
static void reset_options(void) {
 fn_swe_set_ephe_path(data_path);fn_swe_set_jpl_file("de431.eph");
 fn_swe_set_astro_models("0,0,0,0,0,0,0,0",258);fn_swe_set_delta_t_userdef(-1e-10);
 fn_swe_set_tid_acc(999999);fn_swe_set_lapse_rate(0.0065);fn_swe_set_interpolate_nut(0);
 fn_swe_set_sid_mode(0,0,0);fn_swe_set_topo(0,0,0);
}
'''])
 for f in funcs:
  s=describe(f);name=f['name'];args=[];init=[];guards=[];fields=[]
  code.append(f'static void probe_{name}(void) {{')
  inputs=[p['example'] for p in s['parameters'] if p['direction']!='out']
  header={'name':name,'worksheetName':s['worksheetName'],'command':s['command'],'inputs':inputs}
  for p in s['parameters']:
   arg=p['name'];v='p_'+arg;t=p['cType'].replace('const ','').replace(' *','');value=p['example'];n=p['capacity']
   if p['pointer']:
    code.append(f' {t} {v}[{n+1}]={{0}};')
    guard='87' if t=='char' else '123456789' if t in ('int','int32','AS_BOOL','centisec','CSEC') else '9.87654321e199'
    init.append(f' {v}[{n}]={guard};');guards.append(f' if({v}[{n}]!={guard}){{fprintf(stderr,"Buffer overwrite: {name}.{arg}\\n");exit(4);}}')
    if p['direction']!='out':
     if p['string']:init.append(f' strcpy({v}, '+('data_path' if value=='@DATA@' else q(value))+');')
     else:
      for i,val in enumerate(value):init.append(f' {v}[{i}]={repr(val)};')
    args.append(v)
    if p['direction'] in ('out','inout') and arg!='serr':
     if p['string']:fields.append(f' text_field({q(arg)},{v});')
     else:
      indices=range(p['visible']) if arg not in ('cusps','cusp_speed') else range(1,13)
      for i in indices:fields.append(f' number_field({q(arg+"["+str(i)+"]")},{v}[{i}]);')
   else:
    if arg in ('hsys','c','pchar','mchar'):value=ord(value)
    code.append(f' {t} {v}={repr(value)};');args.append(v)
  if f['returns']['kind']=='Function':code.append(f' {f["returns"]["cType"]} returned;')
  code.extend([' reset_options();',' text_codepage='+('CP_UTF8;' if name=='swe_cs2degstr' else 'CP_ACP;'),*init])
  if name=='swe_get_current_file_data':code.append(' {double xx[6];char err[256];fn_swe_calc_ut(2451545,0,258,xx,err);}')
  code.append(' '+('returned=' if f['returns']['kind']=='Function' else '')+'fn_'+name+'('+', '.join(args)+');')
  code.extend(guards)
  code.append(' fputs('+q(json.dumps(header,separators=(',',':'))[:-1]+',"nativeReturn":')+',output);')
  if f['returns']['kind']=='Sub':code.append(' number_json(0);')
  elif f['returns']['vbaType']=='LongPtr':code.append(' text_json(returned);')
  else:code.append(' number_json((double)returned);')
  code.append(' fputs(",\\\"warning\\\":",output); text_json('+('p_serr' if any(p['name']=='serr' for p in s['parameters']) else '""')+');')
  if f['returns']['vbaType']=='LongPtr' and not any(p['string'] and p['direction']=='out' for p in s['parameters']):fields.insert(0,' text_field("Name",returned);')
  if fields and f['returns']['vbaType']!='LongPtr' and f['returns']['kind']=='Function' and s['resultKind'] in ('value','crossing'):fields.insert(0,' number_field("Result",(double)returned);')
  if not fields and f['returns']['kind']=='Function':fields.append(' number_field("Result",(double)returned);')
  code.extend([' fputs('+q(',"fields":[')+',output);field_count=0;',*fields,' fputs('+q('],"bufferGuards":"passed"}')+',output);fflush(output);','}'])
 code.extend(['int main(int argc,char **argv) { HMODULE lib; if(argc!=4)return 2;data_path=argv[2];output=fopen(argv[3],"wb");if(!output)return 2;lib=LoadLibraryExA(argv[1],NULL,LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR|LOAD_LIBRARY_SEARCH_SYSTEM32);if(!lib)return 2;'])
 for f in funcs:code.append(f' fn_{f["name"]}=(FN_{f["name"]})GetProcAddress(lib,"{f["name"]}");if(!fn_{f["name"]})return 2;')
 code.append(' fputs('+q('{"platform":"Windows MSVC x64 typed DLL caller","cases":[')+',output);')
 for i,f in enumerate(funcs):code.append((' fputc(\',\',output);' if i else '')+f' probe_{f["name"]}();')
 code.extend([' fputs("]}",output);fclose(output);return 0;}'])
 return '\n'.join(code)+'\n'
if __name__=='__main__':
 out=ROOT/'build/native-api-probe.c';out.parent.mkdir(exist_ok=True);out.write_text(generate());print(out)
