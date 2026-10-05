#!/usr/bin/env python3
"""Generate reviewed safe marshaling, named worksheet entry points and API docs."""
import argparse,json
from pathlib import Path
from api_contracts import describe
from api_fields import fields
ROOT=Path(__file__).resolve().parents[1]

def q(s):return '"'+s.replace('"','""')+'"'
def procedure(f,spec):
    name=f['name']; public=spec['worksheetName']; inputs=[p for p in spec['parameters'] if p['direction']!='out']
    args=[f'ByVal v_{p["name"]} As Variant' for p in inputs]
    args += ['Optional ByVal options As Variant','Optional ByVal detail As Boolean = False'] if not spec['command'] else []
    start='Public Function '+public+'('+', '.join(args)+') As Variant'
    vals='Array('+', '.join('v_'+p['name'] for p in inputs)+')'
    lines=[start,f'    {public} = SWApiExecute({q(name)}, {vals}, '+('Empty, True' if spec['command'] else 'options, detail')+')','End Function','']
    # Setters and close are commands: a Sub cannot be invoked as a worksheet UDF.
    if spec['command'] and name!='swe_get_current_file_data':
        lines[0]=lines[0].replace('Public Function','Public Sub').replace(') As Variant',')')
        lines[1]='    Dim response As Variant\n    response = SWApiExecute('+q(name)+', '+vals+', Empty, True)\n    If response(2, 2) <> "OK" Then SWRaise CStr(response(3, 2))'
        lines[2]='End Sub'
    if not spec['command']:lines.insert(1,'    If IsMissing(options) Then options = Empty')
    return '\n'.join(lines)

def call_body(f,spec,index):
    name=f['name']; params=spec['parameters']; locals=[]; setup=[]; args=[]; outputs=[]; position=0
    for p in params:
        arg=p['name']; v='p_'+arg; t=p['vbaType']; n=p['capacity']; direction=p['direction']; primary='False' if direction=='inout' and arg!='ascmc' else 'True'
        if p['pointer']:
            locals.append(f'    Dim {v}(0 To {n-1}) As {t}')
            args.append(v+'(0)')
            if direction!='out':
                if p['string']:setup.append(f'    SWApiPutText {v}, CStr(values({position}))')
                else:setup.append(f'    SWApiPutVector {v}, values({position}), {n}')
                position+=1
            if direction in ('out','inout') and arg!='serr':
                if p['string']:outputs.append(f'    SWApiAdd result, {q(arg)}, SWApiOutputText({v}, {q(name)}), {q(p["units"])}, {primary}')
                else:
                    count=str(p['visible']);lo='0'
                    if arg in ('cusps','cusp_speed'):lo='1';count='SWApiHouseCount(p_hsys) + 1'
                    indexed=fields(name,p)
                    unit=q(p['units']) if indexed is None else 'CStr(Array('+', '.join(q(x) for x in (indexed+['reserved / not implemented']*p['visible'])[:p['visible']])+')(k))'
                    outputs.extend([f'    For k = {lo} To {count} - 1',f'        SWApiAdd result, {q(arg+"[")} & CStr(k) & "]", {v}(k), {unit}, {primary}','    Next k'])
        else:
            locals.append(f'    Dim {v} As {t}');args.append(v)
            convert='SWApiNumber' if t=='Double' else 'SWApiInteger'
            if arg in ('hsys','c','pchar','mchar'):convert='SWApiCharacter'
            setup.append(f'    {v} = {convert}(values({position}))');position+=1
    ret=f['returns'];native=name if name.startswith('Native_') else 'Native_'+name
    if ret['kind']=='Function':locals.append('    Dim nativeResult As '+ret['vbaType'])
    lines=[f'Public Sub SWApiCall{index}(ByRef result As SWApiResult, ByVal values As Variant)',*locals,'    Dim k As Long',*setup]
    # Validate potentially unsafe/ill-defined engine inputs before native entry.
    names={p['name'] for p in params}
    for latitude in names & {'geolat'}:lines.append(f'    If Abs(p_{latitude}) > 90# Then SWRaise "Latitude must be within -90..90 degrees."')
    if 'geopos' in names and next(p for p in params if p['name']=='geopos')['direction']!='out':lines.append('    SWApiValidateObserver p_geopos')
    if 'dgeo' in names:lines.append('    SWApiValidateObserver p_dgeo')
    if 'gregflag' in names:lines.append('    If p_gregflag <> 0 And p_gregflag <> 1 Then SWRaise "Calendar must be 0 or 1."')
    if 'ifno' in names:lines.append('    If p_ifno < 0 Or p_ifno > 4 Then SWRaise "File index must be 0..4."')
    if 'hsys' in names:lines.append('    SWApiValidateHouse p_hsys')
    if 'dobs' in names:lines.append('    SWApiValidateHeliacal p_datm, p_dobs')
    if name in ('swe_set_astro_models','swe_get_astro_models'):lines.append('    SWApiValidateModels SWBufferText(p_samod)')
    if 'TypeEvent' in names:lines.append('    If p_TypeEvent < 1 Or p_TypeEvent > 6 Then SWRaise "Event type must be 1..6."')
    if name=='swe_house_pos':
        lines[1:1]=['    Dim sunshineCusps(0 To 36) As Double, sunshineAngles(0 To 9) As Double, sunshineStatus As Long']
        lines.extend(['    If p_hsys = Asc("I") Or p_hsys = Asc("i") Then',
          '        If IsEmpty(result.SunDeclination) Then SWRaise "Sunshine house position requires sun_declination in options."',
          '        sunshineAngles(9) = CDbl(result.SunDeclination)',
          '        sunshineStatus = Native_swe_houses_armc(p_armc, p_geolat, p_eps, p_hsys, sunshineCusps(0), sunshineAngles(0))',
          '        If sunshineStatus < 0 Then SWRaise "Sunshine reference houses are unavailable."', '    End If'])
    call=', '.join(args)
    if ret['kind']=='Sub':lines.append('    '+native+(' '+call if call else ''));lines.append('    result.NativeReturn = 0')
    else:lines.extend(['    nativeResult = '+native+'('+call+')','    result.NativeReturn = nativeResult'])
    if 'serr' in names:lines.append('    result.Warning = SWBufferText(p_serr)')
    if ret['vbaType']=='LongPtr':
        own=next((p for p in params if p['string'] and p['direction']=='out'),None)
        if own:
            lines.append(f'    If nativeResult <> VarPtr(p_{own["name"]}(0)) Then SWRaise "Unexpected native return pointer."')
            # Text already appears in owned buffer output, never expose raw addresses.
            lines.append(f'    result.NativeReturn = SWApiOutputText(p_{own["name"]}, {q(name)})')
        else:
            lines.append('    result.NativeReturn = SWApiBorrowedText(nativeResult, 256)')
            outputs.insert(0,'    SWApiAdd result, "Name", result.NativeReturn, "text"')
    # Determine status before exposing native outputs; errors never become zeroes.
    lines.append(f'    SWApiCheckReturn result, {q(spec["resultKind"])}, '+('p_jd_et' if 'jd_et' in names else 'p_jd_ut' if 'jd_ut' in names else '0#'))
    if spec['resultKind']=='flags':
        flag=next(n for n in ('iflag','flag','ifl') if n in names)
        lines.append(f'    SWApiCheckFlags result, p_{flag}')
    if outputs and ret['vbaType']!='LongPtr' and ret['kind']=='Function' and spec['resultKind'] in ('value','crossing'):
        outputs.insert(0,'    SWApiAdd result, "Result", result.NativeReturn, SWApiReturnUnit('+q(name)+')')
    lines.extend(outputs)
    if not outputs and ret['kind']=='Function':lines.append('    SWApiAdd result, "Result", result.NativeReturn, SWApiReturnUnit('+q(name)+')')
    position=0
    for p in params:
        if p['direction']=='out':continue
        if p['pointer'] and not p['string']:
            lines.extend([f'    For k = 0 To {p["capacity"]-1}',f'        SWApiAdd result, "Effective input {p["name"]}[" & CStr(k) & "]", p_{p["name"]}(k), {q(p["units"])}, False','    Next k'])
        else:
            lines.append(f'    SWApiAdd result, "Input {p["name"]}", values({position}), {q(p["units"])}, False')
        position+=1
    lines.extend(['End Sub',''])
    return '\n'.join(lines)

def generate():
    catalog=json.loads((ROOT/'api/catalog.json').read_text()); funcs=catalog['functions']; outputs={}; specs=[describe(f) for f in funcs]
    dispatch=['Public Sub SWApiDispatch(ByVal functionName As String, ByVal values As Variant, ByRef result As SWApiResult)','    Select Case functionName']
    groups={}
    for i,(f,s) in enumerate(zip(funcs,specs)):
        group=s['family'];groups.setdefault(group,[]).extend([procedure(f,s),call_body(f,s,i)])
        dispatch.extend([f'        Case {q(f["name"])}',f'            SWApiCall{i} result, values'])
    dispatch.extend(['        Case Else','            SWRaise "Unknown Swiss Ephemeris function."','    End Select','End Sub'])
    for group,code in groups.items():
        name='SWApi'+group
        outputs[ROOT/f'src/vba/{name}.bas']='\n'.join([f'Attribute VB_Name = "{name}"','Option Explicit',*(['Option Private Module'] if group=='Configuration' else []),"' Generated by tools/generate_api.py. SPDX-License-Identifier: AGPL-3.0-or-later",'',*code])+'\n'
    outputs[ROOT/'src/vba/SWApiDispatch.bas']='\n'.join(['Attribute VB_Name = "SWApiDispatchModule"','Option Explicit','Option Private Module','',*dispatch])+'\n'
    outputs[ROOT/'api/contracts.json']=json.dumps({'sourceCommit':catalog['engine']['sourceCommit'],'functions':specs},indent=2)+'\n'
    docs=['# Full native API interfaces','', 'Generated from reviewed contracts and the pinned header. `SW_SWE_*` are worksheet functions.','`SW_CMD_*` are VBA commands; the current-file query is VBA-only in intended use.','Pass input vectors as a row/column range or VBA array. Every worksheet function accepts','optional `options` (a two-column key/value array) and `detail` (TRUE for labeled diagnostics).','Configuration commands affect direct native VBA callers; worksheet calls reapply their explicit options.','', 'Units and error conventions follow the [upstream programming manual](https://www.astro.com/swisseph/swephprg.htm).','', '## Functions','']
    for f,s in zip(funcs,specs):
        ins=[p for p in s['parameters'] if p['direction']!='out']
        docs.extend([f'### {s["worksheetName"]}', '',f'Native: `{f["prototype"]}`', '',f'Family: {s["family"]}. Return handling: `{s["resultKind"]}`.'+(' **Upstream experimental.**' if f.get('experimental') else ''), '', 'Data: '+s['requiredFiles'], '', '| Argument | Direction | Capacity | Units |','| --- | --- | ---: | --- |'])
        for p in s['parameters']:docs.append(f'| `{p["name"]}` | {p["direction"]} | {p["capacity"]} | {p["units"]} |')
        for p in s['parameters']:
            indexed=fields(f['name'],p)
            if indexed:
                docs.extend(['',f'Output `{p["name"]}` indices:',''])
                docs.extend(f'- `{i}`: {label}' for i,label in enumerate(indexed[:p['visible']]))
        docs.extend(['','Example inputs in parameter order: `'+json.dumps([p['example'] for p in ins])+'`.',''])
    outputs[ROOT/'docs/API-REFERENCE.md']='\n'.join(docs)
    return outputs

if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--check',action='store_true');a=parser.parse_args()
    for p,s in generate().items():
        if a.check:
            if not p.exists() or p.read_text()!=s:raise SystemExit(f'Generated file differs: {p}')
        else:p.write_text(s)
    print('API contracts, wrappers and reference agree.' if a.check else 'Generated all 106 safe API interfaces.')
