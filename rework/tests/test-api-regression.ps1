[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$PackageDirectory,[string]$Workbook='SWExcel.xlsm',[Parameter(Mandatory=$true)][string]$Report)
$ErrorActionPreference='Stop';$ProgressPreference='SilentlyContinue'
$checks=[Collections.Generic.List[object]]::new();$x=$null;$b=$null
function Call([string]$name,[object[]]$values){
 $args=[Collections.Generic.List[object]]::new();$args.Add("'$($b.Name)'!$name");foreach($v in $values){$args.Add($v)}
 return ,$x.GetType().InvokeMember('Run',[Reflection.BindingFlags]::InvokeMethod,$null,$x,$args.ToArray())
}
function Check($name,[bool]$passed,$actual){$checks.Add(@{name=$name;passed=$passed;actual=$actual});Write-Output "$name : $passed"}
function Near($a,$b,$tol=1e-8){return [math]::Abs([double]$a-[double]$b) -le $tol}
function Options($pairs){if($pairs[0] -isnot [Array]){$pairs=,@($pairs)};$o=New-Object 'object[,]' ($pairs.Count,2);for($i=0;$i -lt $pairs.Count;$i++){$o[$i,0]=$pairs[$i][0];$o[$i,1]=$pairs[$i][1]};return ,$o}
function Detail($name,$values){$wrapper='SW_SWE_'+$name.Substring(4).ToUpperInvariant();return ,(Call $wrapper (@($values)+@([Type]::Missing,$true)))}
try {
 $x=New-Object -ComObject Excel.Application;$x.Visible=$false;$x.DisplayAlerts=$false;$x.EnableEvents=$false;$x.AutomationSecurity=1
 $b=$x.Workbooks.Open((Join-Path $PackageDirectory $Workbook))
 $a=Call 'SW_UTC_JD' @(2000,1,1,17,30,0.0,5.5);$z=Call 'SW_UTC_JD' @(2000,1,1,12,0,0.0,0.0);$m=Call 'SW_UTC_JD' @(2000,1,1,8,30,0.0,-3.5)
 Check 'Fractional positive and negative UTC offsets' ((Near $a[1,1] $z[1,1]) -and (Near $m[1,2] $z[1,2])) @($a[1,1],$m[1,2])
 $leap=Call 'SW_UTC_JD' @(2016,12,31,23,59,60.0);$back=Call 'SW_SWE_JDET_TO_UTC' @($leap[1,1],1)
 Check 'UTC leap second round trip' ($back[1,1] -eq 2016 -and $back[1,3] -eq 31 -and (Near $back[1,6] 60 0.0001)) @($back[1,1],$back[1,6])
 $bad=Call 'SW_UTC_JD' @(2023,2,29,12,0,0.0)
 Check 'Invalid civil date rejected' ($bad -is [Runtime.InteropServices.ErrorWrapper] -or [string]$bad -eq '-2146826273') ([string]$bad)
 $jul=Call 'SW_JULDAY' @(1900,2,29,0.0,0);$greg=Call 'SW_JULDAY' @(2000,2,29,0.0,1)
 Check 'Julian and Gregorian leap dates' ((Near $jul 2415091.5) -and (Near $greg 2451603.5)) @($jul,$greg)
 Check 'Negative angle wraps' ((Call 'SW_SWE_DEGNORM' @(-721.0)) -eq 359.0) (Call 'SW_SWE_DEGNORM' @(-721.0))
 $cs=Call 'SW_SWE_CS2DEGSTR' @(1234567)
 Check 'UTF-8 degree symbol decoded' ([string]$cs -eq (' 3'+[char]176+"25'45")) $cs
 $houses=Call 'SW_HOUSES' @(2451545.0,52.52,13.405,'G')
 Check 'Gauquelin 36 sector buffer' ($houses.GetLength(1) -eq 36) $houses.GetLength(1)
 foreach($sys in @('I','i','J','P','W')){$h=Call 'SW_HOUSES' @(2451545.0,52.52,13.405,$sys);Check "House system $sys" ($h -is [Array] -and $h.GetLength(1) -eq 12) ([string]$h[1,1])}
 $polar=Call 'SW_HOUSES' @(2451545.0,89.0,13.405,'P',0,[Type]::Missing,$true)
 Check 'Unavailable polar houses surfaced' ($polar[2,2] -ne 'OK') $polar[2,2]
 $sun=Options @(@('sun_declination',-23.0));$p=Call 'SW_SWE_HOUSE_POS' @(280.0,52.52,23.439,'I',[double[]]@(280,1),$sun,$true)
 Check 'Sunshine house position explicit declination' ($p[2,2] -eq 'OK') $p[2,2]
 $sid=Options @(@('sidereal_mode',1));$user=Options @(@('sidereal_mode',255),@('sidereal_epoch',2451545.0),@('ayanamsa_epoch',24.0))
 $a=Call 'SW_SWE_CALC_UT' @(2451545.0,0,65794);$c=Call 'SW_SWE_CALC_UT' @(2451545.0,0,65794,$sid);$u=Call 'SW_SWE_CALC_UT' @(2451545.0,0,65794,$user);$again=Call 'SW_SWE_CALC_UT' @(2451545.0,0,65794)
 Check 'Sidereal state reset and custom epoch' ($a -is [Array] -and $c -is [Array] -and $again -is [Array] -and (Near $a[1,1] $again[1,1] 1e-10) -and [math]::Abs(($a[1,1]) - ($c[1,1])) -gt 0.1 -and $u -is [Array]) @($a[1,1],$c[1,1],$u[1,1],$again[1,1])
 $observer=Options @(@('longitude',13.405),@('latitude',52.52),@('altitude',34.0));$a=Call 'SW_SWE_CALC_UT' @(2451545.0,1,33026);$c=Call 'SW_SWE_CALC_UT' @(2451545.0,1,33026,$observer);$again=Call 'SW_SWE_CALC_UT' @(2451545.0,1,33026)
 Check 'Topocentric state reset' ($a -is [Array] -and $c -is [Array] -and $again -is [Array] -and (Near $a[1,1] $again[1,1] 1e-10) -and [math]::Abs(($a[1,1]) - ($c[1,1])) -gt .01) @($a[1,1],$c[1,1],$again[1,1])
 $dt=Call 'SW_SWE_DELTAT' @(2451545.0);$dtOpt=Options @(@('delta_t',0.001));$custom=Call 'SW_SWE_DELTAT' @(2451545.0,$dtOpt);$reset=Call 'SW_SWE_DELTAT' @(2451545.0)
 Check 'User delta T reset' ((Near $custom .001) -and (Near $dt $reset 1e-12)) @($dt,$custom,$reset)
 $series=Call 'SW_POSITIONS' @(2451545.0,1.0,5,0);$last=Call 'SW_SWE_CALC_UT' @(2451549.0,0,258)
 Check 'Date-series helper matches scalar endpoint' ($series.GetLength(0) -eq 5 -and (Near $series[5,2] $last[1,1])) $series[5,2]
 $cross=Call 'SW_SWE_MOONCROSS_NODE_UT' @(2451545.0,258)
 Check 'Node crossing retains date and coordinates' ($cross.GetLength(1) -eq 3 -and $cross[1,1] -gt 2451545) $cross[1,1]
 $star=Call 'SW_SWE_FIXSTAR2_UT' @('Spica',2451545.0,258)
 Check 'Compact star result has six numeric values' ($star.GetLength(1) -eq 6 -and $star[1,1] -is [double]) $star.GetLength(1)
 foreach($test in @(
  @{name='Unknown star';fn='swe_fixstar2_ut';args=@('not-a-real-star-xyz',2451545.0,258)},
  @{name='Wrong vector size';fn='swe_cotrans';args=@([double[]]@(1,2),23.4)},
  @{name='NUL in string';fn='swe_fixstar2_ut';args=@(('Spica'+[char]0+'extra'),2451545.0,258)},
  @{name='Oversize string';fn='swe_fixstar2_ut';args=@(('x'*256),2451545.0,258)},
  @{name='Unknown house code';fn='swe_houses';args=@(2451545.0,52.52,13.405,'?')},
  @{name='Missing optional Eros file';fn='swe_calc_ut';args=@(2451545.0,10433,258)},
  @{name='Invalid model ID';fn='swe_get_astro_models';args=@('99,0,0,0,0,0,0,0',258)}
 )){$r=Detail $test.fn $test.args;Check $test.name ($r[2,2] -eq 'ERROR') $r[3,2]}
 $r=Detail 'swe_calc_ut' @(2451545.0,0,257);Check 'Missing JPL fallback surfaced' ($r[2,2] -eq 'FALLBACK') $r[3,2]
 $r=Detail 'swe_sol_eclipse_how' @(2451545.0,2,[double[]]@(13.405,52.52,34));Check 'No eclipse is not zero success' ($r[2,2] -eq 'NO_EVENT') $r[2,2]
 $r=Detail 'swe_rise_trans' @(2451727.0,0,'',2,1,[double[]]@(0,89,0),1013.25,15.0);Check 'Circumpolar no-rise status' ($r[2,2] -eq 'NO_EVENT') $r[2,2]
 $r=Detail 'swe_calc_ut' @(2360235.5,0,258);Check 'Outside bundled dates reports fallback' ($r[2,2] -eq 'FALLBACK') $r[3,2]
 $r=Detail 'swe_calc_ut' @(9999999.0,0,258);Check 'Outside engine dates reports error' ($r[2,2] -eq 'ERROR') $r[3,2]
 $badOptions=Options @(@('unknown_option',1));$r=Call 'SW_SWE_CALC_UT' @(2451545.0,0,258,$badOptions,$true);Check 'Unknown option rejected' ($r[2,2] -eq 'ERROR') $r[3,2]
 $before=Call 'SW_SWE_GET_TID_ACC' @();[void](Call 'SW_CMD_SET_TID_ACC' @(-100.0));$after=Call 'SW_SWE_GET_TID_ACC' @();Check 'VBA setter cannot contaminate worksheet default' (Near $before $after 1e-12) @($before,$after)
 $original=$b;$peer=$null;$peerPath=Join-Path $PackageDirectory 'SWExcel-regression-peer.xlsm'
 if(Test-Path $peerPath){throw 'Peer fixture already exists.'}
 try {
  $baseline=Call 'SW_SWE_CALC_UT' @(2451545.0,0,65794)
  Copy-Item (Join-Path $PackageDirectory $Workbook) $peerPath
  $peer=$x.Workbooks.Open($peerPath);$b=$peer
  [void](Call 'SW_SWE_CALC_UT' @(2451545.0,0,65794,$sid))
  $b=$original;$after=Call 'SW_SWE_CALC_UT' @(2451545.0,0,65794)
  Check 'Two workbooks isolate explicit sidereal options' (Near $baseline[1,1] $after[1,1] 1e-10) @($baseline[1,1],$after[1,1])
 }finally{$b=$original;if($peer){$peer.Close($false)};if(Test-Path $peerPath){Remove-Item $peerPath}}
 $examples=Get-Content (Join-Path $PackageDirectory 'workbook/api-examples.json') -Raw|ConvertFrom-Json
 $x.CalculateFullRebuild()
 $recipes=$b.Worksheets.Item('Recipes')
 foreach($probe in @(@('B16',2),@('A28',7),@('B43',12),@('B48',6),@('B52',6))){
  $cell=$recipes.Range($probe[0]);$spill=$cell.SpillingToRange
  Check ('Recipe '+$probe[0]+' calculates and spills') ($cell.Value2 -is [double] -and $spill.Columns.Count -eq $probe[1]) $cell.Value2
 }
 Check 'Catalog links reach all worksheet examples' ($b.Worksheets.Item('Function Catalog').Hyperlinks.Count -eq 95) $b.Worksheets.Item('Function Catalog').Hyperlinks.Count
 foreach($sheet in $examples.sheets){foreach($example in $sheet.examples){
  $cell=$b.Worksheets.Item($sheet.name).Range($example.cell)
  $status=[string]$cell.Offset(1,1).Value2
  Check ('Worksheet example '+$example.name) ($cell.Value2 -eq 'Field' -and $status -in @('OK','WARNING','NO_EVENT','BELOW_HORIZON','FALLBACK')) @{cell=$sheet.name+'!'+$example.cell;status=$status}
 }}
} finally {
 if($b){$b.Close($false)};if($x){$x.Quit()}
 [ordered]@{time=[DateTime]::UtcNow.ToString('o');checks=@($checks.ToArray());passed=@($checks|Where-Object passed).Count;failed=@($checks|Where-Object {-not $_.passed}).Count}|ConvertTo-Json -Depth 12|Set-Content -Encoding UTF8 $Report
}
if(@($checks|Where-Object {-not $_.passed}).Count){throw 'API regressions failed.'}
