[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$PackageDirectory, [string]$Workbook='SWExcel-verified.xlsm', [Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
if(Test-Path $OutputDirectory){throw 'Use a fresh acceptance output directory.'}
[void](New-Item -ItemType Directory $OutputDirectory)
$root=(Resolve-Path $PackageDirectory).Path
$checks=[Collections.Generic.List[object]]::new()
function Check($name,[bool]$passed,$actual){$checks.Add([ordered]@{name=$name;passed=$passed;actual=$actual})}
function New-Session($path){
 $x=New-Object -ComObject Excel.Application
 $x.Visible=$false; $x.DisplayAlerts=$false; $x.EnableEvents=$false
 try{$b=$x.Workbooks.Open($path); return @{excel=$x;book=$b;prefix="'"+$b.Name.Replace("'","''")+"'!"}}catch{$x.Quit();throw}
}
function Close-Session($s){if($s){$s.book.Close($false);$s.excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($s.book);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($s.excel);[GC]::Collect();[GC]::WaitForPendingFinalizers()}}
function Copy-Fixture($name){
 $p=Join-Path $OutputDirectory $name
 [void](New-Item -ItemType Directory $p)
 Copy-Item (Join-Path $root 'runtime') $p -Recurse
 Copy-Item (Join-Path $root $Workbook) (Join-Path $p $Workbook)
 return $p
}
function Probe-Fixture($name,$path,$pattern,$runSmoke=$false){
 $s=$null
 try{$s=New-Session (Join-Path $path $Workbook);$status=[string]$s.excel.Run($s.prefix+'SW_RUNTIME_STATUS'); Check $name ($status -match $pattern) $status
 if($runSmoke){$summary=[string]$s.excel.Run($s.prefix+'SW_RunIntegrationChecks');Check "$name integration" ($summary -eq 'PASS=18;FAIL=0') $summary
 $actualPath=[string]$s.excel.Run($s.prefix+'SW_ENGINE_PATH');Check "$name engine path" ($actualPath -ieq (Join-Path $path 'runtime\engine\swexcel-se-2.10.3b-x64.dll')) $actualPath}
 }finally{Close-Session $s}
}
$s=$null
try{
 $s=New-Session (Join-Path $root $Workbook);$x=$s.excel;$b=$s.book;$prefix=$s.prefix
 $summary=[string]$x.Run($prefix+'SW_RunIntegrationChecks');Check 'Fresh Excel integration' ($summary -eq 'PASS=18;FAIL=0') $summary
 $vector=@($x.Run($prefix+'SW_CALC_UT',2451545.0,0))
 $reference=@(& (Join-Path $root 'runtime\engine\swetest64.exe') '-b1.1.2000' '-ut12' '-p0' '-eswe' "-edir$root\runtime\ephe" '-flbRss' '-g,' '-head')
 if($LASTEXITCODE -ne 0){throw 'swetest failed'}
 $numbers=@(($reference -join '').Split(',')|Where-Object {$_.Trim().Length -gt 0}|ForEach-Object{[double]::Parse($_.Trim(),[Globalization.CultureInfo]::InvariantCulture)})
 $differences=@();$numerical=$vector.Count -eq 6 -and $numbers.Count -eq 6
 if($numerical){for($i=0;$i -lt 6;$i++){$d=[math]::Abs([double]$vector[$i]-$numbers[$i]);$differences+=$d;if($d -gt 0.0000001){$numerical=$false}}}
 Check 'Six-value swetest parity (1e-7 printed tolerance)' $numerical @{excel=$vector;reference=$numbers;absoluteDifferences=$differences}
 $a=[double]$x.Run($prefix+'SW_LONGITUDE',2451545.0,0,65794,0)
 $other=[double]$x.Run($prefix+'SW_LONGITUDE',2451545.0,0,65794,1)
 $again=[double]$x.Run($prefix+'SW_LONGITUDE',2451545.0,0,65794,0)
 Check 'Interleaved sidereal modes' ([math]::Abs($a-$again) -lt 1e-10 -and [math]::Abs($a-$other) -gt 0.1) @($a,$other,$again)
 $topo=[double]$x.Run($prefix+'SW_LONGITUDE',2451545.0,1,33026,0,13.405,52.52,34.0)
 $newyork=[double]$x.Run($prefix+'SW_LONGITUDE',2451545.0,1,33026,0,-74.006,40.7128,10.0)
 $topoAgain=[double]$x.Run($prefix+'SW_LONGITUDE',2451545.0,1,33026,0,13.405,52.52,34.0)
 Check 'Interleaved observers' ([math]::Abs($topo-$topoAgain) -lt 1e-10 -and [math]::Abs($topo-$newyork) -gt 0.001) @($topo,$newyork,$topoAgain)
 $sun=[double]$x.Run($prefix+'SW_LONGITUDE',2451545.0,0)
 Check 'Default options restored after state changes' ([math]::Abs($sun-[double]$vector[0]) -lt 1e-10) $sun
 $sheet=$b.Worksheets.Item('Integration')
 $reported=$sheet.Range('C13').Value2
 Check 'Diagnostic decimal preserved as text' ($reported -is [string] -and [math]::Abs([double]::Parse($reported.Replace(',', '.'),[Globalization.CultureInfo]::InvariantCulture)-$sun) -lt 1e-8) $reported
 $sheet.Range('C63').ClearContents();$x.CalculateFullRebuild()
 $spill=$sheet.Range('B63').SpillingToRange
 Check 'Blocked spill restores six values' ($spill.Columns.Count -eq 6 -and [math]::Abs([double]$sheet.Range('B63').Value2-$sun) -lt 1e-10) @($spill.Address(),$sheet.Range('B63').Value2)
 $sheet.Range('B63').ClearContents();$sheet.Range('C63').Value2='Intentional blockage';$sheet.Range('B63').Formula2='=SW_CALC_UT(2451545,0)'
 $sheet.Range('A100').Value2=2451545.0;$sheet.Range('B100').Formula2='=SW_LONGITUDE(A100,0)';$sheet.Range('C100').Formula2='=B100+1';$x.CalculateFullRebuild()
 $before=[double]$sheet.Range('B100').Value2;$sheet.Range('A100').Value2=2451546.0;$x.Calculate()
 $after=[double]$sheet.Range('B100').Value2;$dependent=[double]$sheet.Range('C100').Value2
 Check 'Input edit and dependent recalculation' ([math]::Abs($after-$before) -gt 0.5 -and [math]::Abs($dependent-$after-1) -lt 1e-10) @($before,$after,$dependent)
 for($i=110;$i -lt 310;$i++){$sheet.Range("B$i").Formula2="=SW_LONGITUDE(2451545+$i,0)"}
 $timer=[Diagnostics.Stopwatch]::StartNew();$x.CalculateFullRebuild();$timer.Stop();$batch=$true
 for($i=110;$i -lt 310;$i++){if($sheet.Range("B$i").Value2 -isnot [double]){$batch=$false}}
 Check '200-formula batch calculation' $batch @{milliseconds=$timer.Elapsed.TotalMilliseconds;count=200}
 $sheet.Range('A100:C309').ClearContents();$summary=[string]$x.Run($prefix+'SW_RunIntegrationChecks');Check 'Smoke after spill and recalculation edits' ($summary -eq 'PASS=18;FAIL=0') $summary
 $b.Save()
}finally{Close-Session $s}
try{
 Probe-Fixture 'Saved workbook reopened in fresh Excel' $root '^Ready' $true
 $space=Copy-Fixture 'package with spaces';Probe-Fixture 'Space path' $space '^Ready' $true
 $latin=Copy-Fixture ('M'+[char]0x00FC+'nchen');Probe-Fixture 'ANSI non-ASCII path' $latin '^Ready' $true
 $unicode=Copy-Fixture ([string][char]0x6E2C+[char]0x8A66);Probe-Fixture 'Unrepresentable Unicode path is explicit' $unicode '^ERROR: The package path cannot be represented'
 $missing=Copy-Fixture 'missing dll';Move-Item (Join-Path $missing 'runtime\engine\swexcel-se-2.10.3b-x64.dll') (Join-Path $missing 'runtime\engine\engine.held');Probe-Fixture 'Missing DLL diagnostic' $missing '^ERROR: The package DLL is missing'
 $data=Copy-Fixture 'missing data';Move-Item (Join-Path $data 'runtime\ephe\sepl_18.se1') (Join-Path $data 'runtime\ephe\sepl_18.se1.held');Probe-Fixture 'Missing data diagnostic' $data '^ERROR: Required package data is missing'
 $wrong=Copy-Fixture 'wrong architecture';$fixture='C:\Windows\SysWOW64\version.dll';$bytes=[IO.File]::ReadAllBytes($fixture);$offset=[BitConverter]::ToInt32($bytes,60);$machine=[BitConverter]::ToUInt16($bytes,$offset+4)
 if($machine -ne 0x14c){throw "Wrong-architecture fixture is not x86: $machine"}
 Copy-Item $fixture (Join-Path $wrong 'runtime\engine\swexcel-se-2.10.3b-x64.dll') -Force
 Probe-Fixture 'x86 DLL rejected by 64-bit Excel' $wrong '^ERROR: Windows could not load the x64 engine'
 $s=$null;$peer=$null
 try{
  $s=New-Session (Join-Path $root $Workbook)
  $peerName='SWExcel-peer-'+[guid]::NewGuid().ToString('N')+'.xlsm'
  Copy-Item (Join-Path $root $Workbook) (Join-Path $root $peerName)
  $peer=$s.excel.Workbooks.Open((Join-Path $root $peerName))
  $a=[double]$s.excel.Run($s.prefix+'SW_LONGITUDE',2451545.0,1,33026,0,13.405,52.52,34.0)
  $b=[double]$s.excel.Run("'$peerName'!SW_LONGITUDE",2451545.0,1,33026,0,-74.006,40.7128,10.0)
  $again=[double]$s.excel.Run($s.prefix+'SW_LONGITUDE',2451545.0,1,33026,0,13.405,52.52,34.0)
  Check 'Two workbooks sharing package isolate observer options' ([math]::Abs($a-$again) -lt 1e-10 -and [math]::Abs($a-$b) -gt 0.001) @($a,$b,$again)
 }finally{if($peer){$peer.Close($false)};Close-Session $s}
 $s=$null;$second=$null
 try{$s=New-Session (Join-Path $root $Workbook);[void]$s.excel.Run($s.prefix+'SW_VERSION');$copyPath=Join-Path $space 'SWExcel-conflict.xlsm';Copy-Item (Join-Path $space $Workbook) $copyPath;$second=$s.excel.Workbooks.Open($copyPath);$status=[string]$s.excel.Run("'SWExcel-conflict.xlsm'!SW_RUNTIME_STATUS");Check 'Different package in same Excel process rejected' ($status -match 'different SWExcel package DLL is already loaded') $status}finally{if($second){$second.Close($false)};Close-Session $s}
}finally{
 $report=[ordered]@{recordedAtUtc=[DateTime]::UtcNow.ToString('o');package=$root;workbook=$Workbook;checks=@($checks.ToArray());passed=@($checks|Where-Object passed).Count;failed=@($checks|Where-Object {-not $_.passed}).Count;manualCompile='separate UI gate';downloadOnboarding='not tested'}
 [IO.File]::WriteAllText((Join-Path $OutputDirectory 'acceptance.json'),($report|ConvertTo-Json -Depth 12),[Text.UTF8Encoding]::new($false))
}
$report|ConvertTo-Json -Depth 12
if($report.failed){throw 'Excel acceptance checks failed.'}
