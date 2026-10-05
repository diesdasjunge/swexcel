[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$PackageDirectory,[string]$Workbook='SWExcel.xlsm',[Parameter(Mandatory=$true)][string]$Reference,[Parameter(Mandatory=$true)][string]$Report)
$ErrorActionPreference='Stop';$ProgressPreference='SilentlyContinue'
$cases=(Get-Content $Reference -Raw|ConvertFrom-Json).cases
$contracts=(Get-Content (Join-Path $PackageDirectory 'api/contracts.json') -Raw|ConvertFrom-Json).functions
$checks=[Collections.Generic.List[object]]::new();$x=$null;$b=$null
function Invoke-Macro([string]$name,[object[]]$values){
 $invokeArgs=[Collections.Generic.List[object]]::new();$invokeArgs.Add("'$($b.Name)'!$name")
 foreach($v in $values){$invokeArgs.Add($v)}
 return ,$x.GetType().InvokeMember('Run',[Reflection.BindingFlags]::InvokeMethod,$null,$x,$invokeArgs.ToArray())
}
try {
 $x=New-Object -ComObject Excel.Application;$x.Visible=$false;$x.DisplayAlerts=$false;$x.EnableEvents=$false;$x.AutomationSecurity=1
 $b=$x.Workbooks.Open((Join-Path $PackageDirectory $Workbook))
 foreach($case in $cases){
  $row=[ordered]@{name=$case.name;passed=$false;status=$null;nativeReturn=$null;differences=@();error=$null}
  try {
   if($case.error){throw "Reference failed: $($case.error)"}
   $inputList=[Collections.Generic.List[object]]::new()
   foreach($v in $case.inputs){
    if($v -is [Array]){$inputList.Add([double[]]$v)}
    elseif($case.name -eq 'swe_set_ephe_path'){$inputList.Add((Join-Path $PackageDirectory 'runtime\ephe'))}
    else{$inputList.Add($v)}
   }
   if($case.name -eq 'swe_get_current_file_data'){[void](Invoke-Macro 'SW_CALC_UT' @(2451545.0,0))}
   if($case.command){$result=Invoke-Macro 'SWApiExecute' @($case.name,$inputList.ToArray(),[Type]::Missing,$true)}
   else {$inputList.Add([Type]::Missing);$inputList.Add($true);$result=Invoke-Macro $case.worksheetName $inputList.ToArray()}
   $row.status=[string]$result[2,2];$row.nativeReturn=$result[5,2]
   if($case.command){[void](Invoke-Macro $case.worksheetName $inputList.ToArray());$row.commandInvoked=$true}
   $kind=($contracts|Where-Object name -eq $case.name).resultKind
   if($case.nativeReturn -is [ValueType]){
    if([math]::Abs([double]$row.nativeReturn-[double]$case.nativeReturn) -gt 1e-8){$row.differences+=@{field='nativeReturn';actual=$row.nativeReturn;expected=$case.nativeReturn}}
   }elseif([string]$row.nativeReturn -cne [string]$case.nativeReturn){$row.differences+=@{field='nativeReturn';actual=$row.nativeReturn;expected=$case.nativeReturn}}
   if($kind -eq 'event' -and $case.nativeReturn -eq 0 -and $row.status -ne 'NO_EVENT'){throw 'Zero event result must report NO_EVENT.'}
   if($kind -eq 'rise' -and $case.nativeReturn -eq -2 -and $row.status -ne 'NO_EVENT'){throw 'Circumpolar result must report NO_EVENT.'}
   if($row.status -eq 'ERROR'){throw ([string]$result[3,2])}
   $actual=@{};for($r=9;$r -le $result.GetUpperBound(0);$r++){$actual[[string]$result[$r,1]]=$result[$r,2]}
   if($row.status -notin @('NO_EVENT','BELOW_HORIZON')){
    foreach($field in $case.fields){
     if(-not $actual.ContainsKey($field.label)){throw "Missing output $($field.label)"}
     $v=$actual[$field.label];$expected=$field.value
     if($expected -is [ValueType]){
      if($v -isnot [ValueType]){throw "Non-numeric output $($field.label): $v"}
      $delta=[math]::Abs([double]$v-[double]$expected)
      $tolerance=1e-7
      if([math]::Abs([double]$expected) -gt 1e6){$tolerance=1e-8}
      if($delta -gt $tolerance){$row.differences+=@{field=$field.label;actual=$v;expected=$expected;difference=$delta;tolerance=$tolerance}}
     } elseif([string]$v -cne [string]$expected){$row.differences+=@{field=$field.label;actual=$v;expected=$expected}}
    }
   }
   $row.passed=$row.differences.Count -eq 0
  }catch{$row.error=$_.Exception.Message}
  $checks.Add($row)
  Write-Output "$($row.name): $($row.passed) $($row.status) $($row.error)"
  [ordered]@{time=[DateTime]::UtcNow.ToString('o');workbook=$b.FullName;checks=@($checks.ToArray());passed=@($checks|Where-Object passed).Count;failed=@($checks|Where-Object {-not $_.passed}).Count}|ConvertTo-Json -Depth 12|Set-Content -Encoding UTF8 $Report
 }
} finally {if($b){$b.Close($false)};if($x){$x.Quit()}}
if(@($checks|Where-Object {-not $_.passed}).Count){throw 'Full API comparison failed.'}
