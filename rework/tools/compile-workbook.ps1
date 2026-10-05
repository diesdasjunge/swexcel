# Explicit full-project compilation in a dedicated desktop Excel instance.
# Requires the developer's existing VBA-project access; never changes trust.
[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$Workbook,[Parameter(Mandatory=$true)][string]$Report)
$ErrorActionPreference='Stop';$ProgressPreference='SilentlyContinue'
Add-Type @'
using System; using System.Runtime.InteropServices;
public static class SWCompileWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint process); }
'@
$x=$null;$b=$null;$watch=$null;$failure=$null
$result=[ordered]@{time=[DateTime]::UtcNow.ToString('o');workbook=$Workbook;result='not_completed';enabledBefore=$null;enabledAfter=$null;error=$null}
try {
 $x=New-Object -ComObject Excel.Application;$x.Visible=$false;$x.DisplayAlerts=$false;$x.EnableEvents=$false;$x.AutomationSecurity=1
 $b=$x.Workbooks.Open((Resolve-Path $Workbook).Path)
 $x.VBE.ActiveVBProject=$b.VBProject
 $command=$x.VBE.CommandBars.FindControl(1,578)
 if($null -eq $command){throw 'Compile VBAProject command unavailable.'}
 $result.command=[string]$command.Caption;$result.enabledBefore=[bool]$command.Enabled
 [uint32]$excelPid=0;[void][SWCompileWindow]::GetWindowThreadProcessId([IntPtr]$x.Hwnd,[ref]$excelPid)
 # A compilation error is a modal dialog even when DisplayAlerts is false.
 # Dismiss only that known error notification, retain its text, never accept
 # security, license, or unrelated application dialogs.
 $watch=Start-Job -ArgumentList $excelPid -ScriptBlock {
  param($targetPid)
  Add-Type -AssemblyName UIAutomationClient
  $condition=New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ProcessIdProperty,[int]$targetPid)
  for($attempt=0;$attempt -lt 60;$attempt++){
   $windows=[System.Windows.Automation.AutomationElement]::RootElement.FindAll([System.Windows.Automation.TreeScope]::Children,$condition)
   foreach($window in $windows){
    $nodes=$window.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition)
    $errorNode=@($nodes|Where-Object {$_.Current.Name -like 'Compile error:*'})
    if($errorNode.Count){
     $errorNode[0].Current.Name
     $ok=@($nodes|Where-Object {$_.Current.Name -eq 'OK'})|Select-Object -First 1
     if($ok){
      $patterns=@($ok.GetSupportedPatterns()|ForEach-Object ProgrammaticName)
      if($patterns -contains 'InvokePatternIdentifiers.Pattern'){$ok.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()}
      elseif($patterns -contains 'LegacyIAccessiblePatternIdentifiers.Pattern'){$ok.GetCurrentPattern([System.Windows.Automation.LegacyIAccessiblePattern]::Pattern).DoDefaultAction()}
     }
     return
    }
   }
   Start-Sleep -Milliseconds 500
  }
 }
 if($result.enabledBefore){$command.Execute()}
 $result.enabledAfter=[bool]$command.Enabled
 if($result.enabledAfter){
  $pane=$x.VBE.ActiveCodePane
  $line=0;$column=0;$endLine=0;$endColumn=0
  $pane.GetSelection([ref]$line,[ref]$column,[ref]$endLine,[ref]$endColumn)
  $result.module=$pane.CodeModule.Name;$result.line=$line;$result.source=$pane.CodeModule.Lines($line,1)
  throw 'Compile command remains enabled.'
 }
 $b.Save();$result.result='passed'
} catch {$failure=$_;$result.result='failed';$result.error=$_.Exception.Message}
finally {
 if($watch){Stop-Job $watch;$messages=@(Receive-Job $watch);if($messages.Count){$result.dialog=$messages -join "`n"};Remove-Job $watch}
 if($b){$b.Close($false)};if($x){$x.Quit()}
 $result|ConvertTo-Json -Depth 5|Set-Content -Encoding UTF8 $Report
}
$result|ConvertTo-Json -Depth 5
if($failure){throw $failure}
