# Finalize a locally tested development package. Does not grant public-release acceptance.
[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$PackageDirectory,[Parameter(Mandatory=$true)][string]$Archive)
$ErrorActionPreference='Stop'
$root=(Resolve-Path $PackageDirectory).Path
if(Test-Path $Archive){throw 'Archive already exists; preserve the tested artifact.'}
function Read-Json($path){Get-Content -LiteralPath (Join-Path $root $path) -Raw|ConvertFrom-Json}
function Save-Json($path,$value){[IO.File]::WriteAllText((Join-Path $root $path),($value|ConvertTo-Json -Depth 30)+[Environment]::NewLine,[Text.UTF8Encoding]::new($false))}
$verificationDirectory=Split-Path -Parent (Read-Json 'verification/current-api.json').report
$version=(Get-Content (Join-Path $root 'VERSION') -Raw).Trim()
$build=Read-Json 'evidence/build.json'
$compile=Read-Json 'evidence/compile-full-api.json'
$api=Read-Json 'evidence/full-api.json'
$regression=Read-Json 'evidence/api-regression.json'
$desktop=Read-Json 'evidence/desktop-acceptance.json'
if($compile.result -ne 'passed' -or $build.sourceParityStatus -ne 'passed' -or $build.integrationSummary -ne 'PASS=18;FAIL=0'){throw 'Compilation/source/smoke evidence did not pass.'}
if($api.passed -ne 106 -or $api.failed -ne 0 -or $regression.passed -lt 136 -or $regression.failed -ne 0 -or $desktop.passed -ne 25 -or $desktop.failed -ne 0){throw 'Full API/regression/desktop evidence did not pass.'}
foreach($module in $build.sourceParity){
 if((Get-FileHash (Join-Path $root $module.source) -Algorithm SHA256).Hash.ToLowerInvariant() -ne $module.sourceSha256){throw "Source changed after verification: $($module.source)"}
}
$security=Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$attestation=[ordered]@{version=$version;time=[DateTime]::UtcNow.ToString('o');workbook='SWExcel.xlsm';workbookSha256=(Get-FileHash (Join-Path $root 'SWExcel.xlsm') -Algorithm SHA256).Hash.ToLowerInvariant();engineSha256=(Get-FileHash (Join-Path $root 'runtime\engine\swexcel-se-2.10.3b-x64.dll') -Algorithm SHA256).Hash.ToLowerInvariant();sourceParity=@($build.sourceParity).Count;compiled=$true;nativeInterfacesPassed=$api.passed;apiRegressionPassed=$regression.passed;smokePassed=18;desktopPassed=$desktop.passed;publicReleaseAcceptance='deferred';AccessVBOM=$security.AccessVBOM;VBAWarnings=$security.VBAWarnings}
Save-Json (Join-Path $verificationDirectory 'package-attestation.json') $attestation
Save-Json 'checkpoint-status.json' ([ordered]@{version=$version;engineSourceBuild='passed';vbaCompilation='passed';nativeInterfaces=106;worksheetFunctions=95;vbaCommands=11;localDevelopmentVerification='passed';publicReleaseReady=$false;downloadOnboarding='deferred'})
$files=@(Get-ChildItem $root -Recurse -File|Where-Object Name -ne 'package-files.json'|Sort-Object FullName|ForEach-Object{[ordered]@{path=$_.FullName.Substring($root.Length+1).Replace('\','/');sha256=(Get-FileHash $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant();size=$_.Length}})
# Ordinal order matches the cross-platform manifest contract.
$ordered=[Collections.Generic.SortedDictionary[string,object]]::new([StringComparer]::Ordinal)
foreach($item in $files){$ordered.Add([string]$item.path,$item)}
$files=@($ordered.Values)
Save-Json 'package-files.json' ([ordered]@{schemaVersion=1;projectVersion=$version;files=$files})
# Write portable forward-slash ZIP entry names, including dotfiles.
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$stream=[IO.File]::Open($Archive,[IO.FileMode]::CreateNew)
$zip=New-Object IO.Compression.ZipArchive($stream,[IO.Compression.ZipArchiveMode]::Create,$false)
try {
 foreach($file in Get-ChildItem $root -Recurse -File){
  $entry=(Split-Path $root -Leaf)+'/'+$file.FullName.Substring($root.Length+1).Replace('\','/')
  [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip,$file.FullName,$entry,[IO.Compression.CompressionLevel]::Optimal)
 }
}finally{$zip.Dispose();$stream.Dispose()}
$attestation|ConvertTo-Json
