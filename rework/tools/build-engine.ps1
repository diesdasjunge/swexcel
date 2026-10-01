# Build the pinned latest source; never substitute the unchanged upstream prebuilt.
# Run in an x64 Native Tools Command Prompt for Visual Studio, with Windows SDK.
[CmdletBinding()]
param(
    [string]$OutputRoot = "",
    [string]$BuildDirectory = "",
    [switch]$SkipSwetest
)
$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest
$ReworkRoot = Split-Path -Parent $PSScriptRoot
$SourceDirectory = Join-Path $ReworkRoot "vendor\swisseph\source"
if (-not $OutputRoot) {
    if (Test-Path -LiteralPath (Join-Path $ReworkRoot "checkpoint-status.json") -PathType Leaf) {
        # The script is also shipped inside the already prepared package.
        $OutputRoot = Join-Path $ReworkRoot "runtime\engine"
    } else {
        $PackageVersion = (Get-Content -LiteralPath (Join-Path $ReworkRoot "VERSION") -Raw).Trim()
        if ($PackageVersion -notmatch '^[0-9A-Za-z.-]+$') { throw "Invalid package VERSION." }
        $OutputRoot = Join-Path $ReworkRoot "dist\SWExcel-$PackageVersion\runtime\engine"
    }
}
if (-not $BuildDirectory) { $BuildDirectory = Join-Path $ReworkRoot "build\engine-x64" }
$OutputRoot = [IO.Path]::GetFullPath($OutputRoot)
$BuildDirectory = [IO.Path]::GetFullPath($BuildDirectory)
if (-not (Test-Path -LiteralPath $OutputRoot -PathType Container)) {
    throw "Prepare the package first; its runtime engine directory does not exist: $OutputRoot"
}
$Commit = "f4dcd18e8005dde95fd8a8d2312ed12f9accd1b0"
$LibraryName = "swexcel-se-2.10.3b-x64.dll"

function Invoke-CheckedNative {
    param([string]$Program, [string[]]$Arguments)
    & $Program @Arguments
    if ($LASTEXITCODE -ne 0) { throw "$Program failed with exit code $LASTEXITCODE" }
}

function Assert-Amd64Pe {
    param([string]$Path)
    $Bytes = [IO.File]::ReadAllBytes($Path)
    $Offset = [BitConverter]::ToInt32($Bytes, 60)
    if ([BitConverter]::ToUInt32($Bytes, $Offset) -ne 0x00004550 -or
        [BitConverter]::ToUInt16($Bytes, $Offset + 4) -ne 0x8664 -or
        [BitConverter]::ToUInt16($Bytes, $Offset + 24) -ne 0x20b) {
        throw "Build output is not an AMD64 PE32+ binary: $Path"
    }
}

if ($env:OS -ne "Windows_NT") { throw "This build requires Windows and MSVC." }
$Compiler = (Get-Command "cl.exe" -ErrorAction Stop).Source
$Dumpbin = (Get-Command "dumpbin.exe" -ErrorAction Stop).Source
if ($env:VSCMD_ARG_TGT_ARCH -and $env:VSCMD_ARG_TGT_ARCH -ne "x64") {
    throw "Select the x64 Native Tools environment, not $env:VSCMD_ARG_TGT_ARCH."
}
$SourceManifestPath = Join-Path $SourceDirectory "provenance.json"
$SourceManifest = Get-Content -LiteralPath $SourceManifestPath -Raw | ConvertFrom-Json
if ($SourceManifest.commit -ne $Commit) { throw "Unexpected engine source commit." }
foreach ($Entry in $SourceManifest.files) {
    $Path = Join-Path $SourceDirectory $Entry.path
    if ((Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() -ne $Entry.sha256) {
        throw "Pinned upstream source hash differs: $($Entry.path)"
    }
}
$VersionText = Get-Content -LiteralPath (Join-Path $SourceDirectory "sweph.h") -Raw
$Match = [regex]::Match($VersionText, '#define\s+SE_VERSION\s+"([^"]+)"')
if (-not $Match.Success) { throw "Pinned source has no SE_VERSION." }
$ExpectedRuntimeVersion = $Match.Groups[1].Value

New-Item -ItemType Directory -Force -Path $BuildDirectory | Out-Null
$ObjectDirectory = Join-Path $BuildDirectory "obj"
New-Item -ItemType Directory -Force -Path $ObjectDirectory | Out-Null
$SourceNames = @("swedate.c", "swehouse.c", "swejpl.c", "swemmoon.c", "swemplan.c", "sweph.c", "swephlib.c", "swecl.c", "swehel.c")
$Sources = @($SourceNames | ForEach-Object { Join-Path $SourceDirectory $_ })

# Initialize the handle used by upstream swe_get_library_path(). The pinned
# upstream source remains untouched. This local loader glue exports no API.
$LoaderPath = Join-Path $BuildDirectory "SWEngineDllMain.c"
$Loader = @'
#include <windows.h>
#include "swephexp.h"
BOOL WINAPI DllMain(HINSTANCE instance, DWORD reason, LPVOID reserved) {
    (void) reserved;
    if (reason == DLL_PROCESS_ATTACH) dllhandle = instance;
    return TRUE;
}
'@
[IO.File]::WriteAllText($LoaderPath, $Loader, [Text.Encoding]::ASCII)

# Match the upstream swedll64 Release/x64 recipe: Cdecl, MAKE_DLL, MultiByte,
# static multi-threaded CRT (/MT), WIN32/NDEBUG/_WINDOWS. x64 has one ABI.
$Common = @("/nologo", "/O2", "/W3", "/MT", "/TC", "/Gd", "/DWIN32", "/DNDEBUG", "/D_WINDOWS", "/D_CRT_SECURE_NO_WARNINGS", "/D_CRT_NONSTDC_NO_DEPRECATE", "/I$SourceDirectory", "/Fo$ObjectDirectory\")
$LibraryPath = Join-Path $OutputRoot $LibraryName
$ImportLibrary = Join-Path $BuildDirectory "swexcel-se-2.10.3b-x64.lib"
$DllArguments = $Common + @("/LD", "/DMAKE_DLL") + $Sources + @($LoaderPath, "/link", "/MACHINE:X64", "/OUT:$LibraryPath", "/IMPLIB:$ImportLibrary")
Push-Location $BuildDirectory
try {
    Invoke-CheckedNative -Program $Compiler -Arguments $DllArguments
    Assert-Amd64Pe -Path $LibraryPath
    $ExportOutput = @(& $Dumpbin "/nologo" "/exports" $LibraryPath)
    if ($LASTEXITCODE -ne 0) { throw "dumpbin export inspection failed." }
    $ActualExports = @($ExportOutput | ForEach-Object {
        if ($_ -match '^\s+\d+\s+[0-9A-Fa-f]+\s+[0-9A-Fa-f]+\s+(swe_\w+)') { $Matches[1] }
    } | Sort-Object)
    $Catalog = Get-Content -LiteralPath (Join-Path $ReworkRoot "api\catalog.json") -Raw | ConvertFrom-Json
    $ExpectedExports = @($Catalog.functions.name | Sort-Object)
    if (@(Compare-Object $ExpectedExports $ActualExports).Count -ne 0) {
        throw "Fresh DLL exports differ from the header-derived API catalogue."
    }
    $ExportOutput | Set-Content -LiteralPath (Join-Path $BuildDirectory "exports.txt") -Encoding ASCII

    $Assets = @([ordered]@{ path = $LibraryName; sha256 = (Get-FileHash $LibraryPath -Algorithm SHA256).Hash.ToLowerInvariant(); sizeBytes = (Get-Item $LibraryPath).Length })
    $SwetestArguments = $null
    if (-not $SkipSwetest) {
        $SwetestPath = Join-Path $OutputRoot "swetest64.exe"
        # Compile the same pinned sources into a standalone reference executable.
        # MAKE_DLL is intentionally absent in this second compilation.
        $SwetestArguments = $Common + $Sources + @((Join-Path $SourceDirectory "swetest.c"), "/link", "/MACHINE:X64", "/OUT:$SwetestPath")
        Invoke-CheckedNative -Program $Compiler -Arguments $SwetestArguments
        Assert-Amd64Pe -Path $SwetestPath
        $Assets += [ordered]@{ path = "swetest64.exe"; sha256 = (Get-FileHash $SwetestPath -Algorithm SHA256).Hash.ToLowerInvariant(); sizeBytes = (Get-Item $SwetestPath).Length }
    }

    $BuildEvidence = [ordered]@{
        schemaVersion = 1
        releaseTag = "v2.10.3bfinal"
        sourceCommit = $Commit
        sourceManifestSha256 = (Get-FileHash $SourceManifestPath -Algorithm SHA256).Hash.ToLowerInvariant()
        sourceVersion = $ExpectedRuntimeVersion
        runtimeVersionVerified = $false
        excelExecutionVerified = $false
        compilerPath = $Compiler
        compilerFileVersion = (Get-Item $Compiler).VersionInfo.FileVersion
        compilerSha256 = (Get-FileHash $Compiler -Algorithm SHA256).Hash.ToLowerInvariant()
        dllCompileArguments = $DllArguments
        swetestCompileArguments = $SwetestArguments
        loaderGlueSha256 = (Get-FileHash $LoaderPath -Algorithm SHA256).Hash.ToLowerInvariant()
        exportedNames = $ActualExports
        assets = $Assets
        completedAtUtc = [DateTime]::UtcNow.ToString("o")
        verificationScope = "Fresh pinned-source AMD64 compilation and static exports. DLL calls, numerical parity, and Excel acceptance are separate gates."
    }
    $EvidencePath = Join-Path $OutputRoot "build-provenance.json"
    $BuildEvidence | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $EvidencePath -Encoding UTF8
    Write-Host "Built $LibraryName from $Commit with $($ActualExports.Count) matching exports."
    Write-Host "Source reports $ExpectedRuntimeVersion; runtime and Excel acceptance remain unverified."
} finally {
    Pop-Location
}
