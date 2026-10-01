#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string] $PackageDirectory,
    [string] $SourceDirectory = (Split-Path -Parent $PSScriptRoot),
    [string] $OutputWorkbook = 'SWExcel.xlsm',
    [switch] $RunIntegration
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Release-ComObject($Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

function Get-Property($Value, [string] $Name, $Default = $null) {
    if ($null -eq $Value) { return $Default }
    $property = $Value.PSObject.Properties[$Name]
    if ($null -eq $property) { return $Default }
    return $property.Value
}

function Get-Colour([string] $Hex) {
    return [Convert]::ToInt32($Hex.Substring(0, 2), 16) +
        256 * [Convert]::ToInt32($Hex.Substring(2, 2), 16) +
        65536 * [Convert]::ToInt32($Hex.Substring(4, 2), 16)
}

function Write-Json([string] $Path, $Value) {
    $json = ConvertTo-Json -InputObject $Value -Depth 30
    [IO.File]::WriteAllText($Path, $json + [Environment]::NewLine, [Text.UTF8Encoding]::new($false))
}

function Refresh-PackageFiles([string] $Root, [string] $Version) {
    $relativeRoot = $Root.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    $records = @{}
    foreach ($file in Get-ChildItem -LiteralPath $Root -Recurse -File) {
        if ($file.Name -eq 'package-files.json') { continue }
        $relative = $file.FullName.Substring($relativeRoot.Length).Replace('\', '/')
        $records[$relative] = [ordered]@{
            path = $relative
            sha256 = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
            size = $file.Length
        }
    }
    # Match prepare.py's complete current-file inventory, including evidence.
    [string[]]$paths = @($records.Keys)
    [Array]::Sort($paths, [StringComparer]::Ordinal)
    $sorted = @($paths | ForEach-Object { $records[$_] })
    Write-Json (Join-Path $Root 'package-files.json') ([ordered]@{
        schemaVersion = 1; projectVersion = $Version; files = $sorted
    })
}

function Assert-FreshEngine($Manifest, $Catalog, [string] $EnginePath, [string] $PackageRoot) {
    $proofPath = Join-Path (Split-Path -Parent $EnginePath) 'build-provenance.json'
    if (-not (Test-Path -LiteralPath $proofPath -PathType Leaf)) {
        throw "Fresh source-build evidence is missing: $proofPath. Run tools/build-engine.ps1 -OutputRoot .\runtime\engine first."
    }
    $proof = Get-Content -LiteralPath $proofPath -Raw | ConvertFrom-Json
    if ([string]$proof.sourceCommit -cne [string]$Manifest.engine.commit) {
        throw 'Engine build source commit differs from the pinned package manifest.'
    }
    if ([string]$proof.sourceVersion -cne [string]$Manifest.engine.runtimeVersion) {
        throw 'Engine source version differs from the expected package runtime version.'
    }
    $sourceManifestPath = Join-Path $PackageRoot 'vendor/swisseph/source/provenance.json'
    $sourceManifestHash = (Get-FileHash -LiteralPath $sourceManifestPath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ([string]$proof.sourceManifestSha256 -cne $sourceManifestHash) {
        throw 'Engine build evidence refers to a different pinned-source provenance manifest.'
    }
    $name = [IO.Path]::GetFileName($EnginePath)
    $assets = @($proof.assets | Where-Object { [string]$_.path -ceq $name })
    if ($assets.Count -ne 1) { throw "Engine build evidence must contain exactly one DLL record for $name." }
    $hash = (Get-FileHash -LiteralPath $EnginePath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($hash -cne [string]$assets[0].sha256 -or (Get-Item -LiteralPath $EnginePath).Length -ne [long]$assets[0].sizeBytes) {
        throw 'Engine DLL hash/size differs from the fresh source-build evidence.'
    }
    [string[]]$expected = @($Catalog.functions | ForEach-Object { [string]$_.name })
    [string[]]$actual = @($proof.exportedNames)
    [Array]::Sort($expected, [StringComparer]::Ordinal)
    [Array]::Sort($actual, [StringComparer]::Ordinal)
    if ($expected.Count -ne 106 -or $actual.Count -ne $expected.Count -or ($actual -join "`n") -cne ($expected -join "`n")) {
        throw 'Fresh engine build exports differ from the 106-symbol pinned native catalog.'
    }
    return [ordered]@{
        evidencePath = $proofPath; evidenceSha256 = (Get-FileHash -LiteralPath $proofPath -Algorithm SHA256).Hash.ToLowerInvariant()
        sourceCommit = [string]$proof.sourceCommit; sourceVersion = [string]$proof.sourceVersion
        sourceManifestSha256 = $sourceManifestHash
        dllSha256 = $hash; exportCount = $actual.Count
        verificationScope = 'Source-build evidence and DLL hash matched; runtime and Excel execution are separate checks.'
    }
}

function Normalize-Vba([string] $Text) {
    $lines = $Text.Replace("`r`n", "`n").Replace("`r", "`n").Split("`n")
    $normalized = foreach ($line in $lines) {
        if ($line -notmatch '^\s*Attribute\s+VB_') { $line.TrimEnd() }
    }
    return (($normalized -join "`n").Trim() + "`n")
}

function Expand-SeedValue($Value) {
    if ($Value -isnot [string]) { return $Value }
    foreach ($key in $tokens.Keys) { $Value = $Value.Replace($key, [string]$tokens[$key]) }
    return $Value
}

function Set-Cell($Sheet, [string] $Address, $Value) {
    $cell = $null
    try {
        $cell = $Sheet.Range($Address)
        if ($null -eq $Value) { return }
        if ($Value -is [string]) {
            $cell.NumberFormat = '@'
            $cell.Value2 = [string](Expand-SeedValue $Value)
        } elseif ($Value -is [bool]) {
            $cell.Value2 = [bool]$Value
        } else {
            $cell.Value2 = [double]$Value
        }
    } finally { Release-ComObject $cell }
}

function Column-Name([int] $Number) {
    $name = ''
    while ($Number -gt 0) {
        $Number--
        $name = [char](65 + ($Number % 26)) + $name
        $Number = [math]::Floor($Number / 26)
    }
    return $name
}

function Catalog-Rows($Catalog) {
    $functions = @(Get-Property $Catalog 'functions' @())
    if ($functions.Count -ne 106) { throw "Expected 106 pinned native functions; found $($functions.Count)." }
    foreach ($entry in $functions) {
        $wrapper = 'Pending'
        $status = [string](Get-Property $entry 'status' 'native_inventory_only')
        if (Get-Property $entry 'worksheetWrapped' $false) {
            $wrapper = [string](Get-Property $entry 'worksheetName' 'See checkpoint documentation')
            $status = 'Checkpoint wrapper; Windows verification pending'
        }
        if (Get-Property $entry 'experimental' $false) { $status += '; upstream experimental' }
        # The native catalog is not a claim that all worksheet wrappers exist.
        ,@([string]$entry.name, [string]$entry.prototype,
            [string](Get-Property $entry 'family' (Get-Property $entry 'category' 'Native API')),
            $wrapper, $status)
    }
}

function Render-Sheet($Sheet, $Spec, $Book, $Excel) {
    $ranges = [Collections.Generic.List[object]]::new()
    try {
        $Sheet.Name = [string]$Spec.name
        $lastColumn = Column-Name ([math]::Max(@($Spec.widths).Count, @($Spec.headers).Count))
        $rows = @(Get-Property $Spec 'rows' @())
        if ((Get-Property $Spec 'source' '') -eq 'catalog') { $rows = @(Catalog-Rows $catalog) }

        Set-Cell $Sheet 'A1' $Spec.title
        Set-Cell $Sheet 'A2' $Spec.description
        for ($c = 0; $c -lt @($Spec.headers).Count; $c++) {
            Set-Cell $Sheet ((Column-Name ($c + 1)) + '4') $Spec.headers[$c]
        }
        for ($r = 0; $r -lt $rows.Count; $r++) {
            $values = @($rows[$r])
            for ($c = 0; $c -lt $values.Count; $c++) {
                Set-Cell $Sheet ((Column-Name ($c + 1)) + [string]($r + 5)) $values[$c]
            }
        }
        foreach ($item in @(Get-Property $Spec 'cells' @())) { Set-Cell $Sheet $item.cell $item.value }

        $used = $Sheet.UsedRange; $ranges.Add($used)
        $used.Font.Name = [string]$seed.style.font
        $used.Font.Size = [double]$seed.style.fontSize
        $used.Font.Color = Get-Colour $seed.style.heading
        $used.Interior.Color = Get-Colour $seed.style.background
        $used.VerticalAlignment = -4160 # xlTop
        $used.WrapText = $true
        for ($c = 0; $c -lt @($Spec.widths).Count; $c++) {
            $column = $Sheet.Columns.Item($c + 1); $ranges.Add($column)
            $column.ColumnWidth = [double]$Spec.widths[$c]
        }
        [void]$used.Rows.AutoFit()
        $title = $Sheet.Range("A1:${lastColumn}1"); $ranges.Add($title)
        $title.Merge()
        $title.Interior.Color = Get-Colour $seed.style.heading
        $title.Font.Color = Get-Colour 'FFFFFF'
        $title.Font.Size = 20
        $title.Font.Bold = $true
        $title.RowHeight = 36
        $description = $Sheet.Range("A2:${lastColumn}2"); $ranges.Add($description)
        $description.Merge()
        $description.Font.Color = Get-Colour $seed.style.muted
        $description.RowHeight = 48
        $header = $Sheet.Range('A4:' + (Column-Name (@($Spec.headers).Count)) + '4'); $ranges.Add($header)
        $header.Interior.Color = Get-Colour $seed.style.accent
        $header.Font.Color = Get-Colour 'FFFFFF'
        $header.Font.Bold = $true
        $header.RowHeight = 30
        if ($Sheet.Name -eq 'Function Catalog' -or $Sheet.Name -eq 'Bodies') {
            $table = $Sheet.Range('A4:' + (Column-Name (@($Spec.headers).Count)) + [string]($rows.Count + 4)); $ranges.Add($table)
            [void]$table.AutoFilter()
        }
        if ($Sheet.Name -eq 'Setup' -or $Sheet.Name -eq 'Options') {
            $inputs = $Sheet.Range('B5:B' + [string]($rows.Count + 4)); $ranges.Add($inputs)
            $inputs.Interior.Color = Get-Colour $seed.style.input
        }
        if ($Sheet.Name -eq 'Integration') {
            foreach ($area in @('B27:B31', 'B36:G36', 'B40:D61', 'B63:G63')) {
                $input = $Sheet.Range($area); $ranges.Add($input)
                $input.Interior.Color = Get-Colour $seed.style.input
                $input.NumberFormat = 'General'
            }
            foreach ($row in @(25, 34, 38)) {
                $section = $Sheet.Range("A${row}:H${row}"); $ranges.Add($section)
                $section.Merge()
                $section.Font.Bold = $true
                $section.Interior.Color = Get-Colour $seed.style.heading
                $section.Font.Color = Get-Colour 'FFFFFF'
                $section.RowHeight = 26
            }
            $explanation = $Sheet.Range('A62:H62'); $ranges.Add($explanation)
            $explanation.Merge()
            $explanation.RowHeight = 60
        }
        foreach ($item in @(Get-Property $Spec 'formulas' @())) {
            if ([string]$item.formula -notmatch '^=SW') { throw "Unexpected seed formula: $($item.formula)" }
            $cell = $Sheet.Range([string]$item.cell); $ranges.Add($cell)
            $cell.NumberFormat = 'General'
            $cell.Formula2 = [string]$item.formula
        }
        foreach ($item in @(Get-Property $Spec 'names' @())) {
            $escapedSheet = $Sheet.Name.Replace("'", "''")
            $address = [regex]::Match([string]$item.cell, '^([A-Z]+)([1-9][0-9]*)$')
            if (-not $address.Success) { throw "Named checkpoint cell must be one A1 address: $($item.cell)" }
            $absolute = '$' + $address.Groups[1].Value + '$' + $address.Groups[2].Value
            [void]$Book.Names.Add([string]$item.name, "='$escapedSheet'!" + $absolute)
        }
        $Sheet.Tab.Color = Get-Colour $seed.style.accent
        [void]$Sheet.Activate()
        $window = $Excel.ActiveWindow; $ranges.Add($window)
        $window.DisplayGridlines = $false
        $window.SplitRow = 4
        $window.SplitColumn = 0
        $window.FreezePanes = $true
        $window.Zoom = 90
    } finally {
        for ($i = $ranges.Count - 1; $i -ge 0; $i--) { Release-ComObject $ranges[$i] }
    }
}

if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'Build on Windows with desktop Excel; no workbook is generated on macOS.' }
if (-not [Environment]::UserInteractive -or [Diagnostics.Process]::GetCurrentProcess().SessionId -eq 0) {
    throw 'Use a logged-on interactive Windows desktop session, not a service or SYSTEM scheduled task.'
}
if ([IO.Path]::GetFileName($OutputWorkbook) -ne $OutputWorkbook -or [IO.Path]::GetExtension($OutputWorkbook) -ne '.xlsm') {
    throw '-OutputWorkbook must be an .xlsm filename directly inside PackageDirectory.'
}
$sourceRoot = (Resolve-Path -LiteralPath $SourceDirectory).Path
$packageRoot = (Resolve-Path -LiteralPath $PackageDirectory).Path
$workbookPath = Join-Path $packageRoot $OutputWorkbook
if (Test-Path -LiteralPath $workbookPath) { throw "Refusing to overwrite $workbookPath. Choose a new output filename or preserve the existing generated package first." }
$version = [IO.File]::ReadAllText((Join-Path $sourceRoot 'VERSION')).Trim()
$seed = Get-Content -LiteralPath (Join-Path $sourceRoot 'workbook/seed.json') -Raw | ConvertFrom-Json
$catalog = Get-Content -LiteralPath (Join-Path $sourceRoot 'api/catalog.json') -Raw | ConvertFrom-Json
$manifest = Get-Content -LiteralPath (Join-Path $packageRoot 'package-manifest.json') -Raw | ConvertFrom-Json
$runtimeVersion = [string](Get-Property $manifest.engine 'runtimeVersion' '')
if ([string]::IsNullOrWhiteSpace($runtimeVersion)) { throw 'package-manifest.json engine.runtimeVersion must contain the expected upstream runtime version.' }
$engineFile = [string]$manifest.engine.file
if ([IO.Path]::IsPathRooted($engineFile) -or $engineFile -match '(^|[\\/])\.\.([\\/]|$)') { throw 'Engine file must be a package-relative path.' }
$enginePath = Join-Path $packageRoot $engineFile
if (-not (Test-Path -LiteralPath $enginePath -PathType Leaf)) { throw "Engine is missing: $enginePath. Run tools/build-engine.ps1 first." }
$engineProof = Assert-FreshEngine $manifest $catalog $enginePath $packageRoot
$tokens = @{'@VERSION@' = $version; '@ENGINE_RELEASE@' = [string]$manifest.engine.releaseTag; '@ENGINE_RUNTIME_VERSION@' = $runtimeVersion; '@ENGINE_FILE@' = $engineFile.Replace('/', '\')}
$sourceFiles = @(Get-ChildItem -LiteralPath (Join-Path $sourceRoot 'src/vba') -Filter '*.bas' -File | Sort-Object Name)
if ($sourceFiles.Count -eq 0) { throw 'No src/vba/*.bas modules found.' }
$requiredModules = @('SWRuntime.bas', 'SWNative.bas', 'SWFunctions.bas', 'SWIntegration.bas')
foreach ($required in $requiredModules) {
    if ($sourceFiles.Name -notcontains $required) { throw "Required checkpoint source missing: $required" }
}
foreach ($file in $sourceFiles) {
    if (@([IO.File]::ReadAllBytes($file.FullName) | Where-Object { $_ -gt 127 }).Count -gt 0) {
        throw "VBA source must be ASCII for predictable VBE import/export: $($file.Name)"
    }
}
$runId = [DateTime]::UtcNow.ToString('yyyyMMddTHHmmssfffZ') + '-' + [guid]::NewGuid().ToString('N').Substring(0, 8)
$evidenceRoot = Join-Path $packageRoot 'evidence'
$runRoot = Join-Path $evidenceRoot $runId
$exportRoot = Join-Path $runRoot 'vba-export'
[void][IO.Directory]::CreateDirectory($exportRoot)
$report = [ordered]@{
    schemaVersion = 1; generatedAtUtc = [DateTime]::UtcNow.ToString('o'); version = $version
    workbook = $OutputWorkbook; sourceDirectory = $sourceRoot; packageDirectory = $packageRoot
    engine = [ordered]@{ expectedPath = $enginePath; sha256 = $engineProof.dllSha256; expectedRuntimeVersion = $runtimeVersion; sourceBuild = $engineProof }
    buildStatus = 'started'; sourceParityStatus = 'pending'; sourceParity = @()
    integrationStatus = 'not_run'; integrationSummary = $null; integrationChecks = @(); environment = @{}
    compilationStatus = 'not_verified'; coverage = 'First checkpoint only; full API worksheet and runtime acceptance pending.'
    evidenceDirectory = 'evidence/' + $runId
}
$excel = $null; $book = $null; $sheets = $null; $components = $null; $saved = $false; $failure = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $report.environment = [ordered]@{ applicationVersion = [string]$excel.Version; applicationBuild = [string]$excel.Build; operatingSystem = [string]$excel.OperatingSystem; powershellProcessBits = [IntPtr]::Size * 8 }
    $book = $excel.Workbooks.Add()
    $excel.Calculation = -4135 # xlCalculationManual while assembling
    $excel.CalculateBeforeSave = $false
    $sheets = $book.Worksheets
    while ($sheets.Count -gt 1) {
        $extra = $sheets.Item($sheets.Count)
        try { [void]$extra.Delete() } finally { Release-ComObject $extra }
    }
    for ($i = 0; $i -lt @($seed.sheets).Count; $i++) {
        $sheet = $null
        try {
            if ($i -eq 0) { $sheet = $sheets.Item(1) } else {
                $after = $sheets.Item($sheets.Count)
                try { $sheet = $sheets.Add([Type]::Missing, $after) } finally { Release-ComObject $after }
            }
            Render-Sheet $sheet $seed.sheets[$i] $book $excel
        } finally { Release-ComObject $sheet }
    }
    # Import requires the developer's pre-existing VBOM permission. Do not edit trust settings.
    try { $components = $book.VBProject.VBComponents } catch {
        throw 'VBA import is blocked. For this developer build only, enable Trust access to the VBA project object model in Excel Trust Center, or ask the administrator. This script never changes the registry.'
    }
    foreach ($file in $sourceFiles) {
        $component = $null
        try {
            $component = $components.Import($file.FullName)
            $exportPath = Join-Path $exportRoot $file.Name
            [void]$component.Export($exportPath)
            $sourceText = Normalize-Vba ([IO.File]::ReadAllText($file.FullName, [Text.Encoding]::ASCII))
            $exportText = Normalize-Vba ([IO.File]::ReadAllText($exportPath, [Text.Encoding]::Default))
            $parity = $sourceText -ceq $exportText
            $report.sourceParity += [ordered]@{ source = 'src/vba/' + $file.Name; moduleName = [string]$component.Name; sourceSha256 = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant(); exportSha256 = (Get-FileHash -LiteralPath $exportPath -Algorithm SHA256).Hash.ToLowerInvariant(); normalizedEqual = $parity }
            if (-not $parity) { throw "Imported/exported VBA source differs: $($file.Name). Inspect $exportPath." }
        } finally { Release-ComObject $component }
    }
    $report.sourceParityStatus = 'passed'
    $welcome = $sheets.Item('Welcome')
    try { [void]$welcome.Activate(); [void]$welcome.Range('A1').Select() } finally { Release-ComObject $welcome }
    [void]$book.SaveAs($workbookPath, 52) # xlOpenXMLWorkbookMacroEnabled
    $saved = $true
    # Restoring normal calculation can evaluate workbook UDFs. This is not a recorded test run.
    $excel.Calculation = -4105 # xlCalculationAutomatic
    if ($RunIntegration) {
        $macro = "'" + $book.Name.Replace("'", "''") + "'!SW_RunIntegrationChecks"
        $summary = [string]$excel.Run($macro)
        $report.integrationSummary = $summary
        $report.compilationStatus = 'checkpoint_macro_executed; full_API_execution_not_established'
        $integration = $sheets.Item('Integration')
        try {
            for ($r = 5; $r -le 23; $r++) {
                $name = [string]$integration.Cells.Item($r, 1).Value2
                if ([string]::IsNullOrWhiteSpace($name)) { continue }
                $report.integrationChecks += [ordered]@{ name = $name; expected = [string]$integration.Cells.Item($r, 2).Value2; actual = [string]$integration.Cells.Item($r, 3).Value2; result = [string]$integration.Cells.Item($r, 4).Value2; notes = [string]$integration.Cells.Item($r, 5).Value2 }
            }
            for ($r = 5; $r -le 16; $r++) {
                $name = [string]$integration.Cells.Item($r, 7).Value2
                if (-not [string]::IsNullOrWhiteSpace($name)) { $report.environment[$name] = [string]$integration.Cells.Item($r, 8).Value2 }
            }
        } finally { Release-ComObject $integration }
        Write-Json (Join-Path $runRoot 'integration.json') ([ordered]@{ summary = $summary; checks = $report.integrationChecks; environment = $report.environment; coverage = $report.coverage })
        if ($summary -notmatch '^PASS=[1-9][0-9]*;FAIL=0$') {
            $report.integrationStatus = 'failed'
            [void]$book.Save()
            throw "Windows integration failed: $summary. Workbook and evidence are retained."
        }
        $report.integrationStatus = 'passed_first_checkpoint_only'
    }
    $welcome = $sheets.Item('Welcome')
    try { [void]$welcome.Activate(); [void]$welcome.Range('A1').Select() } finally { Release-ComObject $welcome }
    [void]$book.Save()
    $report.buildStatus = 'saved'
} catch {
    $failure = $_
    $report.buildStatus = 'failed'
    $report.error = $_.Exception.Message
    if ($RunIntegration -and $report.integrationStatus -eq 'not_run') { $report.integrationStatus = 'failed_or_not_reached' }
    if ($null -ne $book -and $saved) { try { [void]$book.Save() } catch {} }
} finally {
    Release-ComObject $components
    Release-ComObject $sheets
    if ($null -ne $book) { try { [void]$book.Close($false) } catch {}; Release-ComObject $book }
    if ($null -ne $excel) { try { [void]$excel.Quit() } catch {}; Release-ComObject $excel }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if (Test-Path -LiteralPath $workbookPath -PathType Leaf) { $report.workbookSha256 = (Get-FileHash -LiteralPath $workbookPath -Algorithm SHA256).Hash.ToLowerInvariant() }
    Write-Json (Join-Path $runRoot 'build.json') $report
    Write-Json (Join-Path $evidenceRoot 'build.json') $report
    Refresh-PackageFiles $packageRoot $version
}
if ($null -ne $failure) { throw $failure }
Write-Output "Generated $workbookPath"
Write-Output "Source parity: $($report.sourceParityStatus); integration: $($report.integrationStatus)"
Write-Output "Evidence: $(Join-Path $runRoot 'build.json')"
