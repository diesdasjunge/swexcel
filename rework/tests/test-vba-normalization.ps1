$ErrorActionPreference = 'Stop'
$builder = Join-Path (Split-Path -Parent $PSScriptRoot) 'tools/build-workbook.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($builder, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
$function = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Normalize-Vba' }, $true)
. ([scriptblock]::Create($function.Extent.Text))
$source = @'
Attribute VB_Name = "Example"
Public Function Sample(ByVal Body As Long) As String
    Sample = "Sun ""Quoted""" ' Keep Comment
    Rem Keep Remark
End Function
'@
$recased = $source.Replace('Sample', 'sample').Replace('Body', 'body')
if ((Normalize-Vba $source) -cne (Normalize-Vba $recased)) { throw 'Identifier recasing must compare equal.' }
foreach ($changed in @(
    $source.Replace('Sun', 'sun'),
    $source.Replace('Quoted', 'quoted'),
    $source.Replace('Keep Comment', 'keep comment'),
    $source.Replace('Keep Remark', 'keep remark'),
    $source.Replace('As Long', 'As Double')
)) {
    if ((Normalize-Vba $source) -ceq (Normalize-Vba $changed)) { throw 'A meaningful source change was hidden.' }
}
Write-Output 'PASS: identifier recasing accepted; literals, comments and type changes detected.'
