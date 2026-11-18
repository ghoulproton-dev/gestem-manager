# Comprueba Find-ExcelAuto y la lista sin repetidos, sin abrir la ventana
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$src = Join-Path $here 'GESTEM.ps1'
$logp = Join-Path (Split-Path -Parent $here) 'check19.log'
$out = New-Object Text.StringBuilder
function W([string]$s) { [void]$out.AppendLine($s) }
try {
    $tokens = $null; $errs = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($src, [ref]$tokens, [ref]$errs)
    W "errores de sintaxis: $($errs.Count)"
    foreach ($s in $ast.EndBlock.Statements) { if ($s -is [System.Management.Automation.Language.FunctionDefinitionAst]) { . ([ScriptBlock]::Create($s.Extent.Text)) } }
    W ("nubes: " + (@(Get-CloudRoots) -join ' | '))
    W ("copias en: " + (Find-CompanyCloud))
    foreach ($p in @('C:\no\existe\CRM 2026CLIENTES (copia real).xlsx', '')) {
        $script:Cfg = @{ ExcelPath = $p }
        $sw = [Diagnostics.Stopwatch]::StartNew(); $r = Find-ExcelAuto; $sw.Stop()
        W ("Excel antes '{0}' -> {1}  ({2} ms)" -f $p, $r, $sw.ElapsedMilliseconds)
    }
    $script:Lists = @{ 'excel:F' = @('1 PASO', '2 PASO', '1 PASO', '', '2 PASO', '3 PASO') }
    W ("lista sin repetidos: " + ((Get-List 'excel:F') -join ' | '))
} catch { W "ERROR: $($_.Exception.Message) (línea $($_.InvocationInfo.ScriptLineNumber))" }
[IO.File]::WriteAllText($logp, $out.ToString(), (New-Object Text.UTF8Encoding($true)))
