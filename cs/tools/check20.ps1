# Comprueba la clave de GESTEM y las descargas de GitHub, sin abrir la ventana y sin tocar tu config real
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$src = Join-Path $here 'GESTEM.ps1'
$base = Split-Path -Parent $here
$logp = Join-Path $base 'check20.log'
$out = New-Object Text.StringBuilder
function W([string]$s) { [void]$out.AppendLine($s) }
try {
    Add-Type -AssemblyName System.Security
    $tokens = $null; $errs = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($src, [ref]$tokens, [ref]$errs)
    W "errores de sintaxis: $($errs.Count)"
    foreach ($s in $ast.EndBlock.Statements) { if ($s -is [System.Management.Automation.Language.FunctionDefinitionAst]) { . ([ScriptBlock]::Create($s.Extent.Text)) } }
    $script:UpdateBase = 'https://raw.githubusercontent.com/ghoulproton-dev/gestem-manager/main'
    $script:RepoApi = 'https://api.github.com/repos/ghoulproton-dev/gestem-manager'
    $script:CfgPath = Join-Path $base 'test\config_clave.json'
    $script:LogPath = Join-Path $base 'test\log_clave.txt'
    $script:Cfg = [ordered]@{ UpdateKey = '' }
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    W ("sin clave, version.txt: " + "$(Get-RepoFile 'version.txt')".Trim())
    Set-GestemKey 'abc-prueba'
    W ("guardada cifrada: " + ($script:Cfg.UpdateKey -ne 'abc-prueba') + " · se lee: " + (Get-GestemKey))
    W ("comprobar clave falsa: " + (Test-GestemKey 'github_pat_falsa_123'))
    $script:Cfg.UpdateKey = ''; Set-GestemKey 'github_pat_falsa_123'
    try { [void](Get-RepoFile 'version.txt'); W 'clave falsa: descarga OK (MAL)' } catch { W ("clave falsa: descarga rechazada · clave borrada: " + (-not (Get-GestemKey))) }
    $real = ''; $kf = Join-Path $base 'clave.txt'
    if (Test-Path -LiteralPath $kf) { $real = (Get-Content -Raw -LiteralPath $kf).Trim() }
    if ($real) {
        W ("clave real: " + (Test-GestemKey $real))
        Set-GestemKey $real
        W ("con clave, version.txt: " + "$(Get-RepoFile 'version.txt')".Trim())
        $t = Join-Path $base 'test\GESTEM_bajado.ps1'; Get-RepoFile 'GESTEM.ps1' $t 30
        W ("con clave, GESTEM.ps1: " + (Get-Item $t).Length + " bytes")
    } else { W 'clave real: no hay clave.txt (se prueba cuando tengas la clave)' }
} catch { W "ERROR: $($_.Exception.Message) (línea $($_.InvocationInfo.ScriptLineNumber))" }
[IO.File]::WriteAllText($logp, $out.ToString(), (New-Object Text.UTF8Encoding($true)))
