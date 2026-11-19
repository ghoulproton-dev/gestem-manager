# Compara lo que dejaron pstest.ps1 (PowerShell) y GESTEM.exe --selftest (C#) en la carpeta de prueba.
# Los datos guardados se comparan por contenido (sin «guardado», la hora de guardar). Resultado en test\result.txt.
param([string]$TestDir)
$out = New-Object Text.StringBuilder
function Get-Canon([string]$p) {
    if (-not (Test-Path -LiteralPath $p)) { return '(no existe)' }
    $j = [IO.File]::ReadAllText($p, [Text.Encoding]::UTF8) | ConvertFrom-Json
    if ($j.PSObject.Properties['guardado']) { $j.PSObject.Properties.Remove('guardado') }
    # los uid nuevos son GUID al azar en los dos lados
    return ($j | ConvertTo-Json -Depth 20 -Compress) -replace '"uid":"[0-9a-f]{32}"', '"uid":"GUID"'
}
$a = [IO.File]::ReadAllLines((Join-Path $TestDir 'ps_summary.txt')); $b = @()
if (Test-Path (Join-Path $TestDir 'cs_summary.txt')) { $b = [IO.File]::ReadAllLines((Join-Path $TestDir 'cs_summary.txt')) }
$d = @(Compare-Object $a $b -SyncWindow 5)
[void]$out.AppendLine("summary: " + $(if ($d.Count) { "DISTINTO ($($d.Count) líneas)" } else { "IGUAL ($($a.Count) líneas)" }))
foreach ($x in ($d | Select-Object -First 20)) { [void]$out.AppendLine("  $($x.SideIndicator) $($x.InputObject)") }
foreach ($f in Get-ChildItem (Join-Path $TestDir 'ps\GESTEM datos') -Filter '*.json') {
    $eq = (Get-Canon $f.FullName) -eq (Get-Canon (Join-Path $TestDir ('cs\GESTEM datos\' + $f.Name)))
    [void]$out.AppendLine("$($f.Name): " + $(if ($eq) { 'IGUAL' } else { 'DISTINTO' }))
}
# pendientes: es un @{} (el orden no cuenta) y los uid nuevos son GUID al azar
function Get-FbCanon([string]$p) {
    if (-not (Test-Path -LiteralPath $p)) { return '(no existe)' }
    $j = [IO.File]::ReadAllText($p, [Text.Encoding]::UTF8) | ConvertFrom-Json
    $l = @($j.pend.PSObject.Properties | ForEach-Object { ($_.Name -replace '^[0-9a-f]{32}$', 'GUID') + "|$($_.Value.t)|$($_.Value.h)|$($_.Value.n)" } | Sort-Object)
    return "$($j.lastMod)|$($j.joined)`n" + ($l -join "`n")
}
$eq = (Get-FbCanon (Join-Path $TestDir 'ps_firebase.json')) -eq (Get-FbCanon (Join-Path $TestDir 'cs_firebase.json'))
[void]$out.AppendLine("firebase.json (pendientes): " + $(if ($eq) { 'IGUAL' } else { 'DISTINTO' }))
[IO.File]::WriteAllText((Join-Path $TestDir 'result.txt'), $out.ToString(), (New-Object Text.UTF8Encoding $true))
$out.ToString()
