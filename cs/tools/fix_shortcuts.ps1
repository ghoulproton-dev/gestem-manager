# vuelve a apuntar los accesos directos de GESTEM a la copia instalada (no a la carpeta de pruebas)
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$out = Join-Path (Split-Path -Parent $here) 'test\fix_shortcuts.txt'
$log = New-Object Text.StringBuilder
$app = Join-Path $env:LOCALAPPDATA 'GESTEM'
if (-not (Test-Path -LiteralPath (Join-Path $app 'GESTEM.ps1'))) { [void]$log.AppendLine("no está instalado en $app"); [IO.File]::WriteAllText($out, $log.ToString()); return }
$ps = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
$ch = Join-Path $env:WINDIR 'System32\conhost.exe'
$me = Join-Path $app 'GESTEM.ps1'
$want = '--headless "' + $ps + '" -NoProfile -ExecutionPolicy Bypass -STA -File "' + $me + '" -Oculto'
$ico = Join-Path $app 'gestem.ico'
$sh = New-Object -ComObject WScript.Shell
foreach ($d in @([Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('Programs'))) {
    $f = Join-Path $d 'GESTEM MANAGER.lnk'
    if (-not (Test-Path -LiteralPath $f)) { [void]$log.AppendLine("no existe: $f"); continue }
    $l = $sh.CreateShortcut($f)
    [void]$log.AppendLine("antes: $f -> $($l.Arguments)")
    $l.TargetPath = $ch; $l.Arguments = $want; $l.WorkingDirectory = $app; $l.WindowStyle = 7
    if (Test-Path -LiteralPath $ico) { $l.IconLocation = "$ico,0" }
    $l.Save()
    [void]$log.AppendLine("ahora: $want")
}
[IO.File]::WriteAllText($out, $log.ToString())
