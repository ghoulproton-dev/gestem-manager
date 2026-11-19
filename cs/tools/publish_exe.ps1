# Prepara GESTEM.exe para publicarlo (instalación de GESTEM.ps1 2.0.0 y autoactualización de GESTEM.exe):
#   <Out>\app\GESTEM.exe, <Out>\app\GESTEM.exe.sha256 y <Out>\app\exe_version.txt (la versión de cs\src\App.cs).
# Por defecto <Out> es el repositorio. No hace commit ni push. Las pruebas (exeinstall.ps1) lo usan para su GitHub de mentira.
param([string]$Out, [string]$Exe, [string]$Version)
$ErrorActionPreference = 'Stop'
$cs = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
if (-not $Out) { $Out = Split-Path -Parent $cs }
if (-not $Exe) { $Exe = Join-Path $cs 'GESTEM.exe' }
if (-not $Version) { $Version = [regex]::Match([IO.File]::ReadAllText((Join-Path $cs 'src\App.cs')), 'const string Version = "([^"]+)"').Groups[1].Value }
if ($Version -notmatch '^\d+(\.\d+){1,3}$') { throw "versión no válida: $Version" }
$app = Join-Path $Out 'app'
New-Item -ItemType Directory -Force -Path $app | Out-Null
Copy-Item -LiteralPath $Exe -Destination (Join-Path $app 'GESTEM.exe') -Force
$h = (Get-FileHash -LiteralPath (Join-Path $app 'GESTEM.exe') -Algorithm SHA256).Hash
[IO.File]::WriteAllText((Join-Path $app 'GESTEM.exe.sha256'), "$h  GESTEM.exe`n")
[IO.File]::WriteAllText((Join-Path $app 'exe_version.txt'), "$Version`n")
"app\GESTEM.exe $Version $h"
