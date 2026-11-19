# Etapa 7: paso a GESTEM.exe (Install-GestemExe / Undo-GestemExe de GESTEM.ps1) sin internet y sin tocar nada de verdad:
# GitHub de mentira (carpeta), instalación, APPDATA y accesos directos (escritorio, Inicio, barra de tareas) en cs\test\exeinst.
# Cada caso comprueba que config.json, firebase.json, admin.dat y los almacenes no cambian ni un byte. Resultado: cs\test\exeinst.txt
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$cs = Split-Path -Parent $here; $root = Split-Path -Parent $cs
$T = Join-Path $cs 'test\exeinst'
if (Test-Path $T) { Remove-Item $T -Recurse -Force }
$G = Join-Path $T 'github'; $A = Join-Path $T 'app'; $AD = Join-Path $T 'appdata'; $X = Join-Path $T 'excel\GESTEM datos'
$SC = @((Join-Path $T 'Escritorio'), (Join-Path $T 'Inicio'), (Join-Path $T 'BarraTareas'))
New-Item -ItemType Directory -Force -Path (@($G, $A, $AD, $X) + $SC) | Out-Null
Copy-Item (Join-Path $root 'GESTEM.ps1') $A; Copy-Item (Join-Path $cs 'gestem.ico') $A

# datos: copias de los de prueba (o de mentira si no hay)
$P = Join-Path $cs 'data\prueba'
foreach ($n in 'config.json', 'firebase.json', 'admin.dat') { $s = Join-Path $P $n; if (Test-Path $s) { Copy-Item $s $AD } else { [IO.File]::WriteAllText((Join-Path $AD $n), "{`"prueba`":`"$n`"}") } }
$st = @(Get-ChildItem (Join-Path $P 'excel\GESTEM datos') -Filter *.json -ErrorAction SilentlyContinue)
if ($st.Count) { $st | Copy-Item -Destination $X } else { [IO.File]::WriteAllText((Join-Path $X 'canje.json'), '{"Clients":[]}') }
function Get-DataPrint { @(foreach ($f in @(Get-ChildItem $AD -File | Where-Object { @('config.json', 'firebase.json', 'admin.dat') -contains $_.Name }) + @(Get-ChildItem $X -File)) { "$($f.Name)=$((Get-FileHash $f.FullName -Algorithm MD5).Hash)" }) -join ';' }
$data0 = Get-DataPrint

# accesos directos como los deja GESTEM.ps1 (conhost --headless) + uno ajeno que no se toca
$sh = New-Object -ComObject WScript.Shell
$ps = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'; $ch = Join-Path $env:WINDIR 'System32\conhost.exe'
foreach ($d in $SC) {
    $l = $sh.CreateShortcut((Join-Path $d 'GESTEM MANAGER.lnk'))
    $l.TargetPath = $ch; $l.Arguments = '--headless "' + $ps + '" -NoProfile -ExecutionPolicy Bypass -STA -File "' + (Join-Path $A 'GESTEM.ps1') + '" -Oculto'
    $l.WorkingDirectory = $A; $l.WindowStyle = 7; $l.IconLocation = (Join-Path $A 'gestem.ico') + ',0'; $l.Save()
}
$l = $sh.CreateShortcut((Join-Path $SC[0] 'Bloc de notas.lnk')); $l.TargetPath = Join-Path $env:WINDIR 'notepad.exe'; $l.Save()
function Get-LnkState { @(foreach ($d in $SC) { foreach ($f in @(Get-ChildItem $d -Filter *.lnk | Sort-Object Name)) { $l = $sh.CreateShortcut($f.FullName); "$(Split-Path -Leaf $d)\$($f.Name)|$($l.TargetPath)|$($l.Arguments)|$($l.WorkingDirectory)|$($l.IconLocation)|$($l.WindowStyle)" } }) -join "`n" }
$lnk0 = Get-LnkState

# las funciones de verdad del GESTEM.ps1 copiado, con las rutas de la prueba
$src = Join-Path $A 'GESTEM.ps1'; $tokens = $null; $errs = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($src, [ref]$tokens, [ref]$errs)
foreach ($fd in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false)) {
    if ($fd.Parent -ne $null -and $fd.Parent.Parent -ne $null -and $fd.Parent.Parent -ne $ast) { continue }
    . ([ScriptBlock]::Create($fd.Extent.Text))
}
$script:AppDir = $A; $script:DataDir = $AD; $script:LogPath = Join-Path $AD 'gestem.log'; $script:User = ''; $script:IsAdmin = $false; $script:IsOwnerPC = $false
# lo que no puede pasar en una prueba: internet, el registro de Windows, otros procesos, abrir programas
$script:Fetched = @()
function Get-RepoFile([string]$name, [string]$outFile, [int]$timeout = 5) {
    $script:Fetched += $name
    $p = Join-Path $G $name
    if (-not (Test-Path -LiteralPath $p)) { throw "(404) $name" }
    if ($outFile) { Copy-Item -LiteralPath $p -Destination $outFile -Force; return }
    return [IO.File]::ReadAllText($p, [Text.Encoding]::UTF8)
}
function Get-ShortcutDirs { return $SC }
$script:Net = 528040; function Get-NetRelease { return $script:Net }
$script:Others = 0; function Get-OtherGestemPs { return $script:Others }
$script:ExeOpen = $false; function Get-ExeRunning([string]$exe) { return $script:ExeOpen }
$script:Started = @(); function Start-GestemExe([string]$exe) { $script:Started += $exe }

$o = New-Object Text.StringBuilder
function Check([string]$name, [bool]$ok, [string]$why) { [void]$o.AppendLine($(if ($ok) { "OK $name" } else { "MAL $name : $why" })) }
function Control([string]$json) { [IO.File]::WriteAllText((Join-Path $G 'control.json'), $json) }
function Publish([string]$ver, [string]$exe) { [void](& (Join-Path $here 'publish_exe.ps1') -Out $G -Exe $exe -Version $ver) }
function Tries { $p = Join-Path $AD 'exe_intentos.txt'; if (Test-Path $p) { (Get-Content -Raw $p).Trim() } else { '-' } }
function Exists([string]$n) { Test-Path -LiteralPath (Join-Path $A $n) }
$good = Join-Path $cs 'GESTEM.exe'
$goodHash = (Get-FileHash $good -Algorithm SHA256).Hash
$volver = Join-Path $SC[0] 'Volver a GESTEM anterior.lnk'
# cada caso: resultado esperado, y siempre los datos iguales
function Case([string]$name, [string]$want) {
    $script:Fetched = @(); $script:Started = @()
    $r = Install-GestemExe
    Check "$name -> $want" ($r -eq $want) "salió $r; registro: $((Get-Content (Join-Path $AD 'gestem.log') -Tail 1 -ErrorAction SilentlyContinue))"
    Check "$name : datos iguales" ((Get-DataPrint) -eq $data0) "$(Get-DataPrint) vs $data0"
    Check "$name : sin GESTEM.new.exe" (-not (Exists 'GESTEM.new.exe')) 'quedó GESTEM.new.exe'
    return $r
}
function Unchanged([string]$name) {
    Check "$name : accesos directos sin tocar" ((Get-LnkState) -eq $lnk0) (Get-LnkState)
    Check "$name : no se abre GESTEM.exe" ($script:Started.Count -eq 0) "$($script:Started)"
}

Publish '2.0.0' $good
# 1) este PC no está en la lista, o «exe»: "no", o sin conexión: nada (ni cuenta como intento)
Control '{"exePcs":["OTRO-PC"]}'; [void](Case 'otro PC' 'no'); Unchanged 'otro PC'
Check 'otro PC : no descarga' ($script:Fetched -notcontains 'app/GESTEM.exe') "$($script:Fetched)"
Control '{"exePcs":"*","exe":"no"}'; [void](Case 'exe no' 'no'); Unchanged 'exe no'
Remove-Item (Join-Path $G 'control.json'); [void](Case 'sin conexión' 'sin conexión'); Unchanged 'sin conexión'
Control ('{"exePcs":["' + $env:COMPUTERNAME.ToLower() + '"]}')
$script:Others = 1; [void](Case 'otra ventana de GESTEM.ps1' 'otra ventana'); Unchanged 'otra ventana'; $script:Others = 0
Check 'sin intentos todavía' ((Tries) -eq '-') (Tries)
# 2) fallos: sin .NET 4.8, SHA-256 que no coincide, --smoke que falla; 3 intentos con la misma versión y luego espera
$script:Net = 461808; [void](Case 'sin .NET 4.8' 'fallo'); Unchanged 'sin .NET'; $script:Net = 528040
Check 'sin .NET : intento 1' ((Tries) -eq '2.0.0|1') (Tries)
[IO.File]::WriteAllText((Join-Path $G 'app\GESTEM.exe.sha256'), ('0' * 64)); [void](Case 'SHA-256 malo' 'fallo'); Unchanged 'SHA-256 malo'
Check 'SHA-256 malo : intento 2' ((Tries) -eq '2.0.0|2') (Tries)
Publish '2.0.0' (Join-Path $env:WINDIR 'System32\where.exe'); [void](Case '--smoke falla' 'fallo'); Unchanged '--smoke falla'
Check '--smoke falla : intento 3' ((Tries) -eq '2.0.0|3') (Tries)
Check 'fallos : sin GESTEM.exe' (-not (Exists 'GESTEM.exe')) 'hay GESTEM.exe'
Publish '2.0.0' $good; [void](Case '4.º arranque, misma versión' 'espera'); Unchanged 'espera'
Check 'espera : no descarga' ($script:Fetched -notcontains 'app/GESTEM.exe') "$($script:Fetched)"
# 3) versión nueva: se instala
Publish '2.0.1' $good; [void](Case 'normal (2.0.1)' 'ok')
Check 'normal : GESTEM.exe es el publicado' ((Exists 'GESTEM.exe') -and (Get-FileHash (Join-Path $A 'GESTEM.exe') -Algorithm SHA256).Hash -eq $goodHash) 'otro archivo'
Check 'normal : se abre GESTEM.exe' ($script:Started.Count -eq 1 -and $script:Started[0] -eq (Join-Path $A 'GESTEM.exe')) "$($script:Started)"
Check 'normal : sin intentos' ((Tries) -eq '-') (Tries)
$exe = Join-Path $A 'GESTEM.exe'
$moved = @(foreach ($d in $SC) { $l = $sh.CreateShortcut((Join-Path $d 'GESTEM MANAGER.lnk')); "$($l.TargetPath)|$($l.Arguments)|$($l.WorkingDirectory)|$($l.IconLocation)" })
Check 'normal : escritorio, Inicio y barra de tareas -> GESTEM.exe --real' (@($moved | Where-Object { $_ -eq "$exe|--real|$A|$exe,0" }).Count -eq 3) ($moved -join ' / ')
$np = $sh.CreateShortcut((Join-Path $SC[0] 'Bloc de notas.lnk')); Check 'normal : el acceso ajeno no se toca' ($np.TargetPath -like '*notepad.exe') $np.TargetPath
$v = $(if (Test-Path $volver) { $sh.CreateShortcut($volver) })
Check 'normal : icono «Volver a GESTEM anterior»' ($v -and $v.TargetPath -eq $ch -and $v.Arguments -like "*-File `"$(Join-Path $A 'GESTEM.ps1')`" -Oculto -Volver") "$($v.TargetPath) $($v.Arguments)"
$ins = Get-Content -Raw (Join-Path $AD 'exe_instalado.txt') -Encoding UTF8 | ConvertFrom-Json
Check 'normal : exe_instalado.txt' ($ins.version -eq '2.0.1' -and @($ins.antes).Count -eq 3) "$($ins | ConvertTo-Json -Compress)"
Check 'normal : registro' ((Get-Content (Join-Path $AD 'gestem.log') -Raw -Encoding UTF8) -match 'GESTEM.exe 2.0.1 instalado') 'falta en el registro'
# 4) volver con el icono: accesos como estaban, este PC ya no cambia solo; con GESTEM.exe abierto no hace nada
$script:ExeOpen = $true; $lnk1 = Get-LnkState
Check 'volver con GESTEM.exe abierto -> abierto' ((Undo-GestemExe) -eq 'abierto') 'no'
Check 'volver con GESTEM.exe abierto : nada cambia' ((Get-LnkState) -eq $lnk1 -and (Test-Path (Join-Path $AD 'exe_instalado.txt'))) (Get-LnkState)
$script:ExeOpen = $false
Check 'volver (icono) -> ok' ((Undo-GestemExe) -eq 'ok') 'no'
Check 'volver (icono) : accesos directos como antes' ((Get-LnkState) -eq $lnk0) (Get-LnkState)
Check 'volver (icono) : exe_desactivado.txt' ((Test-Path (Join-Path $AD 'exe_desactivado.txt')) -and -not (Test-Path (Join-Path $AD 'exe_instalado.txt'))) 'no'
Check 'volver (icono) : datos iguales' ((Get-DataPrint) -eq $data0) 'cambiaron'
[void](Case 'después del icono' 'desactivado'); Unchanged 'después del icono'
# 5) otra vez a GESTEM.exe (quitando exe_desactivado.txt) y vuelta remota (control.json «exe»: "no", la abre GESTEM.exe)
Remove-Item (Join-Path $AD 'exe_desactivado.txt')
[void](Case 'otra vez a GESTEM.exe' 'ok')
Check 'otra vez : accesos -> GESTEM.exe' ((Get-LnkState) -ne $lnk0 -and (Test-Path $volver)) (Get-LnkState)
Control '{"exePcs":"*","exe":"no"}'
Check 'volver (control.json) -> ok' ((Undo-GestemExe -Remote) -eq 'ok') 'no'
Check 'volver (control.json) : accesos directos como antes' ((Get-LnkState) -eq $lnk0) (Get-LnkState)
Check 'volver (control.json) : sin exe_desactivado.txt' (-not (Test-Path (Join-Path $AD 'exe_desactivado.txt'))) 'existe'
[void](Case 'mientras «exe»: "no"' 'no'); Unchanged 'mientras exe no'
Control '{"exePcs":"*"}'; [void](Case 'control.json lo permite otra vez' 'ok')
Check 'otra vez (remoto) : accesos -> GESTEM.exe' ((Get-LnkState) -ne $lnk0) (Get-LnkState)
Check 'GESTEM.ps1 sin errores de sintaxis' ($errs.Count -eq 0) "$($errs.Count) errores"

$out = Join-Path $cs 'test\exeinst.txt'
[IO.File]::WriteAllText($out, $o.ToString(), (New-Object Text.UTF8Encoding($true)))
$bad = @($o.ToString() -split "`r?`n" | Where-Object { $_ -like 'MAL *' }).Count
"$(@($o.ToString() -split "`r?`n" | Where-Object { $_ -like 'OK *' }).Count) OK, $bad MAL -> $out"
