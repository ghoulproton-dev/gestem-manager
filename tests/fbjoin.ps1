# Primera unión a Firebase con tres copias: mismo nº y otra persona, misma persona, solo en un PC y carrera entre dos PC. Clientes ficticios (de canje.json solo se usa el formato); al final se borra todo de Firestore.
# Uso: powershell -NoProfile -ExecutionPolicy Bypass -STA -File tests\fbjoin.ps1 [-Store <carpeta «GESTEM datos»; por defecto la del Excel de la config>]
# No contiene claves: la clave de GESTEM se lee de la config de este PC (protegida con DPAPI).
param([string]$Store = '')
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$here = Join-Path $env:TEMP 'gestem-tests'; New-Item -ItemType Directory -Force $here | Out-Null
$fail = 0
function Check($ok, $msg) { if ($ok) { "OK   $msg" } else { "FAIL $msg"; $script:fail++ } }
Add-Type -AssemblyName System.Security
$errs = $null; $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'GESTEM.ps1'), [ref]$null, [ref]$errs)
Check ($errs.Count -eq 0) "sintaxis ($($errs.Count) errores)"
foreach ($s in $ast.EndBlock.Statements) {
    if ($s -is [System.Management.Automation.Language.FunctionDefinitionAst]) { . ([ScriptBlock]::Create($s.Extent.Text)) }
    elseif ($s.Extent.StartLineNumber -lt 1150 -or $s.Extent.Text -match '^\$script:(Fb\w+|HandRowBase)\s*=') { try { . ([ScriptBlock]::Create($s.Extent.Text)) } catch { } }
}
$script:Logs = New-Object System.Collections.Generic.List[string]
function Write-Log([string]$m) { $script:Logs.Add("[$env:COMPUTERNAME] $m") }
function Show-Toast([string]$text) { $script:Logs.Add("[$env:COMPUTERNAME] AVISO: $text") }
$script:Cfg = Get-Content -Raw (Join-Path $env:APPDATA 'GESTEM\config.json') -Encoding UTF8 | ConvertFrom-Json
$script:Cfg | Add-Member -NotePropertyName FbSync -NotePropertyValue $true -Force    # PCs de prueba: sincronizan aunque este PC tenga la sincronización apagada (solo en memoria)
if (-not $Store) { $Store = Get-StoreDir }    # la del Excel de este PC
$script:Cfg.Language = 'es'
$realPc = $env:COMPUTERNAME
$t = $script:Tramites | Where-Object { $_.Id -eq 'canje' }
$base = [IO.File]::ReadAllText((Join-Path $Store 'canje.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json

function New-Copy($pc, $clients) {
    $d = Join-Path $here "join\$pc"; Remove-Item -Recurse -Force $d -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Force (Join-Path $d 'GESTEM datos') | Out-Null
    $j = $base.PSObject.Copy()
    $j.clients = @(foreach ($c in $clients) {
        $f = [ordered]@{}; foreach ($x in $t.Fields) { $f[$x.Key] = $null }
        $f.name = $c.name; $f.doc = $c.doc; $f.price = $c.price
        [ordered]@{ row = 10 + $c.n; id = 'CANJ-2026-{0:000}' -f $c.n; sync = $true; nuevo = $false; updated = $c.upd; by = $pc; f = $f }
    })
    $j.adoptados = @()
    [IO.File]::WriteAllText((Join-Path $d 'GESTEM datos\canje.json'), ($j | ConvertTo-Json -Depth 20), (New-Object Text.UTF8Encoding($true)))
}
function Use-Pc($pc) {
    if ($script:FbJob) { while (-not $script:FbJob.H.IsCompleted) { Start-Sleep -Milliseconds 300 }; $script:FbJob.PS.Dispose(); $script:FbJob = $null }    # se termina la vuelta del PC anterior
    $env:COMPUTERNAME = $pc; $script:FbLastErr = ''
    $script:Cfg.ExcelPath = Join-Path $here "join\$pc\x.xlsx"; $script:FbStatePath = Join-Path $here "join\$pc\firebase.json"
}
function Tick { Step-FbSync; $w = 0; while ($script:FbJob -and -not $script:FbJob.H.IsCompleted -and $w -lt 90) { Start-Sleep 1; $w++ } }
# Firestore de la gestoría: la prueba solo corre si está vacío (si no, puede haber datos reales) y lo deja vacío
function Get-FbAllDocs { $r = Invoke-Fs Post ':runQuery' @{ structuredQuery = @{ from = @(@{ allDescendants = $true }) } }; return @(foreach ($x in @($r)) { if ($x.document) { $x.document } }) }
function Assert-FbEmpty { $n = @(Get-FbAllDocs).Count; if ($n) { throw "Firestore tiene $n documento(s): pueden ser datos reales, la prueba no sigue" } }
function Clear-FbTest {
    # primero se espera a la vuelta en segundo plano: si no, podría subir algo después de limpiar
    if ($script:FbJob) { while (-not $script:FbJob.H.IsCompleted) { Start-Sleep -Milliseconds 300 }; $script:FbJob.PS.Dispose(); $script:FbJob = $null }
    $a = Connect-Firebase
    foreach ($d in @(Get-FbAllDocs)) { [void](Invoke-FbJson Delete "https://firestore.googleapis.com/v1/$($d.name)" $null @{ Authorization = "Bearer $($a.Id)" }) }
    "Firestore limpio: quedan $(@(Get-FbAllDocs).Count) documentos"
}
function Get-Cloud {
    $r = Invoke-Fs Post ':runQuery' @{ structuredQuery = @{ from = @(@{ collectionId = 'clientes' }) } }
    @(foreach ($x in @($r)) { if ($x.document) { @{ Uid = ($x.document.name -split '/')[-1]; Id = $x.document.fields.id.stringValue; Pc = $x.document.fields.pc.stringValue; O = ($x.document.fields.datos.stringValue | ConvertFrom-Json) } } })
}
function Show-Store { foreach ($c in ((Read-Store $t).Clients | Sort-Object { $_.Row })) { "     {0} fila {1} {2} {3} {4}" -f $c.Id, $c.Row, $c.F['name'], $c.F['doc'], $c.F['price'] } }

$old = '2026-11-01T10:00:00'; $new = '2026-11-20T10:00:00'
New-Copy 'PC-A' @(
    @{ n = 1; name = 'ANA RUIZ'; doc = 'X1111111A'; price = 100; upd = '2026-11-10T10:00:00' }
    @{ n = 2; name = 'BOB LEE'; doc = 'Y2222222B'; price = 100; upd = '2026-11-10T10:00:00' }
    @{ n = 3; name = 'CARL DIAZ'; doc = $null; price = 100; upd = '2026-11-10T10:00:00' })
New-Copy 'PC-B' @(
    @{ n = 1; name = 'Ána  Ruiz'; doc = 'x-1111111-a'; price = 100; upd = $old }    # misma persona (NIE), más antiguo: gana Firebase
    @{ n = 2; name = 'DIEGO MARTIN'; doc = 'Z3333333C'; price = 200; upd = $old }   # mismo nº, otra persona
    @{ n = 3; name = 'carl díaz'; doc = $null; price = 150; upd = $new }            # misma persona (nombre), más reciente: gana B
    @{ n = 4; name = 'EVA SOLER'; doc = 'E4444444E'; price = 300; upd = $old })     # solo en B
New-Copy 'PC-C' @(
    @{ n = 1; name = 'ZOE PAZ'; doc = 'W5555555W'; price = 50; upd = $new })        # se unió "a la vez" que A y cree que la nube está vacía

Assert-FbEmpty
try {

    '--- PC-A se une (nube vacía) ---'
    Use-Pc 'PC-A'
    Tick; Check ((Get-FbState).Joined -eq '' -and @(Get-Cloud).Count -eq 0) '1ª vuelta: solo recibe, no sube nada'
    Tick; Check ((Get-FbState).Joined -eq 'push' -and (Get-FbState).Pend.Count -eq 3) 'comparado: 3 por subir'
    Tick; Check ((Get-FbState).Joined -eq 'yes' -and (Get-FbState).Pend.Count -eq 0) "subidos y unido $($script:FbLastErr)"
    Check (@(Get-Cloud).Count -eq 3) "nube: $(@(Get-Cloud).Count) clientes"

    '--- PC-B se une ---'
    Use-Pc 'PC-B'
    Tick; Check (@(Get-Cloud).Count -eq 3) '1ª vuelta: solo recibe'
    Tick; Tick; Check ((Get-FbState).Joined -eq 'yes' -and (Get-FbState).Pend.Count -eq 0) "unido $($script:FbLastErr)"
    $s = Read-Store $t; Show-Store
    $by = @{}; foreach ($c in $s.Clients) { $by[$c.Id] = $c }
    Check ($by['CANJ-2026-001'].F['name'] -eq 'ANA RUIZ') '001 misma persona por NIE: gana Firebase (más reciente)'
    Check ($by['CANJ-2026-002'].F['name'] -eq 'BOB LEE') '002 es BOB (el de Firebase)'
    Check ($by['CANJ-2026-003'].F['name'] -eq 'carl díaz' -and $by['CANJ-2026-003'].F['price'] -eq 150) '003 misma persona por nombre: gana B (más reciente)'
    Check ($by['CANJ-2026-004'].F['name'] -eq 'EVA SOLER') '004 EVA (solo en B) se queda'
    Check ($by['CANJ-2026-005'].F['name'] -eq 'DIEGO MARTIN') 'DIEGO renumerado al 005, no mezclado con BOB'
    Check (@($script:Logs | Where-Object { $_ -match 'AVISO: .*CANJ-2026-002.*DIEGO MARTIN.*CANJ-2026-005' }).Count -eq 1) 'aviso del renumerado'
    $cl = Get-Cloud; $cb = @{}; foreach ($d in $cl) { $cb[$d.Id] = $d }
    Check ($cl.Count -eq 5 -and $cb['CANJ-2026-002'].O.name -eq 'BOB LEE' -and $cb['CANJ-2026-005'].O.name -eq 'DIEGO MARTIN' -and $cb['CANJ-2026-003'].O.price -eq 150 -and $cb['CANJ-2026-004'].O.name -eq 'EVA SOLER' -and $cb['CANJ-2026-001'].O.name -eq 'ANA RUIZ') 'nube: 5 clientes correctos'
    foreach ($tt in $script:Tramites) { $x = Read-Store $tt; if ($x) { Write-Store $tt $x } }
    Check ((Get-FbState).Pend.Count -eq 0) 'B guardar sin cambios tras unirse: nada por enviar'

    '--- PC-A recibe lo de B ---'
    Use-Pc 'PC-A'
    Tick; Tick; Show-Store
    $by = @{}; foreach ($c in (Read-Store $t).Clients) { $by[$c.Id] = $c }
    Check ($by.Count -eq 5 -and $by['CANJ-2026-005'].F['name'] -eq 'DIEGO MARTIN' -and $by['CANJ-2026-003'].F['price'] -eq 150 -and $by['CANJ-2026-002'].F['name'] -eq 'BOB LEE') 'A tiene los 5 iguales que la nube'

    '--- PC-C: se "unió" antes de que A subiera (carrera) ---'
    Use-Pc 'PC-C'
    $s = Read-Store $t; $z = $s.Clients[0]; Write-Store $t $s
    $st = Get-FbState; $st.Joined = 'push'; $st.Pend[$z.Uid] = @{ T = 'canje'; H = 'nuevo'; N = $true }; Save-FbState $st
    Tick; Tick
    Check ((Get-FbState).Joined -eq '' -and $script:FbLastErr) "subida rechazada (ya existe), vuelve a unirse: $($script:FbLastErr.Substring(0, [Math]::Min(60, $script:FbLastErr.Length)))"
    Check (((Get-Cloud) | Where-Object { $_.Id -eq 'CANJ-2026-001' }).O.name -eq 'ANA RUIZ') 'ANA no se ha sobrescrito con ZOE'
    Tick; Tick; Tick
    Check ((Get-FbState).Joined -eq 'yes') "C unido $($script:FbLastErr)"
    Show-Store
    $zc = (Get-Cloud) | Where-Object { $_.O.name -eq 'ZOE PAZ' }
    Check ($zc.Id -eq 'CANJ-2026-006' -and @(Get-Cloud).Count -eq 6) "ZOE renumerada a $($zc.Id), nube con $(@(Get-Cloud).Count)"
} finally {
    Use-Pc $realPc
    Clear-FbTest
}
''
$script:Logs | ForEach-Object { "log: $_" }
"fallos: $fail"
