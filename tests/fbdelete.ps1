# Eliminar cliente: se va en el otro PC al sincronizar, deshacer lo devuelve en los dos, volver a importar el Excel no lo trae, y clientes de pestañas hechas a mano. Clientes ficticios; al final se borra todo de Firestore.
# Uso: powershell -NoProfile -ExecutionPolicy Bypass -STA -File tests\fbdelete.ps1 [-Store <carpeta «GESTEM datos»; por defecto la del Excel de la config>]
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

$upd = '2026-11-10T10:00:00'
$base3 = @(
    @{ n = 1; name = 'ANA RUIZ'; doc = 'X1111111A'; price = 100; upd = $upd }
    @{ n = 2; name = 'BOB LEE'; doc = 'Y2222222B'; price = 100; upd = $upd }
    @{ n = 3; name = 'CARL DIAZ'; doc = $null; price = 100; upd = $upd })
function Add-HandRow {
    # un cliente leído de una pestaña hecha a mano (solo en este PC, sin id de Firebase)
    $s = Read-Store $t; $FV = @{}; foreach ($f in $t.Fields) { $FV[$f.Key] = $null }; $FV['name'] = 'HANA MANO'
    $s.Clients.Add(@{ Row = $script:HandRowBase + 5; Id = 'MANO-5'; F = $FV; Sync = $true; New = $false; Updated = $upd; By = 'Excel' }); Write-Store $t $s
}
function Get-Visible { @((Build-TramiteData $t (Read-Store $t)).Rows | ForEach-Object { $_.Name }) }
$canjeT = $t
New-Copy 'PC-A' $base3; New-Copy 'PC-B' $base3

Assert-FbEmpty
try {
    Use-Pc 'PC-A'; Tick; Tick; Tick
    Use-Pc 'PC-B'; Tick; Tick; Tick
    Check ((Get-FbState).Joined -eq 'yes' -and @(Get-Cloud).Count -eq 3) "A y B unidos, nube con $(@(Get-Cloud).Count)"
    # fila escrita a mano en la pestaña después de unirse: sigue siendo de la pestaña (las de antes ya se subieron al unirse)
    Use-Pc 'PC-A'; Add-HandRow; Use-Pc 'PC-B'; Add-HandRow

    '--- PC-A elimina BOB y HANA (pestaña a mano) ---'
    Use-Pc 'PC-A'
    $bob = (Read-Store $t).Clients | Where-Object { $_.F['name'] -eq 'BOB LEE' }
    Set-ClientDeleted $t @{ Hand = $false; Uid = $bob.Uid; Name = 'BOB LEE' } $true
    Set-ClientDeleted $t @{ Hand = $true; Name = 'HANA MANO' } $true
    $v = Get-Visible
    Check ($v -notcontains 'BOB LEE' -and $v -notcontains 'HANA MANO' -and $v -contains 'ANA RUIZ') "A ya no los ve ($($v -join ', '))"
    $b2 = (Read-Store $t).Clients | Where-Object { $_.Uid -eq $bob.Uid }
    Check ($b2.Del -and $b2.Del.By -and $b2.Row -eq 12) 'BOB queda como marca (quién, cuándo) y su fila sigue ocupada'
    Tick; Tick
    $cl = Get-Cloud
    $raw = Invoke-Fs Get "clientes/$($bob.Uid)" $null
    Check ($raw.fields.deleted.booleanValue -eq $true) 'Firebase: BOB con deleted=true'
    Check (@($cl | Where-Object { $_.Uid -like 'hand-*' -and $_.O.name -eq 'HANA MANO' }).Count -eq 1) 'Firebase: marca de HANA (pestaña a mano)'

    '--- PC-B lo recibe ---'
    Use-Pc 'PC-B'; Tick; Tick
    $v = Get-Visible
    Check ($v -notcontains 'BOB LEE' -and $v -notcontains 'HANA MANO' -and $v.Count -eq 2) "B ya no los ve ($($v -join ', '))"
    $s = Read-Store $t
    Check (@($s.HandDel | Where-Object { $_.On -and $_.Key -eq (ConvertTo-HandKey 'HANA MANO') }).Count -eq 1) 'B: HANA oculta también en la próxima lectura de la pestaña'

    '--- volver a importar el Excel en B no lo trae ---'
    $xl = New-Object System.Collections.Generic.List[object]
    foreach ($c in $base3) { $FV = @{}; foreach ($f in $t.Fields) { $FV[$f.Key] = $null }; $FV['name'] = $c.name; $xl.Add(@{ Row = 10 + $c.n; Id = ('CANJ-2026-{0:000}' -f $c.n); F = $FV; Sync = $true; Updated = (Get-Date).ToString('s'); By = 'Excel' }) }
    Merge-StoreDeleted $xl $s
    $bx = $xl | Where-Object { $_.F['name'] -eq 'BOB LEE' }
    Check ($bx.Del -and $bx.Updated -eq $b2.Updated) 'BOB del Excel sigue eliminado (y con su fecha, no la de hoy)'
    $xl2 = New-Object System.Collections.Generic.List[object]; foreach ($c in $xl) { if ($c.F['name'] -ne 'BOB LEE') { $xl2.Add($c) } }
    Merge-StoreDeleted $xl2 $s
    Check (@($xl2 | Where-Object { $_.Del }).Count -eq 1) 'si ya no está en el Excel, la marca se queda'

    '--- PC-A deshace ---'
    Use-Pc 'PC-A'
    Set-ClientDeleted $t @{ Hand = $false; Uid = $bob.Uid; Name = 'BOB LEE' } $false
    Set-ClientDeleted $t @{ Hand = $true; Name = 'HANA MANO' } $false
    Check ((Get-Visible) -contains 'BOB LEE') 'A vuelve a ver a BOB'
    Tick; Tick
    Use-Pc 'PC-B'; Tick; Tick
    $v = Get-Visible
    Check ($v -contains 'BOB LEE' -and $v.Count -eq 3) "B vuelve a ver a BOB ($($v -join ', '))"
    Check (@((Read-Store $t).HandDel | Where-Object { $_.On }).Count -eq 0 -and $script:HandStamp -eq 0) 'B: HANA ya no está oculta y la pestaña se vuelve a leer'
    Check ((Get-FbState).Pend.Count -eq 0) 'B no tiene nada por enviar'
} finally {
    Use-Pc $realPc
    Clear-FbTest
}
''
$script:Logs | ForEach-Object { "log: $_" }
"fallos: $fail"
