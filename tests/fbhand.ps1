# Primera sincronización: los clientes de pestañas hechas a mano se suben una vez (sin duplicarse entre PCs) y luego la pestaña no los vuelve a leer. Clientes ficticios; al final se borra todo de Firestore.
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
$base2 = @(
    @{ n = 1; name = 'ANA RUIZ'; doc = 'X1111111A'; price = 100; upd = $upd }
    @{ n = 2; name = 'BOB LEE'; doc = 'Y2222222B'; price = 100; upd = $upd })
function Add-HandRow([string]$name, [int]$k) {
    $s = Read-Store $t; $FV = @{}; foreach ($f in $t.Fields) { $FV[$f.Key] = $null }; $FV['name'] = $name
    $s.Clients.Add(@{ Row = $script:HandRowBase + $k; Id = "MANO-$k"; F = $FV; Sync = $true; New = $false; Updated = $upd; By = 'Excel' }); Write-Store $t $s
}
function Get-Visible { @((Build-TramiteData $t (Read-Store $t)).Rows | ForEach-Object { $_.Name }) }
New-Copy 'PC-A' $base2; New-Copy 'PC-B' $base2
Use-Pc 'PC-A'; Add-HandRow 'HANA MANO' 1; Add-HandRow 'IKER TAB' 2
Use-Pc 'PC-B'; Add-HandRow 'HANA MANO' 1; Add-HandRow 'JUAN SOLO' 3
Check ((Get-FbState).Pend.Count -eq 0) 'leer la pestaña a mano no deja nada por enviar (antes de unirse)'

Assert-FbEmpty
try {
    Use-Pc 'PC-A'; Tick; Tick; Tick
    $s = Read-Store $t
    $h = @($s.Clients | Where-Object { [int]$_.Row -ge $script:HandRowBase })
    Check ($h.Count -eq 0) 'A: HANA e IKER ya son clientes de GESTEM (con nº)'
    $hana = $s.Clients | Where-Object { $_.F['name'] -eq 'HANA MANO' }
    Check ($hana.Uid -like 'mano-*' -and $hana.Row -eq 13 -and $hana.Id) "A: HANA con id fijo y nº libre ($($hana.Id), $($hana.Uid))"
    Check (@($s.Adopted) -contains 'HANA MANO') 'A: la pestaña ya no vuelve a leer a HANA'
    $cl = @(Get-Cloud)
    Check ($cl.Count -eq 4) "nube: 2 normales + HANA + IKER ($($cl.Count))"

    Use-Pc 'PC-B'; Tick; Tick; Tick
    $cl = @(Get-Cloud)
    $names = @($cl | ForEach-Object { $_.O.name })
    Check ($cl.Count -eq 5 -and @($names | Where-Object { $_ -eq 'HANA MANO' }).Count -eq 1) "nube: HANA una sola vez, + JUAN ($($names -join ', '))"
    $v = Get-Visible
    Check ($v.Count -eq 5 -and @($v | Where-Object { $_ -eq 'HANA MANO' }).Count -eq 1) "B ve 5 sin duplicar ($($v -join ', '))"
    $ids = @((Read-Store $t).Clients | ForEach-Object { $_.Id })
    Check (@($ids | Select-Object -Unique).Count -eq $ids.Count) "B: nº sin repetir ($($ids -join ', '))"
    Check ((Get-FbState).Joined -eq 'yes' -and (Get-FbState).Pend.Count -eq 0) 'B unido, nada pendiente'

    Use-Pc 'PC-A'; Tick; Tick
    $v = Get-Visible
    Check ($v.Count -eq 5 -and $v -contains 'JUAN SOLO') "A recibe a JUAN ($($v -join ', '))"
    Tick
    Check ((Get-FbState).Pend.Count -eq 0) 'segunda vuelta normal: A no vuelve a subir nada'
} finally {
    Use-Pc $realPc
    Clear-FbTest
}
''
$script:Logs | ForEach-Object { "log: $_" }
"fallos: $fail"
