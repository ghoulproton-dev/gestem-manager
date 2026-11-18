# Firebase paso 2: seguimiento de cambios, recibir, renumerar y contraseñas locales, sobre una copia del almacén. -Net: también una vuelta real por Firestore con un cliente ficticio (se borra).
# Uso: powershell -NoProfile -ExecutionPolicy Bypass -STA -File tests\fbtest.ps1 [-Store <carpeta «GESTEM datos»; por defecto la del Excel de la config>]
# No contiene claves: la clave de GESTEM se lee de la config de este PC (protegida con DPAPI).
param([switch]$Net, [string]$Store = '')
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
function Write-Log([string]$m) { $script:Logs.Add($m) }
function Show-Toast { }
$script:Cfg = Get-Content -Raw (Join-Path $env:APPDATA 'GESTEM\config.json') -Encoding UTF8 | ConvertFrom-Json
$script:Cfg | Add-Member -NotePropertyName FbSync -NotePropertyValue $true -Force    # PCs de prueba: sincronizan aunque este PC tenga la sincronización apagada (solo en memoria)
if (-not $Store) { $Store = Get-StoreDir }    # la del Excel de este PC
Remove-Item -Recurse -Force (Join-Path $here 't') -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force (Join-Path $here 't\GESTEM datos') | Out-Null
Copy-Item (Join-Path $Store '*.json') (Join-Path $here 't\GESTEM datos')
$script:Cfg.ExcelPath = Join-Path $here 't\x.xlsx'
$script:FbStatePath = Join-Path $here 't\firebase.json'
Remove-Item -LiteralPath $script:FbStatePath -ErrorAction SilentlyContinue

# 1) leer y volver a guardar sin cambios no deja nada por enviar
$n = 0
foreach ($t in $script:Tramites) { $s = Read-Store $t; if ($s) { Write-Store $t $s; $n += $s.Clients.Count } }
$st = Get-FbState
Check ($st.Pend.Count -eq 0) "guardar sin cambios: $($st.Pend.Count) por enviar de $n clientes"
if ($st.Pend.Count) { $st.Pend.Keys | Select-Object -First 5 | ForEach-Object { "     $_" } }

# 2) uid estable al volver a leer
$t = $script:Tramites | Where-Object { $_.Id -eq 'traduccion' }
$s = Read-Store $t; $c = $s.Clients | Where-Object { $_.Row -lt $script:HandRowBase } | Select-Object -First 1
$j = [IO.File]::ReadAllText((Get-StorePath $t)) | ConvertFrom-Json
Check (($j.clients | Where-Object { $_.row -eq $c.Row }).uid -eq $c.Uid -and $c.Uid) "uid guardado ($($c.Uid))"

# 3) cambiar un cliente lo deja por enviar, una sola vez
$c.F['name'] = "$($c.F['name']) X"; Write-Store $t $s
$st = Get-FbState
Check ($st.Pend.Count -eq 1 -and $st.Pend.ContainsKey($c.Uid)) "cliente cambiado: por enviar ($($st.Pend.Count))"
$s = Read-Store $t; Write-Store $t $s
Check ((Get-FbState).Pend.Count -eq 1) 'volver a guardar no lo duplica'

# 4) cliente recibido de otro PC se aplica sin quedar por enviar; el cambio local más reciente gana
$st.Pend.Clear(); Save-FbState $st
$o = Get-StoreFields $t $c; $o['name'] = 'RECIBIDO OTRO PC'
$d = @{ Uid = $c.Uid; T = $t.Id; Id = $c.Id; Datos = ($o | ConvertTo-Json -Compress); Updated = (Get-Date).AddMinutes(1).ToString('s'); By = 'OTRA'; Pc = 'OTROPC' }
Receive-FbDocs @($d)
$c2 = (Read-Store $t).Clients | Where-Object { $_.Uid -eq $c.Uid }
Check ($c2.F['name'] -eq 'RECIBIDO OTRO PC' -and (Get-FbState).Pend.Count -eq 0) "recibido aplicado ($($c2.F['name']))"
$s = Read-Store $t; $cc = $s.Clients | Where-Object { $_.Uid -eq $c.Uid }; $cc.F['name'] = 'MIO'; $cc.Updated = (Get-Date).AddMinutes(5).ToString('s'); Write-Store $t $s
$d.Updated = (Get-Date).AddMinutes(2).ToString('s'); $d.Datos = $d.Datos -replace 'RECIBIDO OTRO PC', 'VIEJO'
Receive-FbDocs @($d)
Check (((Read-Store $t).Clients | Where-Object { $_.Uid -eq $c.Uid }).F['name'] -eq 'MIO') 'cambio local más reciente gana'

# 5) dos clientes nuevos con el mismo nº: el de este PC se renumera
$st = Get-FbState; $st.Pend.Clear(); Save-FbState $st
$s = Read-Store $t
$used = @{}; foreach ($x in $s.Clients) { $used[[int]$x.Row] = 1 }
$row = $t.FirstRow; while ($used.ContainsKey($row)) { $row++ }
$id = Get-RowId $t $s.Meta $row
$s.Clients.Add(@{ Row = $row; Id = $id; F = @{ name = 'NUEVO ESTE PC' }; Sync = $false; New = $true; Updated = (Get-Date).ToString('s'); By = 'YO' }); Write-Store $t $s
$mine = (Read-Store $t).Clients | Where-Object { $_.Row -eq $row }
Check ((Get-FbState).Pend.ContainsKey($mine.Uid)) "cliente nuevo $id por enviar"
$o = Get-StoreFields $t @{ F = @{ name = 'NUEVO OTRO PC' } }
Receive-FbDocs @(@{ Uid = 'otro-' + [guid]::NewGuid().ToString('N'); T = $t.Id; Id = $id; Datos = ($o | ConvertTo-Json -Compress); Updated = (Get-Date).ToString('s'); By = 'OTRA'; Pc = 'OTROPC' })
$s = Read-Store $t
$a = $s.Clients | Where-Object { $_.F['name'] -eq 'NUEVO OTRO PC' }; $b = $s.Clients | Where-Object { $_.Uid -eq $mine.Uid }
Check ($a.Id -eq $id -and $a.Row -eq $row -and $b.Id -ne $id -and $b.Id) "renumerado: otro PC $($a.Id), este PC $($b.Id)"
Check ((Get-FbState).Pend[$mine.Uid].H -eq 'renumerado') 'el renumerado se vuelve a enviar'

# 6) contraseñas: solo en este PC (no se suben, no cuentan como cambio, lo recibido no las borra)
$t = $script:Tramites | Where-Object { $_.Id -eq 'certificado' }
$pk = ($t.Fields | Where-Object { $_.Local }).Key
Check ($pk -and @($t.Fields | Where-Object { $_.Local }).Count -eq 1 -and @((Get-FbFields $t) | Where-Object { $_.Key -eq $pk }).Count -eq 0) "certificado: campo contraseña local ($pk)"
Check (@($script:Tramites | Where-Object { $_.Fields | Where-Object { $_.Local } }).Count -eq 6) 'contraseña local en 6 trámites'
$st = Get-FbState; $st.Pend.Clear(); Save-FbState $st
$s = Read-Store $t; $c = $s.Clients | Where-Object { $_.Row -lt $script:HandRowBase } | Select-Object -First 1
$c.F[$pk] = 'secreto123'; Write-Store $t $s
Check ((Get-FbState).Pend.Count -eq 0) 'cambiar solo la contraseña: nada por enviar'
$o = Get-StoreFields $t $c -Fb
Check (-not $o.Contains($pk) -and ($o | ConvertTo-Json -Compress) -notmatch 'secreto123') 'lo que se sube no lleva la contraseña'
$o['name'] = 'OTRO NOMBRE'
Receive-FbDocs @(@{ Uid = $c.Uid; T = $t.Id; Id = $c.Id; Datos = ($o | ConvertTo-Json -Compress); Updated = (Get-Date).AddMinutes(1).ToString('s'); By = 'OTRA'; Pc = 'OTROPC' })
$c2 = (Read-Store $t).Clients | Where-Object { $_.Uid -eq $c.Uid }
Check ($c2.F['name'] -eq 'OTRO NOMBRE' -and $c2.F[$pk] -eq 'secreto123') 'recibido de otro PC: la contraseña de este PC se queda'

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
if ($Net) {
    Assert-FbEmpty
    # 6) vuelta completa por Firestore con un cliente ficticio (se borra al final)
    $st = Get-FbState; $st.Pend.Clear(); $st.LastMod = ''; $st.Joined = 'yes'; Save-FbState $st    # ya unido (la unión la prueba fbjoin.ps1)
    $s = Read-Store $t
    $used = @{}; foreach ($x in $s.Clients) { $used[[int]$x.Row] = 1 }
    $row = $t.FirstRow; while ($used.ContainsKey($row)) { $row++ }
    $s.Clients.Add(@{ Row = $row; Id = 'ZZTEST'; F = @{ name = 'ZZ PRUEBA GESTEM' }; Sync = $false; New = $true; Updated = (Get-Date).ToString('s'); By = 'PRUEBA' }); Write-Store $t $s
    $uid = ((Read-Store $t).Clients | Where-Object { $_.Id -eq 'ZZTEST' }).Uid
    try {
        Step-FbSync; $w = 0; while (-not $script:FbJob.H.IsCompleted -and $w -lt 60) { Start-Sleep 1; $w++ }; Step-FbSync
        Check (-not $script:FbLastErr) "sync sin error $($script:FbLastErr)"
        Check (-not (Get-FbState).Pend.ContainsKey($uid)) 'enviado y quitado de pendientes'
        $doc = Invoke-Fs Get "clientes/$uid" $null
        Check ($doc.fields.id.stringValue -eq 'ZZTEST' -and $doc.fields.pc.stringValue -eq $env:COMPUTERNAME -and $doc.fields.mod.timestampValue) "en Firestore: $($doc.fields.id.stringValue) $($doc.fields.mod.timestampValue)"
        # otro PC (sin LastMod) lo recibe
        $r = & $script:FbSyncScript @{ Key = (Get-GestemKey); ApiKey = $script:FbApiKey; Email = $script:FbEmail; Project = $script:FbProject; Gest = $script:FbGest; Auth = $null; LastMod = ''; Push = @(); Pc = 'OTROPC' }
        Check (-not $r.Err -and @($r.Pulled | Where-Object { $_.Uid -eq $uid }).Count -eq 1) "otro PC lo recibe $($r.Err)"
        $r2 = & $script:FbSyncScript @{ Key = (Get-GestemKey); ApiKey = $script:FbApiKey; Email = $script:FbEmail; Project = $script:FbProject; Gest = $script:FbGest; Auth = $r.Auth; LastMod = $r.Mod; Push = @(); Pc = 'OTROPC' }
        Check (-not $r2.Err -and @($r2.Pulled).Count -eq 0) "desde la última vez no recibe nada más ($(@($r2.Pulled).Count)) $($r2.Err)"
    } finally {
        Clear-FbTest
    }
}
''
$script:Logs | ForEach-Object { "log: $_" }
"fallos: $fail"
