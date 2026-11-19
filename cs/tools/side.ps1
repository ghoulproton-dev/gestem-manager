# Etapa 6: GESTEM.ps1 y GESTEM.exe lado a lado, cada uno con su COPIA de los datos de prueba (la copia real del Excel).
#   side.ps1 prep            copias nuevas: PowerShell en <S>\ps (APPDATA propio en <S>\ps_appdata), C# en cs\data\prueba (modo prueba)
#   side.ps1 start           abre los dos (PowerShell desde una copia del script en <S>, para no escribir nada en el repo)
#   side.ps1 dump ps|cs F    lee por UI Automation todas las pantallas (trámites, lista, ficha del 1.º cliente, resumen,
#                            colaboradores, ajustes) y lo escribe en F
#   side.ps1 stop            cierra los dos
# Firebase apagado en los dos (FbSync = false). No toca los datos de verdad ni la config de verdad (solo la lee).
param([string]$Action, [string]$Which = '', [string]$Out = '')
$ErrorActionPreference = 'Stop'
$cs = Split-Path -Parent $PSScriptRoot
$S = Join-Path $cs 'test\side'
$realCfgPath = Join-Path $env:APPDATA 'GESTEM\config.json'

function Get-SrcDir { $c = Get-Content -Raw $realCfgPath -Encoding UTF8 | ConvertFrom-Json; return (Split-Path -Parent $c.ExcelPath) }

if ($Action -eq 'prep') {
    $src = Get-SrcDir; $cfg = Get-Content -Raw $realCfgPath -Encoding UTF8 | ConvertFrom-Json
    if ($src -notmatch 'copia|GESTEM_MANAGER_1') { throw "el Excel de la config no parece la copia de prueba: $src" }
    Remove-Item -Recurse -Force $S -ErrorAction SilentlyContinue
    foreach ($d in (Join-Path $cs 'data\prueba'), (Join-Path $cs 'data\docs'), (Join-Path $cs 'data\nube')) { Remove-Item -Recurse -Force $d -ErrorAction SilentlyContinue }
    $xl = Split-Path -Leaf $cfg.ExcelPath
    # PowerShell: Excel + «GESTEM datos» + documentos, y su propia carpeta de APPDATA
    New-Item -ItemType Directory -Force (Join-Path $S 'ps\GESTEM datos'), (Join-Path $S 'ps_appdata\GESTEM') | Out-Null
    Copy-Item -LiteralPath $cfg.ExcelPath (Join-Path $S "ps\$xl")
    Copy-Item (Join-Path $src 'GESTEM datos\*.json') (Join-Path $S 'ps\GESTEM datos')
    $docs = Join-Path $src 'GESTEM documentos'
    if (Test-Path -LiteralPath $docs) { Copy-Item -Recurse -LiteralPath $docs (Join-Path $S 'ps\GESTEM documentos') }
    $cfg.ExcelPath = Join-Path $S "ps\$xl"
    $cfg | Add-Member DocsDir (Join-Path $S 'ps\GESTEM documentos') -Force
    $cfg | Add-Member FbSync $false -Force
    if ("$($cfg.ExtBackupDir)") { $cfg.ExtBackupDir = Join-Path $S 'ps\nube\GESTEM copias' }
    ($cfg | ConvertTo-Json -Depth 10) | Set-Content (Join-Path $S 'ps_appdata\GESTEM\config.json') -Encoding UTF8
    foreach ($f in 'admin.dat') { $p = Join-Path $env:APPDATA "GESTEM\$f"; if (Test-Path $p) { Copy-Item $p (Join-Path $S "ps_appdata\GESTEM\$f") } }
    Set-Content (Join-Path $S 'ps_appdata\GESTEM\firebase_apagado.txt') 'prueba lado a lado' -Encoding UTF8
    Copy-Item (Join-Path (Split-Path -Parent $cs) 'GESTEM.ps1') (Join-Path $S 'GESTEM.ps1')
    # C#: el modo prueba copia solo la config, el Excel y «GESTEM datos» al abrir; los documentos de prueba son los mismos
    if (Test-Path -LiteralPath $docs) { Copy-Item -Recurse -LiteralPath $docs (Join-Path $cs 'data\docs') } else { New-Item -ItemType Directory -Force (Join-Path $cs 'data\docs') | Out-Null }
    "prep OK: $S"
    return
}

function Get-Proc([string]$w) {
    if ($w -eq 'ps') { return @(Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.CommandLine -like "*$S\GESTEM.ps1*" }) }
    return @(Get-CimInstance Win32_Process -Filter "Name='GESTEM.exe'" | Where-Object { $_.ExecutablePath -eq (Join-Path $cs 'GESTEM.exe') })
}

if ($Action -eq 'start') {
    $env:APPDATA = Join-Path $S 'ps_appdata'    # solo para este proceso y sus hijos (GESTEM.ps1 se vuelve a abrir oculto y lo hereda)
    Start-Process powershell.exe -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$S\GESTEM.ps1`""
    Start-Process (Join-Path $cs 'GESTEM.exe') -WorkingDirectory $cs
    $t0 = Get-Date
    while (((Get-Date) - $t0).TotalSeconds -lt 90) { if ((Get-Proc 'ps').Count -and (Get-Proc 'cs').Count) { break }; Start-Sleep 2 }
    Start-Sleep 25
    "ps: $(@(Get-Proc 'ps').ProcessId -join ',') · cs: $(@(Get-Proc 'cs').ProcessId -join ',')"
    return
}

Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$AE = [Windows.Automation.AutomationElement]; $TS = [Windows.Automation.TreeScope]; $CT = [Windows.Automation.ControlType]
function Get-Win([string]$w) {
    foreach ($p in (Get-Proc $w)) { $x = $AE::RootElement.FindFirst($TS::Children, (New-Object Windows.Automation.PropertyCondition($AE::ProcessIdProperty, [int]$p.ProcessId))); if ($x) { return $x } }
    throw "sin ventana de $w"
}

if ($Action -eq 'stop') {
    foreach ($w in 'ps', 'cs') { try { (Get-Win $w).GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern).Close() } catch { } }
    Start-Sleep 10
    "quedan: ps=$(@(Get-Proc 'ps').Count) cs=$(@(Get-Proc 'cs').Count)"
    return
}

# ---------- lectura de pantallas ----------
$win = Get-Win $Which
try { $wp = $win.GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern); if ($wp.Current.WindowVisualState -eq 'Minimized') { $wp.SetWindowVisualState('Normal'); Start-Sleep 1 } } catch { }
function ById($id, $root = $win) { $root.FindFirst($TS::Descendants, (New-Object Windows.Automation.PropertyCondition($AE::AutomationIdProperty, $id))) }
function Visible($e) { $e -and -not $e.Current.IsOffscreen }
function Texts($root) { @($root.FindAll($TS::Descendants, (New-Object Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Text))) | Where-Object { -not $_.Current.IsOffscreen -and $_.Current.Name } | ForEach-Object { $_.Current.Name }) }
function Val($e) { $p = $null; if ($e.TryGetCurrentPattern([Windows.Automation.ValuePattern]::Pattern, [ref]$p)) { return $p.Current.Value }; if ($e.TryGetCurrentPattern([Windows.Automation.TogglePattern]::Pattern, [ref]$p)) { return "$($p.Current.ToggleState)" }; if ($e.TryGetCurrentPattern([Windows.Automation.SelectionPattern]::Pattern, [ref]$p)) { return (@($p.Current.GetSelection() | ForEach-Object { $_.Current.Name }) -join ',') }; return '' }
function Inputs($root) {
    # cajas, listas y casillas visibles, en orden de pantalla, con su nombre de XAML (o su posición)
    $l = @($root.FindAll($TS::Descendants, [Windows.Automation.Condition]::TrueCondition) | Where-Object { -not $_.Current.IsOffscreen -and $_.Current.ControlType -in @($CT::Edit, $CT::ComboBox, $CT::CheckBox, $CT::RadioButton) })
    $i = 0
    foreach ($e in $l) {
        $c = $e.Current; if ($c.AutomationId -in 'PART_EditableTextBox', 'PART_TextBox') { continue }
        $id = $(if ($c.AutomationId) { $c.AutomationId } else { "#$i" }); $i++
        $v = Val $e
        if ($c.ControlType -eq $CT::ComboBox -and -not $v) { $tb = ById 'PART_EditableTextBox' $e; if ($tb) { $v = Val $tb } }
        if ($c.ControlType -eq $CT::Custom -or ($c.ClassName -eq 'DatePicker')) { $tb = ById 'PART_TextBox' $e; if ($tb) { $v = Val $tb } }
        "  {0} [{1}] {2} = {3}" -f $id, $c.ControlType.ProgrammaticName.Replace('ControlType.', ''), $c.Name, $v
    }
    foreach ($e in @($root.FindAll($TS::Descendants, (New-Object Windows.Automation.PropertyCondition($AE::ClassNameProperty, 'DatePicker'))) | Where-Object { -not $_.Current.IsOffscreen })) {
        $tb = ById 'PART_TextBox' $e; "  fecha {0} = {1} (activa={2})" -f $e.Current.AutomationId, $(if ($tb) { Val $tb }), $e.Current.IsEnabled
    }
}
function Nav($id) { (ById $id).GetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern).Select(); Start-Sleep -Milliseconds 1500 }
function Wait-Ready { $t0 = Get-Date; while (((Get-Date) - $t0).TotalSeconds -lt 60) { $st = ById 'StatusText'; if (-not $st -or $st.Current.Name -notmatch 'Cargando|Chargement|Loading|تحميل') { return }; Start-Sleep -Milliseconds 400 }; }
function Grid-Rows($gid) {
    $g = ById $gid; if (-not (Visible $g)) { return @() }
    $seen = New-Object System.Collections.Generic.List[string]; $sp = $null
    [void]$g.TryGetCurrentPattern([Windows.Automation.ScrollPattern]::Pattern, [ref]$sp)
    if ($sp -and $sp.Current.VerticallyScrollable) { $sp.SetScrollPercent(-1, 0); Start-Sleep -Milliseconds 300 }
    for ($k = 0; $k -lt 400; $k++) {
        foreach ($it in $g.FindAll($TS::Children, (New-Object Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::DataItem)))) {
            if ($it.Current.IsOffscreen) { continue }
            $t = (@($it.FindAll($TS::Descendants, (New-Object Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Text))) | ForEach-Object { $_.Current.Name }) -join ' | ')
            if (-not $seen.Contains($t)) { $seen.Add($t) }
        }
        if (-not $sp -or -not $sp.Current.VerticallyScrollable -or $sp.Current.VerticalScrollPercent -ge 99.9) { break }
        $sp.ScrollVertical([Windows.Automation.ScrollAmount]::LargeIncrement); Start-Sleep -Milliseconds 250
    }
    if ($sp -and $sp.Current.VerticallyScrollable) { $sp.SetScrollPercent(-1, 0) }
    return $seen
}
function Tramite-Items { $cb = ById 'TramiteBox'; $ec = $cb.GetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern); $ec.Expand(); Start-Sleep -Milliseconds 500
    $items = @($cb.FindAll($TS::Descendants, (New-Object Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::ListItem)))); return @{ Cb = $cb; Ec = $ec; Items = $items } }

$o = New-Object System.Collections.Generic.List[string]
Nav 'NavClients'; Wait-Ready
$ti = Tramite-Items; $n = $ti.Items.Count; $ti.Ec.Collapse()
for ($i = 0; $i -lt $n; $i++) {
    $ti = Tramite-Items; $it = $ti.Items[$i]
    $name = (@($it.FindAll($TS::Descendants, (New-Object Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Text))) | ForEach-Object { $_.Current.Name }) -join ' ')
    $it.GetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern).Select(); try { $ti.Ec.Collapse() } catch { }
    Start-Sleep -Milliseconds 800; Wait-Ready; Nav 'NavClients'; Wait-Ready
    $o.Add("=== TRÁMITE $i $name")
    $o.Add("estado: $((ById 'StatusText').Current.Name)")
    $b = ById 'ReqBadge'; $o.Add("aviso: $(if (Visible $b) { (ById 'ReqBadgeText').Current.Name } else { '(no)' })")
    $rows = @(Grid-Rows 'ClientsGrid'); $o.Add("filas: $($rows.Count)"); foreach ($r in $rows) { $o.Add("  $r") }
    # ficha del primer cliente de la lista (solo leer; se sale con «Vaciar/Cancelar» sin guardar)
    $g = ById 'ClientsGrid'; $first = $g.FindFirst($TS::Children, (New-Object Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::DataItem)))
    if ($first) {
        $first.GetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern).Select(); Start-Sleep -Milliseconds 300
        (ById 'EditBtn').GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke(); Start-Sleep -Milliseconds 1500
        $pf = $win    # los paneles (Border) no salen en UI Automation: lo visible es la página abierta
        $o.Add("ficha:"); foreach ($l in (Inputs $pf)) { $o.Add($l) }
        $o.Add("  textos: " + ((Texts $pf) -join ' | '))
        $cb = ById 'ClearBtn'; if ($cb) { $cb.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke(); Start-Sleep -Milliseconds 800 }
        Nav 'NavClients'
    }
}
Nav 'NavStats'; Wait-Ready; Start-Sleep 1
$o.Add('=== RESUMEN'); foreach ($t in (Texts $win)) { $o.Add("  $t") }
foreach ($r in (Grid-Rows 'StatsGrid')) { $o.Add("  fila: $r") }
Nav 'NavColab'; Wait-Ready; Start-Sleep 1
$o.Add('=== COLABORADORES'); foreach ($t in (Texts $win)) { $o.Add("  $t") }
foreach ($r in (Grid-Rows 'ColabGrid')) { $o.Add("  fila: $r") }
Nav 'NavSettings'; Wait-Ready; Start-Sleep 1
$o.Add('=== AJUSTES'); foreach ($l in (Inputs $win)) { $o.Add($l) }
foreach ($t in (Texts $win)) { $o.Add("  texto: $t") }
Nav 'NavClients'
[IO.File]::WriteAllLines($Out, $o, (New-Object Text.UTF8Encoding($false)))
"dump ${Which}: $($o.Count) líneas -> $Out"
