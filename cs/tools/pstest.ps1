# Igual que «GESTEM.exe --selftest», pero con las funciones reales de GESTEM.ps1 (para comparar).
param([string]$TestDir)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$src = Join-Path (Split-Path -Parent (Split-Path -Parent $here)) 'GESTEM.ps1'    # el del repo (1.9.2+)
$lines = [IO.File]::ReadAllLines($src, [Text.Encoding]::UTF8)
$tokens = $null; $errs = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($src, [ref]$tokens, [ref]$errs)
# errores de GESTEM.ps1 que el C# trae ya arreglados: la prueba compara con PowerShell con SOLO estas líneas corregidas
#  - Fill-CmMenu (1.9.2): @(Get-List …) no deshace el ,$arr de Get-List -> el submenú «Estado» sale como UNA opción con todos
#    los estados juntos, y al pulsarla se guarda ese texto como estado del cliente
$psFixes = @{ 'Fill-CmMenu' = @(, @('foreach ($s in @(Get-List $f.List))', 'foreach ($s in (Get-List $f.List))')) }
foreach ($fd in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false)) {
    if ($fd.Parent -ne $null -and $fd.Parent.Parent -ne $null -and $fd.Parent.Parent -ne $ast) { continue }
    $code = $fd.Extent.Text
    if ($psFixes.ContainsKey($fd.Name)) {
        foreach ($fx in $psFixes[$fd.Name]) {
            if ($code.Contains($fx[0])) { $code = $code.Replace($fx[0], $fx[1]) } else { Write-Warning "pstest: arreglo de $($fd.Name) ya no aplica (¿corregido en GESTEM.ps1?): $($fx[0])" }
        }
    }
    . ([ScriptBlock]::Create($code))
}
foreach ($code in (& (Join-Path $here 'psblocks.ps1') $lines)) { . ([ScriptBlock]::Create($code)) }
$script:ES = [Globalization.CultureInfo]::GetCultureInfo('es-ES')
$script:BackupDir = Join-Path $TestDir 'copias_ps'
$script:LogPath = Join-Path $TestDir 'ps_log.txt'
$script:FbStatePath = Join-Path $TestDir 'ps_firebase.json'; $script:FbApplying = $false
$script:HandRowBase = 100000    # en GESTEM.ps1 está antes de todo: Read-Store / Write-Store lo usan
$script:CfgPath = Join-Path $TestDir 'config_ps.json'
$script:User = ''
New-Item -ItemType Directory -Force -Path $script:BackupDir | Out-Null
$script:Cfg = [ordered]@{ ExcelPath = (Join-Path $TestDir 'ps\x.xlsx'); Language = 'es'; Colleagues = @('SAFAE', 'CHAYMAE') }
$inv = [Globalization.CultureInfo]::InvariantCulture
function N($v) { return ([double]$v).ToString('0.##', $inv) }
$o = New-Object Text.StringBuilder
foreach ($t in $script:Tramites) {
    $s = Read-Store $t
    if (-not $s) { [void]$o.Append("T $($t.Id) (sin datos)`r`n"); continue }
    $d = Build-TramiteData $t $s
    $pend = 0.0; $price = 0.0; $paid = 0.0; $net = 0.0
    foreach ($r in $d.Rows) { $pend += $r.Pending; $price += $r.PriceV; $paid += $r.PaidV; $net += $r.NetV }
    [void]$o.Append("T $($t.Id) rows=$($d.Rows.Count) firstEmpty=$($d.FirstEmpty) nextId=$($d.NextId) pend=$(N $pend) price=$(N $price) paid=$(N $paid) net=$(N $net)`r`n")
    foreach ($r in $d.Rows) {
        $fe = ''; if ($r.Fecha) { $fe = $r.Fecha.ToString('yyyy-MM-dd', $inv) }
        [void]$o.Append("R $($t.Id) $($r.Row) $($r.Id)|$($r.Name)|$($r.Doc)|$($r.Tipo)|$($r.Status)|$($r.Color)|$($r.Pay)|$(N $r.Pending)|$($r.PendingText)|$($r.Days)|$($r.Loc)|$($r.Resp)|$(N $r.PriceV)|$(N $r.PaidV)|$(N $r.NetV)|$fe`r`n")
    }
    [string[]]$keys = @($d.Lists.Keys); [Array]::Sort($keys, [StringComparer]::Ordinal)
    foreach ($k in $keys) { [void]$o.Append("L $($t.Id) $k=$(@($d.Lists[$k]) -join '|')`r`n") }
    foreach ($c in $s.Clients) {
        $fv = @{}; foreach ($k in $c.F.Keys) { $fv[$k] = $c.F[$k] }
        $fv = Resolve-Record $t $fv $s.Meta
        [void]$o.Append("V $($t.Id) $($c.Row) price=$($fv['price']) paid=$($fv['paid'])`r`n")
    }
    Write-Store $t $s
}
$script:HandRowBase = 100000; $script:HandTabs = @{}
$xl = $script:Cfg.ExcelPath
$dv = Read-XlsxValidationLists $xl
[string[]]$dk = @($dv.Keys); [Array]::Sort($dk, [StringComparer]::Ordinal)
foreach ($k in $dk) {
    $sh = $dv[$k]
    [string[]]$ck = @($sh.Cols.Keys); [Array]::Sort($ck, [StringComparer]::Ordinal)
    foreach ($c in $ck) { $v = $sh.Cols[$c]; $vs = $(if ($v -is [string]) { $v } else { @($v) -join '|' }); [void]$o.Append("D $k [$($sh.Name)] $c=$vs`r`n") }
}
foreach ($t in $script:Tramites) { [void]$o.Append("U $($t.Id) $(Update-ListsFromFile $t $dv)`r`n") }
[void]$o.Append("S changed=$(Sync-HandmadeTabs -Force)`r`n")
[string[]]$ht = @($script:HandTabs.Keys); [Array]::Sort($ht, [StringComparer]::Ordinal)
foreach ($k in $ht) { [void]$o.Append("HT $k=$($script:HandTabs[$k])`r`n") }
foreach ($t in $script:Tramites) {
    $s = Read-Store $t
    if (-not $s) { [void]$o.Append("T2 $($t.Id) (sin datos)`r`n"); continue }
    $d = Build-TramiteData $t $s
    [void]$o.Append("T2 $($t.Id) rows=$($d.Rows.Count)`r`n")
    $sc = @($s.Clients | Sort-Object { $_.Row })
    for ($i = 0; $i -lt $sc.Count; $i++) {
        $r = $d.Rows[$i]
        $line = "R2 $($t.Id) $($r.Row) $($r.Id)"
        foreach ($f in $t.Fields) { $line += "|$($f.Key)=$(ConvertTo-StoreValue $f $sc[$i].F[$f.Key])" }
        [void]$o.Append("$line`r`n")
    }
    [string[]]$keys = @($d.Lists.Keys); [Array]::Sort($keys, [StringComparer]::Ordinal)
    foreach ($k in $keys) { [void]$o.Append("L2 $($t.Id) $k=$(@($d.Lists[$k]) -join '|')`r`n") }
}
function WR($r) { if (-not $r) { return 'null' }; return "$($r.Row):$($r.Id)" }
function WRL($l) { if (-not $l) { return 'null' }; return (@($l | ForEach-Object { WR $_ }) -join ',') }
$script:Cfg.Backup = $false
$script:Cfg.ExcelRepaired = $false
Repair-ExcelFile
[void]$o.Append("RX repaired=$($script:Cfg.ExcelRepaired)`r`n")
$t = $script:Tramites | Where-Object { $_.Id -eq 'canje' }; $script:T = $t
$s = Read-Store $t; $d = Build-TramiteData $t $s; $script:Tariff = $d.Tariff
$byRow = @{}; foreach ($c in $s.Clients) { if (-not $byRow.ContainsKey([int]$c.Row)) { $byRow[[int]$c.Row] = $c } }
$v1 = @{}; foreach ($k in $byRow[13].F.Keys) { $v1[$k] = $byRow[13].F[$k] }
$v1['status'] = 'Finalizado'; $v1['paystate'] = 'Parcial'; $v1['paid'] = [double]50
[void]$o.Append("W1 $(WR (Save-ClientToFile (Get-ClientOps $v1) 13 "$($v1['name'])"))`r`n")
$v2 = @{}; foreach ($f in $t.Fields) { $v2[$f.Key] = $null }
$v2['name'] = 'PRUEBA GESTEM CS'; $v2['doc'] = 'X1234567Z'; $v2['tipo'] = 'B (Coche)'; $v2['status'] = 'Documentación'
$v2['paystate'] = 'Pendiente'; $v2['fsol'] = (New-Object DateTime 2026, 11, 14); $v2['obs'] = ' espacio '
$v3 = @{}; foreach ($k in $byRow[20].F.Keys) { $v3[$k] = $byRow[20].F[$k] }
$v3['obs'] = 'nota de prueba'; $v3['price'] = [double]99
$batch = @(@{ Ops = (Get-ClientOps $v2); Row = $d.FirstEmpty; Expect = 'PRUEBA GESTEM CS'; IsNew = $true }, @{ Ops = (Get-ClientOps $v3); Row = 20; Expect = "$($v3['name'])"; IsNew = $false })
[void]$o.Append("W2 $(WRL (Save-ClientsToFile $batch))`r`n")
$allIt = @(); foreach ($c in $s.Clients) { $allIt += @{ Ops = (Get-ClientOps $c.F); Row = $c.Row; Expect = "$($c.F['name'])"; IsNew = $false } }
$w3 = Save-ClientsToFile $allIt -OnlyBroken
[void]$o.Append("W3 $(WRL $w3) fixed=$($script:SumFixed)`r`n")
$ta = $script:Tramites | Where-Object { $_.Id -eq 'autonomo' }; $script:T = $ta
$sa = Read-Store $ta; $da = Build-TramiteData $ta $sa; $script:Tariff = $da.Tariff
[void]$o.Append("W4 $(WR (Save-ClientToFile (Get-ClientOps $sa.Clients[0].F) $sa.Clients[0].Row "$($sa.Clients[0].F['name'])"))`r`n")
$script:MissingTabs = @{}; $script:MissingWarned = @{}
try {
    $script:T = $script:Tramites | Where-Object { $_.Id -eq 'canje' }
    $h = Open-Book
    $d = Read-TramiteData $h
    [void]$o.Append("C1 rows=$($d.Rows.Count) firstEmpty=$($d.FirstEmpty) nextId=$($d.NextId) def=$($d.DefaultPrice)`r`n")
    foreach ($r in $d.Rows) {
        $fe = ''; if ($r.Fecha) { $fe = $r.Fecha.ToString('yyyy-MM-dd', $inv) }
        [void]$o.Append("CR $($r.Row) $($r.Id)|$($r.Name)|$($r.Doc)|$($r.Tipo)|$($r.Status)|$($r.Pay)|$(N $r.Pending)|$($r.Days)|$($r.Resp)|$(N $r.PriceV)|$(N $r.PaidV)|$(N $r.NetV)|$fe`r`n")
    }
    [string[]]$lk2 = @($d.Lists.Keys); [Array]::Sort($lk2, [StringComparer]::Ordinal)
    foreach ($k in $lk2) { [void]$o.Append("CL $k=$(@($d.Lists[$k]) -join '|')`r`n") }
    [string[]]$tk = @($d.Tariff.Keys); [Array]::Sort($tk, [StringComparer]::Ordinal)
    foreach ($k in $tk) { [void]$o.Append("CT $k=$(N $d.Tariff[$k])`r`n") }
    Close-Book $h $false
    $s = Read-Store $script:T; $sd = Build-TramiteData $script:T $s; $script:Tariff = $sd.Tariff
    $c13 = $s.Clients | Where-Object { $_.Row -eq 14 } | Select-Object -First 1
    $v = @{}; foreach ($k in $c13.F.Keys) { $v[$k] = $c13.F[$k] }
    $v['status'] = 'Finalizado'; $v['paystate'] = 'Pagado'; $v['obs'] = 'guardado con Excel'; $v['fres'] = (New-Object DateTime 2026, 11, 13)
    $h = Open-Book -ForWrite
    $fr = Save-Client $v 14 -ExpectName "$($v['name'])" -Handle $h
    $vn = @{}; foreach ($f in $script:T.Fields) { $vn[$f.Key] = $null }
    $vn['name'] = 'NUEVO CON EXCEL'; $vn['tipo'] = 'B (Coche)'; $vn['status'] = 'Documentación'; $vn['paystate'] = 'Pendiente'
    $fn = Save-Client $vn $null -ExpectName 'NUEVO CON EXCEL' -IsNew -Handle $h
    $script:T = $script:Tramites | Where-Object { $_.Id -eq 'autonomo' }
    $sa = Read-Store $script:T; $script:Tariff = $sa.Meta.Tariff
    $va = @{}; foreach ($k in $sa.Clients[0].F.Keys) { $va[$k] = $sa.Clients[0].F[$k] }
    $fa = Save-Client $va $null -ExpectName "$($sa.Clients[0].F['name'])" -IsNew -Handle $h
    Close-Book $h $true
    [void]$o.Append("C2 $(WR $fr) $(WR $fn) $(WR $fa)`r`n")
} catch {
    [void]$o.Append("C! $($_.Exception.Message)`r`n")
    [IO.File]::WriteAllText((Join-Path $TestDir 'ps_com_error.txt'), "$($_.Exception.Message)`r`nlínea $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line)`r`n$($_.ScriptStackTrace)")
    try { if ($h) { $h.Wb.Close($false); if ($h.Mine) { $h.Xl.Quit() } } } catch { }
}
# etapa 2d: huella para Firebase (pendientes), cliente eliminado, avisos de requerimientos
$t = $script:Tramites | Where-Object { $_.Id -eq 'sis' }
$s = Read-Store $t
$s.Clients[0].F['obs'] = 'cambio 2d'
$s.Clients[1].Del = @{ By = 'PRUEBA'; At = '2026-11-19T10:00:00' }
Write-Store $t $s
$s = Read-Store $t; $d = Build-TramiteData $t $s
[void]$o.Append("FB rows=$($d.Rows.Count) firstEmpty=$($d.FirstEmpty) nextId=$($d.NextId) del=$(@($s.Clients | Where-Object { $_.Del }).Count)`r`n")
$st = Get-FbState; [string[]]$pk = @($st.Pend.Keys); [Array]::Sort($pk, [StringComparer]::Ordinal)
foreach ($k in $pk) { [void]$o.Append("FP $k t=$($st.Pend[$k].T) h=$($st.Pend[$k].H) n=$($st.Pend[$k].N)`r`n") }
$t = $script:Tramites | Where-Object { $_.Id -eq 'requerimiento' }
$s = Read-Store $t; $i = 0
foreach ($c in $s.Clients) { $c.F['xF'] = (New-Object DateTime 2026, 11, 1).AddDays(3 * $i); $c.F['xH'] = $null; $c.F['xI'] = $(if ($i -eq 2) { 'Cerrado' } else { $null }); $i++ }
$d = Build-TramiteData $t $s
foreach ($day in 1..30) { $al = Get-ReqAlerts $d.Rows (New-Object DateTime 2026, 11, $day); [void]$o.Append("RQ $day due=$($al.Due) late=$($al.Late)`r`n") }
# etapa 3b: tasas DGT, IVA, comisión y margen de la ficha
$script:IvaRate = 0.21; $script:Cfg.TasasCanje = 'B=33; C=100; D=100'
foreach ($tid in 'canje', 'traduccion', 'homologacion') {
    $t = $script:Tramites | Where-Object { $_.Id -eq $tid }; $s = Read-Store $t; $i = 0
    foreach ($c in $s.Clients) {
        $v = @{}; foreach ($k in $c.F.Keys) { $v[$k] = $c.F[$k] }
        if ($i % 3 -eq 1) { $v['comtipo'] = '%'; $v['comval'] = 10 } elseif ($i % 3 -eq 2) { $v['comtipo'] = 'EUR'; $v['comval'] = '12,5' }
        $i++
        $tx = Get-TasasFor $t $v; $mg = Get-Margin $t $v; $p = ConvertTo-CellNumber $v['price']
        $line = "MG $tid $($c.Row) tasas=$(N $tx)"
        if ($null -ne $p -and [double]$p -gt 0) { $sp = Split-Iva $p $tx; $line += " base=$(N $sp.Base) iva=$(N $sp.Iva) tx=$(N $sp.Tasas)" }
        if ($mg) { $line += " m=$(N $mg.Margin) pct=$($mg.Pct.ToString('0.###', $inv)) com=$(N $mg.Com)" } else { $line += ' m=-' }
        [void]$o.Append("$line`r`n")
    }
}
# etapa 3b: la ficha (misma XAML): cada cliente se abre para editar; valores leídos de la ficha y textos de la tarjeta
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
$xs = [Array]::IndexOf($lines, "`$xaml = @'"); $xe = $xs + 1; while ($lines[$xe] -ne "'@") { $xe++ }
$script:Win = [Windows.Markup.XamlReader]::Parse(($lines[($xs + 1)..($xe - 1)] -join "`n"))
$script:UI = @{}
foreach ($n in 'FormGrid', 'FormTitle', 'SaveBtn', 'ClearBtn', 'DupWarn', 'CardTag', 'CardId', 'CardName', 'CardSub', 'CardTipo', 'CardStatusB', 'CardStatus', 'SumPrice', 'SumIva', 'SumMargin', 'SumPaid', 'SumPend', 'SumPendL') { $script:UI[$n] = $script:Win.FindName($n) }
$script:Cfg.DefaultNationality = 'MARROQUI'; $script:Cfg.TradProfit = [double]3; $script:Cfg.Collaborators = @(); $script:Cache = $null
$script:PayColors = @{ 'Pagado' = '#A9D08E'; 'Parcial' = '#FFE699'; 'Pendiente' = '#FF9B9B' }
function Get-Card { $u = $script:UI; return "id=$($u.CardId.Text)|tag=$($u.CardTag.Text)|name=$($u.CardName.Text)|sub=$($u.CardSub.Text)|tipo=$($u.CardTipo.Text)|st=$(if ($u.CardStatusB.Visibility -eq 'Visible') { $u.CardStatus.Text })|price=$($u.SumPrice.Text)|iva=$(if ($u.SumIva.Visibility -eq 'Visible') { $u.SumIva.Text })|mg=$(if ($u.SumMargin.Visibility -eq 'Visible') { $u.SumMargin.Text })|paid=$($u.SumPaid.Text)|pend=$($u.SumPend.Text)|title=$($u.FormTitle.Text)" }
foreach ($t in $script:Tramites) {
    $s = Read-Store $t; if (-not $s) { continue }
    $script:T = $t; $d = Build-TramiteData $t $s
    $script:Lists = $d.Lists; $script:Tariff = $d.Tariff; $script:DefaultPrice = $d.DefaultPrice; $script:Rows = $d.Rows; $script:NextId = $d.NextId
    Build-Form; Reset-Form
    $nv = Get-FormValues; $line = "FN $($t.Id)"; foreach ($f in $t.Fields) { $line += "|$($f.Key)=$(ConvertTo-StoreValue $f $nv[$f.Key])" }
    [void]$o.Append("$line`r`nFC $($t.Id) new $(Get-Card)`r`n")
    foreach ($r in $d.Rows) {
        Fill-FormForEdit $r
        $fv = Get-FormValues; $line = "FE $($t.Id) $($r.Row)"; foreach ($f in $t.Fields) { $line += "|$($f.Key)=$(ConvertTo-StoreValue $f $fv[$f.Key])" }
        [void]$o.Append("$line`r`nFC $($t.Id) $($r.Row) $(Get-Card)`r`n")
    }
}
# etapa 3c: menú del clic derecho, atajos (estado, pagado), responsable, eliminar y deshacer — con la hora fija
$script:FixedNow = New-Object DateTime 2026, 11, 19, 12, 0, 0
function Get-Date { if ($args.Count) { Microsoft.PowerShell.Utility\Get-Date @args } else { $script:FixedNow } }
$script:User = 'PRUEBA'; $script:Cfg.Colleagues = @('ana', ' LUIS', 'ANA'); $script:Cfg.OpenAfterSave = $false
foreach ($n in 'ClientsGrid', 'StatusText', 'NavClients', 'NavNew') { $script:UI[$n] = $script:Win.FindName($n) }
foreach ($n in 'CmFolder', 'CmStatus', 'CmPaid', 'CmResp', 'CmCopyName', 'CmCopyDoc', 'CmCopyTel', 'CmExcel') { $script:UI[$n] = New-Object Windows.Controls.MenuItem }
# sin Excel ni ventana completa: Load-Data relee el trámite y rellena la lista; los avisos se apuntan
function Load-Data { $s = Read-Store $script:T; $d = Build-TramiteData $script:T $s; $script:Rows = $d.Rows; $script:Lists = $d.Lists; $script:Tariff = $d.Tariff; $script:DefaultPrice = $d.DefaultPrice; $script:NextId = $d.NextId; $script:UI.ClientsGrid.ItemsSource = $script:Rows; return $true }
function Show-Toast([string]$text, [string]$btnText = '', $action = $null, [int]$seconds = 7) { [void]$o.Append("TOAST $text|$btnText`r`n") }
function Show-Error([string]$msg) { [void]$o.Append("ERR $($msg -replace '\s+', ' ')`r`n") }
function Get-CmText {
    $u = $script:UI
    $st = @($u.CmStatus.Items | ForEach-Object { "$($_.Tag)" + $(if ($_.IsChecked) { '*' }) }) -join ','
    $rs = @($u.CmResp.Items | ForEach-Object { "$($_.Tag)" + $(if ($_.IsChecked) { '*' }) }) -join ','
    return "st=$st|$($u.CmStatus.IsEnabled) resp=$($u.CmResp.Header)|$rs|$($u.CmResp.IsEnabled) paid=$($u.CmPaid.IsEnabled) doc=$($u.CmCopyDoc.IsEnabled) tel=$($u.CmCopyTel.IsEnabled)"
}
function Get-TestRow([int]$skip) {
    # cliente de GESTEM sin NIE repetido (si no, guardar pregunta)
    $k = 0
    foreach ($r in $script:Rows) {
        if ([int]$r.Row -ge $script:HandRowBase) { continue }
        $dc = "$($r.Doc)".Trim(); if ($dc -and @($script:Rows | Where-Object { "$($_.Doc)".Trim() -ieq $dc }).Count -ne 1) { continue }
        if ($k -eq $skip) { return $r }; $k++
    }
    return $null
}
function Get-RowText($tag, [int]$row) { $r = @($script:Rows | Where-Object { [int]$_.Row -eq $row })[0]; if (-not $r) { return "$tag $row null" }; return "$tag $row $($r.Id)|$($r.Name)|$($r.Status)|$($r.Pay)|$(N $r.Pending)|$($r.Resp) n=$($script:Rows.Count)" }
function Select-Row($r) { $script:UI.ClientsGrid.SelectedItem = $r }
foreach ($t in $script:Tramites) {
    if (-not (Read-Store $t)) { continue }
    $script:T = $t; [void](Load-Data)
    foreach ($r in $script:Rows) { Select-Row $r; Fill-CmMenu $r; [void]$o.Append("CM $($t.Id) $($r.Row) $(Get-CmText)`r`n") }
}
$t = $script:Tramites | Where-Object { $_.Id -eq 'canje' }; $script:T = $t; [void](Load-Data); Build-Form; Reset-Form
$q = Get-TestRow 0; $qn = [int]$q.Row
$all = Get-List (@($t.Fields | Where-Object { $_.Key -eq 'status' })[0]).List; $sts = @($all | Where-Object { "$_".Trim() -and "$_" -ne "$($q.Status)" })
Select-Row $q; Quick-SetField 'status' $sts[-1]; [void]$o.Append("$(Get-RowText 'QS' $qn) undo=$($script:UndoInfo.Key)/$($script:UndoInfo.Old)`r`n")
Invoke-QuickUndo; [void]$o.Append("$(Get-RowText 'QU' $qn)`r`n")
Select-Row (@($script:Rows | Where-Object { [int]$_.Row -eq $qn })[0]); Quick-SetField 'paystate' 'Pagado'; [void]$o.Append("$(Get-RowText 'QS' $qn)`r`n")
Invoke-QuickUndo; [void]$o.Append("$(Get-RowText 'QU' $qn)`r`n")
Invoke-QuickUndo
Select-Row (@($script:Rows | Where-Object { [int]$_.Row -eq $qn })[0]); Set-RespBulk 'ANA'; [void]$o.Append("$(Get-RowText 'RS' $qn)`r`n")
Invoke-QuickUndo; [void]$o.Append("$(Get-RowText 'RU' $qn)`r`n")
$q2 = Get-TestRow 1; $q2n = [int]$q2.Row
$u = @{ Kind = 'delete'; T = $t; Hand = $false; Name = "$($q2.Name)"; Id = "$($q2.Id)"; Uid = $null }
$u.Uid = (Read-Store $t).Clients | Where-Object { [int]$_.Row -eq $q2n } | Select-Object -First 1 | ForEach-Object { $_.Uid }
Set-ClientDeleted $t $u $true; [void](Load-Data); $script:UndoInfo = $u; Show-Toast ((L 'delDone') -f $u.Name) (L 'cmUndo') $null 15
[void]$o.Append("$(Get-RowText 'DL' $q2n)`r`n")
Invoke-QuickUndo; [void]$o.Append("$(Get-RowText 'DU' $q2n)`r`n")
# pestaña hecha a mano (tasas): responsable (pasa a GESTEM), eliminar y deshacer, eliminar otro y dejarlo
$t = $script:Tramites | Where-Object { $_.Id -eq 'tasas' }; $script:T = $t; [void](Load-Data); Build-Form; Reset-Form
$h = @($script:Rows | Where-Object { [int]$_.Row -ge $script:HandRowBase })
Select-Row $h[0]; Set-RespBulk 'LUIS'; $nid = @($script:UndoInfo.Items)[0].Id
$nr = @($script:Rows | Where-Object { "$($_.Id)" -eq $nid })[0]; [void]$o.Append("$(Get-RowText 'RH' ([int]$nr.Row))`r`n")
Invoke-QuickUndo; [void]$o.Append("$(Get-RowText 'RHU' ([int]$nr.Row))`r`n")
foreach ($i in 1, 2) {
    $h = @($script:Rows | Where-Object { [int]$_.Row -ge $script:HandRowBase }); $hr = $h[$i]
    $u = @{ Kind = 'delete'; T = $t; Hand = $true; Name = "$($hr.Name)"; Id = "$($hr.Id)"; Uid = $null }
    Set-ClientDeleted $t $u $true; [void](Load-Data); $script:UndoInfo = $u; Show-Toast ((L 'delDone') -f $u.Name) (L 'cmUndo') $null 15
    [void]$o.Append("DH $i n=$($script:Rows.Count)`r`n")
    if ($i -eq 1) { Invoke-QuickUndo; [void]$o.Append("DHU n=$($script:Rows.Count)`r`n") }
}
Remove-Item function:Get-Date
[IO.File]::WriteAllText((Join-Path $TestDir 'ps_summary.txt'), $o.ToString(), (New-Object Text.UTF8Encoding $false))
