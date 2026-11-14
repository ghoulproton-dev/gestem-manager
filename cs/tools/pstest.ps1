# Igual que «GESTEM.exe --selftest», pero con las funciones reales de GESTEM.ps1 (para comparar).
param([string]$TestDir)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$src = Join-Path $here 'GESTEM.ps1'
$lines = [IO.File]::ReadAllLines($src, [Text.Encoding]::UTF8)
$tokens = $null; $errs = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($src, [ref]$tokens, [ref]$errs)
foreach ($fd in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false)) {
    if ($fd.Parent -ne $null -and $fd.Parent.Parent -ne $null -and $fd.Parent.Parent -ne $ast) { continue }
    . ([ScriptBlock]::Create($fd.Extent.Text))
}
. ([ScriptBlock]::Create(($lines[60..786] -join "`n")))
. ([ScriptBlock]::Create(($lines[887..890] -join "`n")))
. ([ScriptBlock]::Create(($lines[936])))
$script:ES = [Globalization.CultureInfo]::GetCultureInfo('es-ES')
$script:BackupDir = Join-Path $TestDir 'copias_ps'
$script:LogPath = Join-Path $TestDir 'ps_log.txt'
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
[IO.File]::WriteAllText((Join-Path $TestDir 'ps_summary.txt'), $o.ToString(), (New-Object Text.UTF8Encoding $false))
