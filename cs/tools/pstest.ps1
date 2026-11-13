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
[IO.File]::WriteAllText((Join-Path $TestDir 'ps_summary.txt'), $o.ToString(), (New-Object Text.UTF8Encoding $false))
