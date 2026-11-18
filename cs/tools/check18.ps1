# Comprobación rápida de 1.8 sin abrir la ventana: sintaxis, XAML, textos, cuentas del IVA con tasas, comisiones.
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$src = Join-Path $here 'GESTEM.ps1'
$logp = Join-Path (Split-Path -Parent $here) 'check18.log'
$out = New-Object Text.StringBuilder
function W([string]$s) { [void]$out.AppendLine($s) }
$fails = 0
function Check([string]$name, $got, $exp) {
    $ok = ("$got" -eq "$exp")
    if (-not $ok) { $script:fails++ }
    W ("{0}  {1}: {2}{3}" -f $(if ($ok) { 'OK  ' } else { 'FAIL' }), $name, $got, $(if ($ok) { '' } else { "  (esperado $exp)" }))
}
try {
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    $tokens = $null; $errs = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($src, [ref]$tokens, [ref]$errs)
    Check 'errores de sintaxis' $errs.Count 0
    foreach ($e in $errs) { W "     línea $($e.Extent.StartLineNumber): $($e.Message)" }
    $code = [IO.File]::ReadAllText($src, [Text.Encoding]::UTF8)
    Check 'versión' ([regex]::Match($code, "AppVersion = '([^']+)'").Groups[1].Value) '1.8.0'

    # XAML: se carga y tiene todos los nombres nuevos
    $i = $code.IndexOf("`$xaml = @'"); $i = $code.IndexOf("`n", $i) + 1; $j = $code.IndexOf("`n'@", $i)
    $x = $code.Substring($i, $j - $i)
    $win = [Windows.Markup.XamlReader]::Parse($x)
    Check 'XAML carga' ($null -ne $win) 'True'
    $miss = @()
    foreach ($n in 'NavColab','PanelColab','ColabTitle','ColabHint','ColabList','ColabOnlyPend','ColabPaidBtn','ColabName','ColabEmpty','ColabGrid','ColK1','ColK6L','TasasL','TasasBox','SumMargin') {
        if (-not $win.FindName($n)) { $miss += $n }
    }
    Check 'nombres nuevos en XAML que faltan' ($miss -join ',') ''
    Check 'columnas ColabGrid' $win.FindName('ColabGrid').Columns.Count 11

    # funciones y definiciones (sin ejecutar el programa)
    foreach ($s in $ast.EndBlock.Statements) { if ($s -is [System.Management.Automation.Language.FunctionDefinitionAst]) { . ([ScriptBlock]::Create($s.Extent.Text)) } }
    $lines = $code -split "`n"
    $from = 0; $to = 0
    for ($k = 0; $k -lt $lines.Count; $k++) {
        if (-not $from -and $lines[$k].StartsWith('$script:Nacionalidades')) { $from = $k + 1 }
        if (-not $to -and $lines[$k].StartsWith('$script:GenStatusList')) { $to = $k + 1 }
    }
    foreach ($s in $ast.EndBlock.Statements) {
        if ($s -is [System.Management.Automation.Language.FunctionDefinitionAst]) { continue }
        if ($s.Extent.StartLineNumber -ge $from -and $s.Extent.StartLineNumber -le $to) { . ([ScriptBlock]::Create($s.Extent.Text)) }
    }
    $script:ES = [Globalization.CultureInfo]::GetCultureInfo('es-ES')
    $script:IvaRate = 0.21
    $script:Cfg = [ordered]@{ Language = 'es'; TasasCanje = 'B=33; C=100; D=100'; TradProfit = [double]3; Collaborators = @() }

    # textos 1.8 en los 4 idiomas
    $noKey = @()
    foreach ($lg in 'en', 'fr', 'ar') { foreach ($k in $script:Text18['es'].Keys) { if (-not $script:Text18[$lg].ContainsKey($k)) { $noKey += "$lg.$k" } } }
    foreach ($k in 'colStPaid', 'colStPend', 'navColab', 'seccolab') { foreach ($lg in 'es','en','fr','ar') { if (-not $script:Text[$lg].ContainsKey($k)) { $noKey += "$lg.$k" } } }
    Check 'textos que faltan' ($noKey -join ',') ''

    # canje: tasas e IVA (lo que pidió la gestoría)
    $inv = [Globalization.CultureInfo]::InvariantCulture
    function F2($v) { ([double]$v).ToString('0.00', $inv) }
    $a = Split-Iva 140 33;  Check 'Canje B 140 € (33 tasas)' ("{0} + {1} + {2}" -f (F2 $a.Tasas), (F2 $a.Base), (F2 $a.Iva)) '33.00 + 88.43 + 18.57'
    $a = Split-Iva 240 100; Check 'Canje C/D 240 € (100 tasas)' ("{0} + {1} + {2}" -f (F2 $a.Tasas), (F2 $a.Base), (F2 $a.Iva)) '100.00 + 115.70 + 24.30'
    $a = Split-Iva 121;     Check 'sin tasas 121 €' ("{0} + {1}" -f (F2 $a.Base), (F2 $a.Iva)) '100.00 + 21.00'
    $canje = $script:Tramites | Where-Object { $_.Id -eq 'canje' }
    Check 'tasa B (Coche)' (Get-TasasFor $canje @{ tipo = 'B (Coche)' }) 33
    Check 'tasa C' (Get-TasasFor $canje @{ tipo = 'C' }) 100
    Check 'tasa D' (Get-TasasFor $canje @{ tipo = 'D (Autobús)' }) 100
    Check 'tasa en otro trámite' (Get-TasasFor ($script:Tramites | Where-Object { $_.Id -eq 'extranjeria' }) @{ tipo = 'B' }) 0
    $script:Cfg.TasasCanje = 'B=33,5; C=100'
    Check 'tasa con decimales' (Get-TasasFor $canje @{ tipo = 'B' }) 33.5
    $script:Cfg.TasasCanje = 'B=33; C=100; D=100'

    # comisiones y margen
    Check 'comisión 10 % de 300' (Get-Commission @{ comtipo = '%'; comval = 10 } 300) 30
    Check 'comisión fija 50' (Get-Commission @{ comtipo = '€'; comval = 50 } 300) 50
    $ext = $script:Tramites | Where-Object { $_.Id -eq 'extranjeria' }
    $m = Get-Margin $ext @{ price = 300; gastos = 38.28; comtipo = '%'; comval = 10 }
    Check 'margen extranjería 300-38,28-30' ("{0} / {1}%" -f (F2 $m.Margin), ([double]$m.Pct).ToString('0.0', $inv)) '231.72 / 77.2%'
    $m = Get-Margin $canje @{ price = 140; tipo = 'B (Coche)'; gastos = $null }
    Check 'margen canje B 140' (F2 $m.Margin) '107.00'

    # campos nuevos
    foreach ($t in $script:Tramites) {
        $keys = @($t.Fields | ForEach-Object { $_.Key })
        $hasC = ($keys -contains 'colab')
        $want = ($t.Id -ne 'tasas')
        if ($hasC -ne $want) { Check "colab en $($t.Id)" $hasC $want }
        foreach ($f in $t.Fields) { if ($f.Section -eq 'colab' -and $f.Col) { Check "$($t.Id).$($f.Key) sin columna Excel" $f.Col '' } }
    }
    $tr = $script:Tramites | Where-Object { $_.Id -eq 'traduccion' }
    Check 'traducción tiene ndocs' (@($tr.Fields | Where-Object { $_.Key -eq 'ndocs' }).Count) 1

    # página Colaboradores con datos de prueba
    $script:Cache = @{}
    $mk = { param($row, $colab, $price, $pend, $ct, $cv, $cp) [pscustomobject]@{ Row = $row; Id = "X$row"; Name = "CLIENTE $row"; Status = 'Presentado'; Color = '#FFFFFF'; PriceV = [double]$price; Pending = [double]$pend; Fecha = (Get-Date); Data = @{ colab = $colab; comtipo = $ct; comval = $cv; compag = $cp } } }
    $script:Cache['extranjeria'] = @{ Rows = @((& $mk 10 'abogado ruiz' 300 100 '%' 10 $null), (& $mk 11 'ABOGADO RUIZ' 200 0 '€' 40 'Pagada'), (& $mk 12 '' 150 0 $null $null $null)) }
    $script:Cache['canje'] = @{ Rows = @((& $mk 20 'PEPE' 140 140 '€' 20 'Pendiente')) }
    $it = Get-ColabItems
    Check 'colab: trámites con colaborador' @($it).Count 3
    $ruiz = @($it | Where-Object { $_.Colab -eq 'ABOGADO RUIZ' })
    $s = 0.0; $sp = 0.0; foreach ($y in $ruiz) { $s += $y.Com; if (-not $y.ComPaid) { $sp += $y.Com } }
    Check 'colab: comisiones RUIZ / por pagar' "$s / $sp" '70 / 30'
    Check 'colab: lista de nombres' ((Get-CollaboratorNames) -join ',') 'ABOGADO RUIZ,PEPE'
} catch {
    $fails++
    W "ERROR: $($_.Exception.Message)  (línea $($_.InvocationInfo.ScriptLineNumber))"
}
W ''
W $(if ($fails) { "RESULTADO: $fails fallo(s)" } else { 'RESULTADO: todo bien' })
[IO.File]::WriteAllText($logp, $out.ToString(), (New-Object Text.UTF8Encoding($true)))
