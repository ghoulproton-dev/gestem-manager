# Etapa 7 (sección 0): lo mismo que src\SelfTest7.cs con las funciones reales de GESTEM.ps1 (se carga con «.» desde pstest.ps1)
# rutas de la prueba iguales en los dos: <T>\X en vez de <prueba>\cs o <prueba>\ps
function Norm7([string]$s) { return ($s.Replace($TestDir, '<T>') -replace '<T>\\(cs|ps)', '<T>\X') }
function Start-Process { param([string]$FilePath, $ArgumentList) [void]$o.Append("OPEN $(Norm7 $FilePath)$(if ($ArgumentList) { ' ' + (Norm7 "$ArgumentList") })`r`n") }

# abrir el Excel: sin archivo -> aviso; con archivo y sin Excel abierto con él -> se abre el archivo
$script:T = $script:Tramites | Where-Object { $_.Id -eq 'canje' }
$keepX = $script:Cfg.ExcelPath
$script:Cfg.ExcelPath = Join-Path $TestDir 'no-existe.xlsx'; [void]$o.Append("XL0`r`n"); Open-ExcelWindow 5
$script:Cfg.ExcelPath = $keepX; [void]$o.Append("XL1`r`n"); Open-ExcelWindow 5; [void]$o.Append("XL2`r`n"); Open-ExcelWindow ($script:HandRowBase + 3)

# factura PDF: se dibuja en imágenes (huella de los píxeles) en vez de imprimir; datos, nº y aviso, iguales que C#
function Get-Date { if ($args.Count) { Microsoft.PowerShell.Utility\Get-Date @args } else { New-Object DateTime 2026, 11, 20, 10, 0, 0 } }
function Render7([bool]$landscape, [scriptblock]$draw) {
    $wpx = $(if ($landscape) { 1169 } else { 827 }); $hpx = $(if ($landscape) { 827 } else { 1169 })
    $hs = New-Object System.Collections.Generic.List[string]; $page = 0
    do {
        $bmp = New-Object Drawing.Bitmap($wpx, $hpx, [Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $bmp.SetResolution(100, 100)
        $g = [Drawing.Graphics]::FromImage($bmp)
        $g.Clear([Drawing.Color]::White)
        $g.PageUnit = [Drawing.GraphicsUnit]::Millimeter; $g.SmoothingMode = [Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $g.TextRenderingHint = [Drawing.Text.TextRenderingHint]::AntiAliasGridFit; $g.InterpolationMode = [Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $more = [bool](& $draw $g $page)
        $g.Dispose()
        $bd = $bmp.LockBits((New-Object Drawing.Rectangle(0, 0, $wpx, $hpx)), [Drawing.Imaging.ImageLockMode]::ReadOnly, [Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $px = New-Object byte[] ($bd.Stride * $hpx); [Runtime.InteropServices.Marshal]::Copy($bd.Scan0, $px, 0, $px.Length); $bmp.UnlockBits($bd); $bmp.Dispose()
        $md5 = [Security.Cryptography.MD5]::Create(); $hs.Add([BitConverter]::ToString($md5.ComputeHash($px)).Replace('-', '').Substring(0, 12)); $md5.Dispose()
        $page++
    } while ($more -and $page -lt 40)
    return "pages=$page px=$($hs -join ',')"
}
function Dump7($d) {
    [string[]]$ks = @($d.Keys); [Array]::Sort($ks, [StringComparer]::Ordinal)
    return (@(foreach ($k in $ks) {
        $v = $d[$k]
        $s = $(if ($k -eq 'Rows') { "n=$(@($v).Count)" } elseif ($null -eq $v) { '' } elseif ($v -is [double]) { $v.ToString('0.##', [Globalization.CultureInfo]::InvariantCulture) } elseif ($v -is [array]) { (@($v | ForEach-Object { "$($_[0])=$($_[1])" }) -join ';') } else { "$v" })
        "$k=$s"
    }) -join ' | ')
}
function Invoke-PdfPrint([string]$path, [bool]$landscape, [scriptblock]$draw) {
    [void]$o.Append("PDF $(Norm7 $path) land=$landscape $(Render7 $landscape $draw)`r`n")
    [void]$o.Append("PDFD $(Dump7 $script:PdfData)`r`n")
}
$script:PdfPrinter = 'Microsoft Print to PDF'   # en GESTEM.ps1 se pone fuera de las funciones
$c = $script:Cfg
$c.InvName = 'GESTORÍA PRUEBA'; $c.InvNif = ''; $c.InvActivity = 'Gestión de trámites'; $c.InvAddress = 'Calle Uno, 1'; $c.InvTown = '30700 Torre Pacheco'
$c.InvIban = 'ES00 0000 0000 0000'; $c.CompanyName = ''; $c.CompanyLogo = ''; $c.TasasCanje = 'B=33; C=100; D=100'
$script:T = $script:Tramites | Where-Object { $_.Id -eq 'canje' }; [void](Load-Data)
$withT = $null; $paidAll = $null
foreach ($r in $script:Rows) {
    if (-not $withT -and [double]$r.PriceV -gt 0 -and [double]$r.Pending -gt 0 -and (Get-TasasFor $script:T $r.Data) -gt 0) { $withT = $r }
    if (-not $paidAll -and [double]$r.PriceV -gt 0 -and [double]$r.Pending -le 0) { $paidAll = $r }
}
[void]$o.Append("IN rows=$(($script:Rows | Measure-Object).Count) conTasas=$(if ($withT) { $withT.Id } else { '-' }) pagado=$(if ($paidAll) { $paidAll.Id } else { '-' })`r`n")
[void]$o.Append("IN0 sin emisor`r`n"); New-ClientInvoice $withT
$c.InvNif = 'X1234567L'
[void]$o.Append("IN1 $($withT.Id)`r`n"); New-ClientInvoice $withT; [void]$o.Append("ST $($script:UI.StatusText.Text)`r`n")
[void]$o.Append("IN2 otra vez $($withT.Id)`r`n"); New-ClientInvoice $withT
[void]$o.Append("IN3 $($paidAll.Id)`r`n"); New-ClientInvoice $paidAll
$keepP = $withT.PriceV; $withT.PriceV = 0; [void]$o.Append("IN4 sin precio $($withT.Id)`r`n"); New-ClientInvoice $withT; $withT.PriceV = $keepP
$script:T = $script:Tramites | Where-Object { $_.Id -eq 'traduccion' }; [void](Load-Data)
$aut = @($script:Rows | Where-Object { [double]$_.PriceV -gt 0 })[0]
$c.InvIban = ''; $c.CompanyName = 'EMPRESA X'
[void]$o.Append("IN5 $(if ($aut) { $aut.Id } else { '(ninguno)' })`r`n"); if ($aut) { New-ClientInvoice $aut }

# exportar la lista: vacía -> aviso; CSV (huella del archivo); PDF apaisado (huella de cada página); diálogo cancelado -> nada
$script:T = $script:Tramites | Where-Object { $_.Id -eq 'canje' }; [void](Load-Data); $c.CompanyName = ''; $keepU = $script:User; $script:User = 'PRUEBA'
$script:ExpPath = $null
function Get-ExportFile([string]$ext, [string]$filter) { [void]$o.Append("EXF $ext $filter`r`n"); return $script:ExpPath }
$script:UI.ClientsGrid.ItemsSource = @(); [void]$o.Append("EX0 vacía`r`n"); Export-ClientsCsv; Export-ClientsPdf
$script:UI.ClientsGrid.ItemsSource = $script:Rows
$script:ExpPath = Join-Path $TestDir 'ps\lista.csv'; [void]$o.Append("EX1 csv`r`n"); Export-ClientsCsv
$cb = [IO.File]::ReadAllBytes($script:ExpPath); $md5 = [Security.Cryptography.MD5]::Create()
[void]$o.Append("CSV bytes=$($cb.Length) md5=$([BitConverter]::ToString($md5.ComputeHash($cb)).Replace('-', '').Substring(0, 12))`r`n"); $md5.Dispose()
[void]$o.Append("ST $($script:UI.StatusText.Text)`r`n")
$script:ExpPath = Join-Path $TestDir 'ps\lista.pdf'; [void]$o.Append("EX2 pdf`r`n"); Export-ClientsPdf
$script:ExpPath = $null; [void]$o.Append("EX3 cancelado`r`n"); Export-ClientsCsv; Export-ClientsPdf
$script:User = $keepU; ${function:Get-ExportFile} = $null
${function:Start-Process} = $null; ${function:Get-Date} = $null
