# ¿Excel abre estos archivos? (original, copia de prueba antes y después de escribir directo)
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here
$out = Join-Path $root 'test\comtest.txt'
$files = @(
    (Join-Path (Split-Path -Parent $root) 'GESTEM MANAGER\CRM 2026CLIENTES (copia real).xlsx'),
    (Join-Path $root 'test\orig.xlsx'),
    (Join-Path $root 'test\cs\x.xlsx'),
    (Join-Path $root 'test\ps\x.xlsx'),
    (Join-Path (Split-Path -Parent $root) 'GESTEM MANAGER\CRM_PRUEBA3.xlsx'),
    (Join-Path $env:APPDATA 'GESTEM\copias\CRM 2026CLIENTES (copia real)_20261113_214443.xlsx'),
    (Join-Path $root 'test\nuevo.xlsx')
)
$files += @(Get-ChildItem -Path (Join-Path $root 'test\fix') -Filter '*.xlsx' | Sort-Object Name | ForEach-Object { $_.FullName })
$files += @(Get-ChildItem -Path (Join-Path $env:APPDATA 'GESTEM\copias') -Filter 'CRM 2026CLIENTES*.xlsx' | Sort-Object Name | ForEach-Object { $_.FullName })
$log = New-Object Text.StringBuilder
try { $act = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application'); [void]$log.AppendLine("Excel ya abierto: libros=$($act.Workbooks.Count) visible=$($act.Visible)") } catch { [void]$log.AppendLine('Excel no estaba abierto') }
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
[void]$log.AppendLine("version=$($xl.Version)")
try { $nb = $xl.Workbooks.Add(); $nb.Worksheets.Item(1).Range('A1').Value2 = 'prueba'; $nb.SaveAs((Join-Path $root 'test\nuevo.xlsx'), 51); $nb.Close($false); [void]$log.AppendLine('libro nuevo: OK') } catch { [void]$log.AppendLine("libro nuevo: ERR $($_.Exception.Message)") }
try { [void]$log.AppendLine("ProtectedView=$($xl.ProtectedViewWindows.Count) Interactive=$($xl.Interactive) Ready=$($xl.Ready)") } catch { }
foreach ($f in $files) {
    try {
        $wb = $xl.Workbooks.Open([IO.Path]::GetFullPath($f))
        [void]$log.AppendLine("OK  $f  hojas=$($wb.Worksheets.Count) readonly=$($wb.ReadOnly)")
        $wb.Close($false)
    } catch { [void]$log.AppendLine("ERR $f  $($_.Exception.Message) hr=$($_.Exception.HResult)") }
}
$xl.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
[IO.File]::WriteAllText($out, $log.ToString())
