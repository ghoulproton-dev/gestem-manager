# cierra SOLO los Excel ocultos que tienen abiertos archivos de la carpeta de pruebas (sin guardar)
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here
$testDir = (Join-Path $root 'test').ToLowerInvariant()
$log = New-Object Text.StringBuilder
for ($i = 0; $i -lt 10; $i++) {
    $xl = $null
    try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { break }
    $vis = $false; try { $vis = [bool]$xl.Visible } catch { }
    $books = @(); try { foreach ($w in $xl.Workbooks) { $books += "$($w.FullName)" } } catch { }
    [void]$log.AppendLine("instancia: visible=$vis libros=$($books -join ' | ')")
    $onlyTest = ($books.Count -gt 0) -and -not ($books | Where-Object { -not $_.ToLowerInvariant().StartsWith($testDir) })
    if ($vis -or -not $onlyTest) { [void]$log.AppendLine('  no se toca'); break }
    try { foreach ($w in @($xl.Workbooks)) { $w.Close($false) }; $xl.Quit() } catch { [void]$log.AppendLine("  error: $($_.Exception.Message)") }
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl); $xl = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); Start-Sleep -Seconds 2
    [void]$log.AppendLine('  cerrada')
}
[IO.File]::WriteAllText((Join-Path $root 'test\cleanup.txt'), $log.ToString())
