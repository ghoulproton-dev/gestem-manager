# Etapa 7 (sección 0): lo mismo que src\SelfTest7.cs con las funciones reales de GESTEM.ps1 (se carga con «.» desde pstest.ps1)
# rutas de la prueba iguales en los dos: <T>\X en vez de <prueba>\cs o <prueba>\ps
function Norm7([string]$s) { return ($s.Replace($TestDir, '<T>') -replace '<T>\\(cs|ps)', '<T>\X') }
function Start-Process { param([string]$FilePath, $ArgumentList) [void]$o.Append("OPEN $(Norm7 $FilePath)$(if ($ArgumentList) { ' ' + (Norm7 "$ArgumentList") })`r`n") }

# abrir el Excel: sin archivo -> aviso; con archivo y sin Excel abierto con él -> se abre el archivo
$script:T = $script:Tramites | Where-Object { $_.Id -eq 'canje' }
$keepX = $script:Cfg.ExcelPath
$script:Cfg.ExcelPath = Join-Path $TestDir 'no-existe.xlsx'; [void]$o.Append("XL0`r`n"); Open-ExcelWindow 5
$script:Cfg.ExcelPath = $keepX; [void]$o.Append("XL1`r`n"); Open-ExcelWindow 5; [void]$o.Append("XL2`r`n"); Open-ExcelWindow ($script:HandRowBase + 3)

${function:Start-Process} = $null
