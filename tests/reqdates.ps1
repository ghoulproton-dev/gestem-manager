# Requerimientos: plazo de 10 días hábiles, aviso del menú, fechas mm/dd de la pestaña y pestañas -> trámites. Sin red ni datos reales.
$fail = 0
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\GESTEM.ps1'), [ref]$null, [ref]$null)
foreach ($s in $ast.EndBlock.Statements) { if ($s -is [System.Management.Automation.Language.FunctionDefinitionAst]) { . ([ScriptBlock]::Create($s.Extent.Text)) } elseif ($s.Extent.StartLineNumber -lt 1150) { try { . ([ScriptBlock]::Create($s.Extent.Text)) } catch { } } }
function D($s) { [DateTime]::ParseExact($s, 'yyyy-MM-dd', $null) }
function C($ok, $m) { if ($ok) { "OK   $m" } else { "FAIL $m"; $script:fail++ } }
C ((Get-ReqDeadline $null (D '2026-11-02')) -eq (D '2026-11-16')) 'notificación lun 02/11 -> lun 16/11'
C ((Get-ReqDeadline (D '2026-11-06') (D '2026-11-02')) -eq (D '2026-11-20')) 'apertura vie 06/11 manda -> vie 20/11'
C ((Get-ReqDeadline (D '2026-11-07') $null) -eq (D '2026-11-20')) 'apertura sáb 07/11 -> vie 20/11'
C ($null -eq (Get-ReqDeadline $null $null)) 'sin fechas -> sin plazo'
C ((Get-WorkDaysUntil (D '2026-11-13') (D '2026-11-16')) -eq 1) 'vie->lun = 1 día hábil'
$rows = New-Object System.Collections.Generic.List[object]; foreach ($x in @(@{ Data = @{ xF = (D '2026-11-02') } }, @{ Data = @{ xF = (D '2026-11-10') } }, @{ Data = @{ xF = (D '2026-11-02'); xI = 'Respondido' } }, @{ Data = @{ xF = (D '2026-11-17') } })) { $rows.Add($x) }    # List, como en la caché de la app
$a = Get-ReqAlerts $rows (D '2026-11-18'); C ($a.Late -eq 1 -and $a.Due -eq 0) "aviso a 18/11 (10/11 vence el 24/11, a 4 días hábiles): vencidos $($a.Late), por vencer $($a.Due)"
C ((ConvertFrom-HandDate '11/03/2026' -MDY) -eq (D '2026-11-03') -and (ConvertFrom-HandDate '11/03/2026') -eq (D '2026-03-11')) 'texto mm/dd sin cambiar día y mes'
$t = $script:Tramites | Where-Object Id -eq 'requerimiento'
$cells = @{ 1 = @{ 2 = 'NOTIFICACION'; 3 = 'RESPUESTA' }; 3 = @{ 1 = 'PRUEBA UNO'; 2 = [double]46063; 3 = '11/03/2026' } }
$r = (ConvertFrom-HandSheet $cells $t)[0]
C ($r.F['xF'] -eq [DateTime]::FromOADate(46063).Date -and $r.F['xH'] -eq (D '2026-11-03') -and $r.F['flim'] -is [DateTime]) "pestaña: notif $($r.F['xF'].ToString('dd/MM/yyyy')), resp $($r.F['xH'].ToString('dd/MM/yyyy')), límite $($r.F['flim'].ToString('dd/MM/yyyy'))"
$names = 'REQUERIMIENTOS', 'REGULARIZACIÓN EXTRAORDINARIA', 'EXTRANJERIA'
foreach ($n in $names) { C (@($script:Tramites | Where-Object { (Get-TabCandidates @($n) $_).Count }).Id -join ',') "pestaña «$n» -> $(@($script:Tramites | Where-Object { (Get-TabCandidates @($n) $_).Count }).Id -join ',')" }
$a = Get-ReqAlerts $rows (D '2026-11-19'); C ($a.Late -eq 1 -and $a.Due -eq 1) "aviso a 19/11: vencidos $($a.Late), por vencer $($a.Due)"
$e = Get-ReqAlerts (New-Object System.Collections.Generic.List[object]) (D '2026-11-19'); C ($e.Late + $e.Due -eq 0) 'lista vacía: sin aviso'
"fallos: $fail"
