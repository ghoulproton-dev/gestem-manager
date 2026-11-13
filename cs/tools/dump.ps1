# Exporta los datos fijos de GESTEM.ps1 (trámites, textos, traducciones) a JSON para la versión C#.
# Se ejecuta con dump.bat; lee tools\GESTEM.ps1 (misma versión que el repo).
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$lines = [IO.File]::ReadAllLines((Join-Path $here 'GESTEM.ps1'), [Text.Encoding]::UTF8)
function Run-Range([int]$a, [int]$b) { $code = ($lines[($a - 1)..($b - 1)] -join "`n"); . ([ScriptBlock]::Create($code)) }
. ([ScriptBlock]::Create(($lines[60..786] -join "`n")))
. ([ScriptBlock]::Create(($lines[887..890] -join "`n")))
. ([ScriptBlock]::Create(($lines[936] )))
. ([ScriptBlock]::Create(($lines[6371..6507] -join "`n")))
$out = Join-Path (Split-Path -Parent $here) 'data'
New-Item -ItemType Directory -Force -Path $out | Out-Null
function Save([string]$name, $obj) {
    $j = ConvertTo-Json -InputObject $obj -Depth 12 -Compress
    [IO.File]::WriteAllText((Join-Path $out $name), $j, (New-Object Text.UTF8Encoding $false))
}
foreach ($t in $script:Tramites) { if ($t.HandCols) { $h = @{}; foreach ($k in $t.HandCols.Keys) { $h["$k"] = $t.HandCols[$k] }; $t.HandCols = $h } }
Save 'tramites.json' $script:Tramites
Save 'text.json' $script:Text
Save 'lists.json' @{ nationalities = $script:Nacionalidades; payments = $script:FormasPago; genStatus = $script:GenStatusList; sheetAliases = $script:SheetAliases; langs = @($script:Langs.Keys); langNames = @($script:Langs.Values) }
Save 'admin_tr.json' @{ adm = $script:AdmTr; help = $script:HelpTr; about = $script:AboutTr }
"OK tramites=$($script:Tramites.Count) textKeys=$($script:Text['es'].Count) adm=$($script:AdmTr.Count) help=$($script:HelpTr.Count)" | Set-Content -Encoding UTF8 (Join-Path $out 'dump.log')
