# Bloques de datos de GESTEM.ps1 (fuera de funciones) buscados por marcas, no por nº de línea.
# Uso: $code = & (Join-Path $here 'psblocks.ps1') $lines; foreach ($c in $code) { . ([ScriptBlock]::Create($c)) }
param([string[]]$lines)
function At([string]$re, [int]$from = 0) { for ($i = $from; $i -lt $lines.Count; $i++) { if ($lines[$i] -match $re) { return $i } }; throw "no se encuentra: $re" }
$a = At '^\$script:Nacionalidades = '; $b = At '^function L\(' $a
$c = At '^\$script:SheetAliases = @\{'; $d = At '^\}' $c
$e = At '^\$script:AdmTr = @\{'; $f = At '^\$script:TrText = ' $e
@(($lines[$a..($b - 1)] -join "`n"), ($lines[$c..$d] -join "`n"), $lines[(At '^\$script:GenStatusList = ')], ($lines[$e..$f] -join "`n"), $lines[(At '^\$script:OwnerPrints = ')], $lines[(At '^\$script:LockKeys = ')], $lines[(At '^\$script:AppVersion = ')])
