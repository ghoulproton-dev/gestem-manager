# Carpetas de documentos de clientes renumerados: al unirse (mismo nº, otra persona), al recibir un nº nuevo de otro PC,
# clientes de pestañas hechas a mano y carpetas que ya estaban mal. Sin red: lo «de Firebase» se pasa a mano. Clientes ficticios,
# solo carpetas temporales (de canje.json de este PC solo se lee el formato).
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$here = Join-Path $env:TEMP 'gestem-tests\docsrenum'; Remove-Item -Recurse -Force $here -ErrorAction SilentlyContinue; New-Item -ItemType Directory -Force $here | Out-Null
$fail = 0
function Check($ok, $msg) { if ($ok) { "OK   $msg" } else { "FAIL $msg"; $script:fail++ } }
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'GESTEM.ps1'), [ref]$null, [ref]$null)
foreach ($s in $ast.EndBlock.Statements) {
    if ($s -is [System.Management.Automation.Language.FunctionDefinitionAst]) { . ([ScriptBlock]::Create($s.Extent.Text)) }
    elseif ($s.Extent.StartLineNumber -lt 1150 -or $s.Extent.Text -match '^\$script:(Fb\w+|HandRowBase)\s*=') { try { . ([ScriptBlock]::Create($s.Extent.Text)) } catch { } }
}
$script:Logs = New-Object System.Collections.Generic.List[string]
function Write-Log([string]$m) { $script:Logs.Add($m) }
function Show-Toast { }
$realCfg = Get-Content -Raw (Join-Path $env:APPDATA 'GESTEM\config.json') -Encoding UTF8 | ConvertFrom-Json
$base = [IO.File]::ReadAllText((Join-Path (Join-Path (Split-Path -Parent $realCfg.ExcelPath) 'GESTEM datos') 'canje.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
$shared = Join-Path $here 'nube\GESTEM documentos'
$script:Cfg = [ordered]@{ FbSync = $true; DocsDir = $shared; ExcelPath = (Join-Path $here 'PC-A\x.xlsx'); Language = 'es' }
$script:CfgPath = Join-Path $here 'config.json'; $script:FbStatePath = Join-Path $here 'PC-A\firebase.json'
$env:COMPUTERNAME = 'PC-A'
$t = $script:Tramites | Where-Object { $_.Id -eq 'canje' }; $script:T = $t
$tdir = Join-Path $shared (Get-SafeName $t.Sheet)
function Put($rel, $text) { $p = Join-Path $tdir $rel; New-Item -ItemType Directory -Force (Split-Path -Parent $p) | Out-Null; [IO.File]::WriteAllText($p, $text) }
function Txt($rel) { $p = Join-Path $tdir $rel; if (Test-Path -LiteralPath $p) { [IO.File]::ReadAllText($p) } }
function Folders { @(Get-ChildItem -LiteralPath $tdir -Directory | ForEach-Object { $_.Name } | Sort-Object) }
function Fields($name, $doc) { $o = [ordered]@{}; foreach ($f in $t.Fields) { $o[$f.Key] = $null }; $o.name = $name; $o.doc = $doc; return $o }
function Cloud($n, $name, $doc, $uid, $pc = 'PC-B') {
    @{ T = 'canje'; Uid = $uid; Id = ('CANJ-2026-{0:000}' -f $n); Datos = ((Fields $name $doc) | ConvertTo-Json -Compress); Updated = '2026-11-18T12:00:00'; By = 'OTRA'; Pc = $pc; Deleted = $false }
}
function Id-Of($name) { "$(((Read-Store $t).Clients | Where-Object { $_.F['name'] -eq $name } | Select-Object -First 1).Id)" }

# PC-A: ANA 001, BOB 002, CARL 003 y una de pestaña hecha a mano (HANA); en Firebase el 002 es ZOE (otra persona)
New-Item -ItemType Directory -Force (Join-Path $here 'PC-A\GESTEM datos') | Out-Null
$j = $base.PSObject.Copy()
$j.clients = @(foreach ($c in @(@{ n = 1; name = 'ANA RUIZ'; doc = 'X1111111A' }, @{ n = 2; name = 'BOB LEE'; doc = 'Y2222222B' }, @{ n = 3; name = 'CARL DIAZ'; doc = 'Z3333333C' })) {
    [ordered]@{ row = 10 + $c.n; id = 'CANJ-2026-{0:000}' -f $c.n; sync = $true; nuevo = $false; updated = '2026-11-10T10:00:00'; by = 'PC-A'; f = (Fields $c.name $c.doc) } })
$j.adoptados = @()
[IO.File]::WriteAllText((Join-Path $here 'PC-A\GESTEM datos\canje.json'), ($j | ConvertTo-Json -Depth 20), (New-Object Text.UTF8Encoding($true)))
$s = Read-Store $t; $FV = Fields 'HANA MANO' $null; $h = @{}; foreach ($k in $FV.Keys) { $h[$k] = $FV[$k] }
$s.Clients.Add(@{ Row = $script:HandRowBase + 5; Id = ''; F = $h; Sync = $true; New = $false; Updated = '2026-11-10T10:00:00'; By = 'Excel' }); Write-Store $t $s
Put 'CANJ-2026-001 - ANA RUIZ\a.pdf' 'ana'
Put 'CANJ-2026-002 - BOB LEE\nie.pdf' 'bob'
Put 'CANJ-2026-003 - CARL DIAZ\c.pdf' 'carl'
Put 'fila 100005 - HANA MANO\h.pdf' 'hana'
Put 'CANJ-2026-002 - ZOE PEREZ\z.pdf' 'zoe'    # la de ZOE, subida por el otro PC a la carpeta compartida

'--- carpeta del cliente: con dos carpetas del mismo nº, la de su nombre ---'
Check ((Split-Path -Leaf (Get-ClientDocsDir @{ Id = 'CANJ-2026-002'; Name = 'ZOE PEREZ'; Row = 12 })) -eq 'CANJ-2026-002 - ZOE PEREZ') 'ZOE abre la suya, no la de BOB'

'--- PC-A se une: el 002 de Firebase es ZOE, BOB pasa a otro nº ---'
$cloud = @((Cloud 1 'ANA RUIZ' 'X1111111A' 'canje-CANJ-2026-001'), (Cloud 2 'ZOE PEREZ' 'W4444444D' 'canje-CANJ-2026-002'))
Join-FbDocs $cloud ''
$bob = Id-Of 'BOB LEE'; $hana = Id-Of 'HANA MANO'
Check ($bob -and $bob -ne 'CANJ-2026-002') "BOB renumerado a $bob"
Check ((Txt "$bob - BOB LEE\nie.pdf") -eq 'bob' -and -not (Test-Path (Join-Path $tdir 'CANJ-2026-002 - BOB LEE'))) "su carpeta pasa a «$bob - BOB LEE»"
Check ((Txt 'CANJ-2026-002 - ZOE PEREZ\z.pdf') -eq 'zoe') 'la de ZOE no se toca'
Check ($hana -like 'CANJ-*' -and (Txt "$hana - HANA MANO\h.pdf") -eq 'hana' -and -not (Test-Path (Join-Path $tdir 'fila 100005 - HANA MANO'))) "HANA (pestaña a mano) pasa a $hana con su carpeta"
Check ((Txt 'CANJ-2026-001 - ANA RUIZ\a.pdf') -eq 'ana' -and (Txt 'CANJ-2026-003 - CARL DIAZ\c.pdf') -eq 'carl') 'las de ANA y CARL (sin cambio de nº) siguen igual'
Check ((Get-DocsCount @{ Id = $bob; Row = 0 }) -ge 0 -and (Split-Path -Leaf (Get-ClientDocsDir @{ Id = $bob; Name = 'BOB LEE' })) -eq "$bob - BOB LEE") 'BOB abre su carpeta con el nº nuevo'

'--- otro PC renumera a CARL (003 -> 009): se recibe y su carpeta le sigue ---'
$carlUid = "$(((Read-Store $t).Clients | Where-Object { $_.F['name'] -eq 'CARL DIAZ' }).Uid)"
$d = Cloud 9 'CARL DIAZ' 'Z3333333C' $carlUid; $d.Updated = '2026-11-18T13:00:00'
Receive-FbDocs @($d)
Check ((Id-Of 'CARL DIAZ') -eq 'CANJ-2026-009' -and (Txt 'CANJ-2026-009 - CARL DIAZ\c.pdf') -eq 'carl' -and -not (Test-Path (Join-Path $tdir 'CANJ-2026-003 - CARL DIAZ'))) 'carpeta de CARL pasa a 009'

'--- carpetas que ya estaban mal (al abrir) ---'
Put 'CANJ-2026-040 - ANA RUIZ\a.pdf' 'otra ana'    # nº de nadie, nombre de ANA: se junta con la suya sin pisar
Put 'CANJ-2026-001 - ANA\vieja.pdf' 'ana corta'    # mismo cliente con el nombre de antes (más corto): no se toca
Put 'CANJ-2026-050 - NADIE\x.pdf' 'x'               # nadie con ese nombre: no se toca
$s = Read-Store $t; $FV = Fields 'SARA YA' $null; $h = @{}; foreach ($k in $FV.Keys) { $h[$k] = $FV[$k] }
$s.Clients.Add(@{ Row = 30; Id = 'CANJ-2026-020'; Uid = 'x1'; F = $h; Sync = $true; New = $false; Updated = '2026-11-10'; By = 'A' })
$h2 = @{}; foreach ($k in $h.Keys) { $h2[$k] = $h[$k] }; $s.Clients.Add(@{ Row = 31; Id = 'CANJ-2026-021'; Uid = 'x2'; F = $h2; Sync = $true; New = $false; Updated = '2026-11-10'; By = 'A' })
$script:FbApplying = $true; Write-Store $t $s; $script:FbApplying = $false
Put 'CANJ-2026-060 - SARA YA\s.pdf' 'sara'          # dos SARA YA: no se sabe cuál, no se toca
Repair-AllDocsFolders
Check ((Txt 'CANJ-2026-001 - ANA RUIZ\a.pdf') -eq 'ana' -and (Txt 'CANJ-2026-001 - ANA RUIZ\a (2).pdf') -eq 'otra ana' -and -not (Test-Path (Join-Path $tdir 'CANJ-2026-040 - ANA RUIZ'))) 'nº equivocado: se junta con la de ANA, los dos archivos'
Check ((Txt 'CANJ-2026-001 - ANA\vieja.pdf') -eq 'ana corta') 'nombre corregido del mismo cliente: no se mueve'
Check ((Txt 'CANJ-2026-050 - NADIE\x.pdf') -eq 'x') 'sin cliente con ese nombre: no se mueve'
Check ((Txt 'CANJ-2026-060 - SARA YA\s.pdf') -eq 'sara') 'dos clientes con el nombre: no se mueve'
$before = (Folders) -join '|'; Repair-AllDocsFolders
Check (((Folders) -join '|') -eq $before) 'otra vez: no cambia nada'
"carpetas: $((Folders) -join ' · ')"

Remove-Item -Recurse -Force $here -ErrorAction SilentlyContinue
''
$script:Logs | Where-Object { $_ -match '^documentos|renumer|pasa a|es otro' } | ForEach-Object { "log: $_" }
"fallos: $fail"
