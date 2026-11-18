# Carpeta de documentos compartida: dos PCs encuentran la misma «GESTEM documentos» en la nube y pasan allí sus documentos
# sin sobrescribir ninguno. Solo carpetas temporales y una config temporal; no toca la config ni los documentos reales.
$fail = 0
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\GESTEM.ps1'), [ref]$null, [ref]$null)
foreach ($s in $ast.EndBlock.Statements) { if ($s -is [System.Management.Automation.Language.FunctionDefinitionAst]) { . ([ScriptBlock]::Create($s.Extent.Text)) } elseif ($s.Extent.StartLineNumber -lt 1150) { try { . ([ScriptBlock]::Create($s.Extent.Text)) } catch { } } }
function C($ok, $m) { if ($ok) { "OK   $m" } else { "FAIL $m"; $script:fail++ } }
$script:Logs = New-Object System.Collections.Generic.List[string]
function Write-Log($m) { $script:Logs.Add("$m") }
function Show-Toast { }
$root = Join-Path $env:TEMP 'gestem-tests\docs'
Remove-Item -Recurse -Force $root -ErrorAction SilentlyContinue
$script:CfgPath = Join-Path $root 'config.json'    # nunca la config real
$script:T = $script:Tramites[0]
function Put($path, $text) { New-Item -ItemType Directory -Force (Split-Path -Parent $path) | Out-Null; [IO.File]::WriteAllText($path, $text) }
function Txt($path) { if (Test-Path -LiteralPath $path) { [IO.File]::ReadAllText($path) } }
$cloudRoot = Join-Path $root 'OneDrive'
$shared = Join-Path $cloudRoot 'Oficina\GESTEM documentos'
$ana = 'DGT\CANJ-2026-001 - ANA'
Put (Join-Path $shared "$ana\pasaporte.pdf") 'nube'
Put (Join-Path $shared "$ana\igual.pdf") 'mismo'
function Get-CloudRoots { @($script:Roots) }
function Use-Pc($pc, [bool]$sync = $true, $roots = @($cloudRoot)) {
    $script:Roots = $roots
    $script:Cfg = [ordered]@{ FbSync = $sync; DocsDir = ''; ExcelPath = (Join-Path $root "$pc\CRM.xlsx"); Language = 'es' }
    return (Join-Path $root "$pc\GESTEM documentos")
}

'--- PC-A: tiene documentos propios, uno con el mismo nombre que en la nube ---'
$locA = Use-Pc 'PC-A'
Put (Join-Path $locA "$ana\pasaporte.pdf") 'de A'
Put (Join-Path $locA "$ana\igual.pdf") 'mismo'
Put (Join-Path $locA "$ana\nuevo.pdf") 'nuevo A'
Put (Join-Path $locA 'DGT\CANJ-2026-002 - BOB\nie.pdf') 'bob'
C ((Get-DocsRoot) -eq $locA) 'antes: la carpeta de al lado del Excel'
Initialize-DocsFolder
C ($script:Cfg.DocsDir -eq $shared -and (Get-DocsRoot) -eq $shared) 'A usa la carpeta compartida'
C ((Txt (Join-Path $shared "$ana\pasaporte.pdf")) -eq 'nube') 'el de la nube no se sobrescribe'
C ((Txt (Join-Path $shared "$ana\pasaporte (2).pdf")) -eq 'de A') 'el de A, mismo nombre: se guardan los dos'
C (@(Get-ChildItem (Join-Path $shared $ana) -Filter 'igual*').Count -eq 1) 'mismo nombre y mismo contenido: uno solo'
C ((Txt (Join-Path $shared "$ana\nuevo.pdf")) -eq 'nuevo A' -and (Txt (Join-Path $shared 'DGT\CANJ-2026-002 - BOB\nie.pdf')) -eq 'bob') 'los nuevos de A pasan'
C (-not (Test-Path $locA)) 'la carpeta local de A queda vacía y se quita'
C ((Get-Content -Raw $script:CfgPath | ConvertFrom-Json).DocsDir -eq $shared) 'guardado en la config'

'--- PC-B: la misma carpeta, otro pasaporte distinto con el mismo nombre ---'
$locB = Use-Pc 'PC-B'
Put (Join-Path $locB "$ana\pasaporte.pdf") 'de B'
Put (Join-Path $locB "$ana\nuevo.pdf") 'nuevo A'
Initialize-DocsFolder
C ((Get-DocsRoot) -eq $shared) 'B usa la misma carpeta que A'
C ((Txt (Join-Path $shared "$ana\pasaporte (3).pdf")) -eq 'de B' -and (Txt (Join-Path $shared "$ana\pasaporte (2).pdf")) -eq 'de A') 'el de B también se guarda, sin pisar el de A'
C (@(Get-ChildItem (Join-Path $shared $ana) -Filter 'nuevo*').Count -eq 1) 'B no duplica lo que A ya subió'
$script:T = $script:Tramites[0]; $tdA = Get-TramiteDocsDir
C ($tdA.StartsWith($shared)) "los dos PCs ven los documentos del trámite en $tdA"
$n = @(Get-ChildItem $shared -Recurse -File).Count
Initialize-DocsFolder
C (@(Get-ChildItem $shared -Recurse -File).Count -eq $n) 'al volver a abrir no hace nada más'

'--- PC con la sincronización apagada (datos de prueba): no pasa nada ---'
$locC = Use-Pc 'PC-C' $false
Put (Join-Path $locC "$ana\prueba.pdf") 'prueba'
Initialize-DocsFolder
C (-not $script:Cfg.DocsDir -and (Test-Path (Join-Path $locC "$ana\prueba.pdf")) -and -not (Test-Path (Join-Path $shared "$ana\prueba.pdf"))) 'sigue en su carpeta y no llega a la compartida'

'--- sin carpeta en la nube: como antes ---'
$locD = Use-Pc 'PC-D' $true @()
Initialize-DocsFolder
C (-not $script:Cfg.DocsDir -and (Get-DocsRoot) -eq $locD) 'sigue la de al lado del Excel'

'--- Ajustes: A cambia a otra carpeta: desde la compartida solo se copia ---'
[void](Use-Pc 'PC-A'); $script:Cfg.DocsDir = $shared
$other = Join-Path $root 'Drive\GESTEM documentos'
Put (Join-Path $other "$ana\pasaporte.pdf") 'otra'
[void](Set-DocsRoot $other)
C ((Get-DocsRoot) -eq $other) 'A usa la nueva carpeta'
C (@(Get-ChildItem $shared -Recurse -File).Count -eq $n) 'la compartida sigue entera para los otros PCs'
C ((Txt (Join-Path $other "$ana\pasaporte.pdf")) -eq 'otra' -and (Txt (Join-Path $other "$ana\pasaporte (4).pdf")) -in 'nube', 'de A', 'de B') 'en la nueva no se sobrescribe nada'
C (@(Get-ChildItem $other -Recurse -File).Count -eq $n + 1) "la nueva tiene todo ($(@(Get-ChildItem $other -Recurse -File).Count))"

'--- Ajustes: desde la carpeta local se mueve ---'
$locE = Use-Pc 'PC-E' $true @()
Put (Join-Path $locE 'x\a.pdf') 'a'
$dst = Join-Path $root 'E2\GESTEM documentos'
[void](Set-DocsRoot $dst)
C ((Txt (Join-Path $dst 'x\a.pdf')) -eq 'a' -and -not (Test-Path $locE)) 'pasados y la local se quita'
$m = Merge-DocsFolder $dst $dst; C ($m.Moved -eq 0 -and @(Get-ChildItem $dst -Recurse -File).Count -eq 1) 'a la misma carpeta: no hace nada'

Remove-Item -Recurse -Force $root -ErrorAction SilentlyContinue
''
$script:Logs | ForEach-Object { "log: $_" }
"fallos: $fail"
