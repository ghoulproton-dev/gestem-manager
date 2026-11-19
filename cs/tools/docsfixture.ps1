# Carpeta de documentos de prueba (etapa 3d): la misma para pstest.ps1 y GESTEM.exe --selftest.
# Usa los clientes de canje (store de entrada) con nombre único. Fechas fijas para que la lista salga igual.
param([string]$Root, [string]$StoreJson, [string]$SrcDir)
$ErrorActionPreference = 'Stop'
$s = [IO.File]::ReadAllText($StoreJson, [Text.Encoding]::UTF8) | ConvertFrom-Json
function Safe([string]$x) { return (("$x" -replace '[\\/:*?"<>|]', ' ') -replace '\s+', ' ').Trim().TrimEnd('.') }
function Key([string]$x) { $d = $x.Trim().ToUpperInvariant().Normalize([Text.NormalizationForm]::FormD); $d = -join ($d.ToCharArray() | ? { [Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne 'NonSpacingMark' }); return (($d -replace '[^A-Z0-9 ]', ' ') -replace '\s+', ' ').Trim() }
$cl = @($s.clients | ? { -not $_.borrado -and [int]$_.row -lt 100000 -and "$($_.id)" -and "$($_.f.name)".Trim() } | Sort-Object { [int]$_.row })
$cnt = @{}; foreach ($c in $cl) { $k = Key (Safe $c.f.name); $cnt[$k] = 1 + [int]$cnt[$k] }
$c = @($cl | ? { $cnt[(Key (Safe $_.f.name))] -eq 1 -and ("$($_.f.name)".Trim() -split '\s+').Count -ge 2 })
if ($c.Count -lt 9) { throw "pocos clientes para la prueba de documentos: $($c.Count)" }
$tdir = Join-Path $Root 'DGT ACTUALIZADO'
$day = New-Object DateTime 2026, 11, 10, 9, 30, 0
function F([string]$dir, [string]$name, $content, [double]$min) {
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $p = Join-Path $dir $name
    if ($content -is [int]) { [IO.File]::WriteAllBytes($p, (New-Object byte[] $content)) } else { [IO.File]::WriteAllText($p, $content) }
    (Get-Item -LiteralPath $p).LastWriteTime = $day.AddMinutes($min)
}
function D($x) { return (Join-Path $tdir ("$(Safe $x.id) - $(Safe $x.f.name)")) }
F (D $c[0]) 'a.pdf' 'AAA' 0; F (D $c[0]) 'b.jpg' 'BB' 1; F (D $c[0]) 'grande.docx' 1572864 2; F (D $c[0]) 'c.pdf' 'CC' 2.5    # se ve con la misma hora que grande.docx (09:32) pero no empata: Sort-Object de PS 5.1 no fija el orden de los empates
F (D $c[0]) 'desktop.ini' 'x' 3; (Get-Item -LiteralPath (Join-Path (D $c[0]) 'desktop.ini') -Force).Attributes = 'Hidden'
$mis = Join-Path $tdir ("$(Safe $c[2].id) - $(Safe $c[1].f.name)")              # nº de otro cliente: pasa al de c1 (que ya tiene carpeta)
F $mis 'x.pdf' 'XX' 10; F $mis 'y.pdf' 'Y1' 11
F (D $c[1]) 'x.pdf' 'XX' 10; F (D $c[1]) 'y.pdf' 'Y22' 12
F (Join-Path $tdir ("$(Safe $c[3].id) - $(Safe $c[4].f.name)")) 'z.pdf' 'ZZ' 20    # pasa a c4 (sin carpeta)
F (Join-Path $tdir ("$(Safe $c[5].id) - NADIE CON ESTE NOMBRE")) 'n.pdf' 'N' 30     # nadie: no se toca
F (Join-Path $tdir ("$(Safe $c[6].id) - $((Safe $c[6].f.name).Split(' ')[0])")) 'p.pdf' 'P' 40   # nombre corto del mismo: no se toca
F (Join-Path $tdir (Safe $c[7].id)) 'q.pdf' 'Q' 50                                   # sin « - »: cuenta igual
F $SrcDir 'nuevo.pdf' 'NUEVO' 60; F $SrcDir 'a.pdf' 'OTRA A' 61
# para las pruebas: qué clientes son (nº)
Set-Content -LiteralPath (Join-Path $Root 'clientes.txt') -Value (@($c[0..8] | % { "$($_.id)" }) -join '|') -Encoding UTF8
