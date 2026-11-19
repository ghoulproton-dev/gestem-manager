@echo off
cd /d "%~dp0"
rem datos de prueba: ?GESTEM MANAGER? junto a cs (copia de trabajo) o en Descargas
set SRC=..\GESTEM MANAGER
if not exist "%SRC%\GESTEM datos" set SRC=%USERPROFILE%\Downloads\GESTEM_MANAGER_1\GESTEM MANAGER
rmdir /s /q test 2>nul
mkdir "test\cs\GESTEM datos"
mkdir "test\ps\GESTEM datos"
copy /Y "%SRC%\GESTEM datos\*.json" "test\cs\GESTEM datos\" >nul
copy /Y "%SRC%\GESTEM datos\*.json" "test\ps\GESTEM datos\" >nul
copy /Y "%SRC%\CRM 2026CLIENTES (copia real).xlsx" "test\cs\x.xlsx" >nul
copy /Y "%SRC%\CRM 2026CLIENTES (copia real).xlsx" "test\ps\x.xlsx" >nul
rem documentos de prueba (etapa 3d): la misma carpeta para los dos
powershell -NoProfile -ExecutionPolicy Bypass -File tools\docsfixture.ps1 "%cd%\test\ps_docs" "%SRC%\GESTEM datos\canje.json" "%cd%\test\docs_src" > test\fixture_err.txt 2>&1
powershell -NoProfile -ExecutionPolicy Bypass -File tools\docsfixture.ps1 "%cd%\test\cs_docs" "%SRC%\GESTEM datos\canje.json" "%cd%\test\docs_src" >> test\fixture_err.txt 2>&1
powershell -NoProfile -ExecutionPolicy Bypass -File tools\pstest.ps1 "%cd%\test" > test\ps_err.txt 2>&1
start /wait "" GESTEM.exe --selftest "%cd%\test"
powershell -NoProfile -ExecutionPolicy Bypass -File tools\compare.ps1 "%cd%\test" > test\compare_err.txt 2>&1
rem ventana (etapa 3): datos sin tocar en test\ui -> test\ui\ui_summary.txt
mkdir "test\ui\cs\GESTEM datos"
copy /Y "%SRC%\GESTEM datos\*.json" "test\ui\cs\GESTEM datos\" >nul
copy /Y "%SRC%\CRM 2026CLIENTES (copia real).xlsx" "test\ui\cs\x.xlsx" >nul
start /wait "" GESTEM.exe --uitest "%cd%\test\ui"
copy /Y "%SRC%\CRM 2026CLIENTES (copia real).xlsx" "test\orig.xlsx" >nul
powershell -NoProfile -ExecutionPolicy Bypass -File tools\comtest.ps1 > test\comtest_err.txt 2>&1
echo done> test\done.txt
