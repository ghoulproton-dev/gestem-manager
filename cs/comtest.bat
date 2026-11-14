@echo off
cd /d "%~dp0"
copy /Y "..\GESTEM MANAGER\CRM 2026CLIENTES (copia real).xlsx" "test\orig.xlsx" >nul
powershell -NoProfile -ExecutionPolicy Bypass -File tools\comtest.ps1 > test\comtest_err.txt 2>&1
