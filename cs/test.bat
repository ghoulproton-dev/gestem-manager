@echo off
cd /d "%~dp0"
rmdir /s /q test 2>nul
mkdir "test\cs\GESTEM datos"
mkdir "test\ps\GESTEM datos"
copy /Y "..\GESTEM MANAGER\GESTEM datos\*.json" "test\cs\GESTEM datos\" >nul
copy /Y "..\GESTEM MANAGER\GESTEM datos\*.json" "test\ps\GESTEM datos\" >nul
powershell -NoProfile -ExecutionPolicy Bypass -File tools\pstest.ps1 "%cd%\test" > test\ps_err.txt 2>&1
start /wait "" GESTEM.exe --selftest "%cd%\test"
echo done> test\done.txt
