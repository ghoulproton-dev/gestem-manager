@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File tools\cleanup.ps1 > test\cleanup_err.txt 2>&1
