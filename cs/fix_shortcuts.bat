@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File tools\fix_shortcuts.ps1 > test\fix_shortcuts_err.txt 2>&1
