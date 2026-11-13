@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File dump.ps1 > dump_err.log 2>&1
