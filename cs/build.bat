@echo off
cd /d "%~dp0"
set FW=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319
set WPF=%FW%\WPF
rem se compila a GESTEM.new.exe: el GESTEM.exe anterior solo se cambia si la compilacion y la prueba de arranque salen bien
del build.log smoke.log GESTEM.new.exe 2>nul
set RES=/resource:main.xaml,main.xaml /resource:data\tramites.json,tramites.json /resource:data\text.json,text.json /resource:data\lists.json,lists.json /resource:data\admin_tr.json,admin_tr.json /resource:gestem.ico,gestem.ico
"%FW%\csc.exe" /nologo /target:winexe /optimize+ /out:GESTEM.new.exe /win32icon:gestem.ico /r:"%WPF%\PresentationFramework.dll" /r:"%WPF%\PresentationCore.dll" /r:"%WPF%\WindowsBase.dll" /r:"%FW%\System.Xaml.dll" /r:System.dll /r:System.Core.dll /r:System.Xml.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:System.Net.Http.dll /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll /r:System.Security.dll /r:Microsoft.CSharp.dll /r:Microsoft.VisualBasic.dll %RES% src\*.cs > build.log 2>&1
set CSC=%ERRORLEVEL%
echo csc_exit=%CSC% >> build.log
if not "%CSC%"=="0" goto keep
if not exist GESTEM.new.exe goto keep
start /wait "" GESTEM.new.exe --smoke
if exist smoke.log (type smoke.log >> build.log & echo.>> build.log)
findstr /b /c:"OK xaml" smoke.log >nul 2>&1 || goto keep
move /y GESTEM.new.exe GESTEM.exe >nul 2>&1 || goto locked
echo build=OK >> build.log
exit /b 0
:locked
echo build=NUEVO EN GESTEM.new.exe: no se pudo cambiar GESTEM.exe (abierto?) >> build.log
exit /b 1
:keep
del GESTEM.new.exe 2>nul
echo build=FALLO: se conserva el GESTEM.exe anterior >> build.log
exit /b 1
