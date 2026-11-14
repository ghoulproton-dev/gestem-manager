@echo off
cd /d "%~dp0"
set FW=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319
set WPF=%FW%\WPF
del build.log smoke.log GESTEM.exe 2>nul
set RES=/resource:main.xaml,main.xaml /resource:data\tramites.json,tramites.json /resource:data\text.json,text.json /resource:data\lists.json,lists.json /resource:data\admin_tr.json,admin_tr.json
"%FW%\csc.exe" /nologo /target:winexe /optimize+ /out:GESTEM.exe /win32icon:gestem.ico /r:"%WPF%\PresentationFramework.dll" /r:"%WPF%\PresentationCore.dll" /r:"%WPF%\WindowsBase.dll" /r:"%FW%\System.Xaml.dll" /r:System.dll /r:System.Core.dll /r:System.Xml.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:System.Net.Http.dll /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll /r:System.Security.dll /r:Microsoft.CSharp.dll %RES% src\*.cs > build.log 2>&1
echo csc_exit=%ERRORLEVEL% >> build.log
if exist GESTEM.exe start /wait "" GESTEM.exe --smoke
if exist smoke.log type smoke.log >> build.log
