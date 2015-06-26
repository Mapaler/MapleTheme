@echo off
title 安装《枫谷主题》运行库 / Setup "MapleTheme" Runtime
for /f "tokens=1* delims=[" %%a in ('ver') do set b=%%b
set b=%b:* =%
goto %b:~0,4%%PROCESSOR_ARCHITECTURE:~-1%

:5.1.6
echo You are xp_32bit
set sys=system32
goto setup
pause

:5.1.4
echo You are xp_64bit
set sys=system64
goto setup
pause

:5.2.6
echo You are 2003_32bit
set sys=system32
goto setup
pause

:5.2.4
echo You are 2003_64bit
set sys=system64
goto setup
pause

:6.0.6
echo You are vista_32bit
set sys=system32
goto setup
pause

:6.0.4
echo You are vista_64bit
set sys=SysWOW64
goto setup
pause

:6.1.6
echo You are win7_32bit
set sys=system32
goto setup
pause

:6.1.4
echo You are win7_64bit
set sys=SysWOW64
goto setup
pause

:6.2.6
echo You are win8_32bit
set sys=system32
goto setup
pause

:6.2.4
echo You are win8_64bit
set sys=SysWOW64
goto setup
pause

:setup
cd /d "%~dp0"
If "%PROCESSOR_ARCHITECTURE%"=="AMD64" (Set b=%SystemRoot%\SysWOW64) Else (Set b=%SystemRoot%\system32)
Rd "%b%\test_wifi" >nul 2>nul
Md "%b%\test_wifi" 2>nul||(Echo 请使用右键管理员身份运行&&Echo Please right-click me and run as administrator!&&Pause >nul&&Exit)
Rd "%b%\test_wifi" >nul 2>nul
:: cpoy
copy /y OCX\*.ocx %windir%\%sys%\
copy /y DLL\*.dll %windir%\%sys%\
:: ocx
regsvr32.exe %windir%\%sys%\COMDLG32.OCX /s
regsvr32.exe %windir%\%sys%\MSCOMCT2.OCX /s
regsvr32.exe %windir%\%sys%\MSCOMCTL.OCX /s
echo. & exit