ECHO OFF

REM locate the ipy.exe correctly on 64bit and 32bit OSes 
SET IPVERSION=IronPython 2.7
SET IPEXE=%IPVERSION%\IPY.exe
SET IPEXEFULLPATH="%ProgramFiles(x86)%\%IPEXE%"
IF NOT EXIST %IPEXEFULLPATH% SET IPEXEFULLPATH="%ProgramFiles%\%IPEXE%"

REM Launch the IronPython interactive shell 
%IPEXEFULLPATH% -D -X:TabCompletion -X:ColorfulConsole -X:AutoIndent -i %~dp0visio.py
