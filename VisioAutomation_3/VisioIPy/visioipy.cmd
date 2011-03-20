ECHO OFF

REM ----------------------------------------
REM Prepare the environment 

REM Setting env var IRONPYTHONSTARTUP will make ip.exe launch and start this script automatically - very useful for interactive sessions
REM because it saves the user from having to always import that module first
SET IRONPYTHONSTARTUP=%~dp0visioipy.py

REM Now locate the ipy.exe correctly on 64bit and 32bit OSes 
SET IPVERSION=IronPython 2.7
SET IPEXE=%IPVERSION%\IPY.exe
SET IPEXEFULLPATH="%ProgramFiles(x86)%\%IPEXE%"
IF NOT EXIST %IPEXEFULLPATH% SET IPEXEFULLPATH="%ProgramFiles%\%IPEXE%"

REM ----------------------------------------
REM Launch the IronPython interactive shell 

%IPEXEFULLPATH% -D -X:TabCompletion -X:ColorfulConsole

REM ----------------------------------------
REM Cleanup the environment

SET IRONPYTHONSTARTUP=
