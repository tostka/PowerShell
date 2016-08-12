::: bootalert.cmd
:::	===remote WIN2K8 system===
::: on 2008R2 new /rp params For Task Scheduler 2.0 tasks:
::: /ru "NT AUTHORITY\LOCALSERVICE"
::: /ru "NT AUTHORITY\NETWORKSERVICE"

::: 1) schtasks /create /s USEA-fdhubcas2 /tn "bootalert" /tr e:\scripts\bootalert.ps1 /sc onstart /ru "NT AUTHORITY\NETWORKSERVICE" 
::: 2) launch it to verify function:
:::		schtasks /run /tn bootalert /s USEA-fdhubcas2
::: 3) Verify that you receive an email.

@ echo off

::: 7:50 AM 5/8/2013 need to do some path detection since we're running this on diff dirs
::: PERFORM STANDARD TESTS
::: DERIVE SCRIPTPATH DRIVE LETTER
SET SCRIPTPATH=%~d0
::: DERIVE & ADD SCRIPTPATH PATH
SET SCRIPTPATH=%SCRIPTPATH%%~p0
ECHO SCRIPTPATH=%SCRIPTPATH%
::: echos as 'c:\Scripts\'
set PSExec=@%WINDIR%\system32\windowspowershell\v1.0\powershell.exe
::: DEFAULT install locations, correct to actual installs loc
set Ex10BinPath=%ProgramFiles%\Microsoft\Exchange Server\V14\bin\
set Ex7BinPath=%ProgramFiles%\Microsoft\Exchange Server\bin\
::: TargetScript
set targPS1=%SCRIPTPATH%bootalert.ps1
ECHO targPS1=%targPS1%
:RERUNIT

::: ***EMS10 Powershell call***
::IF EXIST %targPS1% (@%PSExec% -noexit -command "& {. '%Ex10BinPath%RemoteExchange.ps1'; Connect-ExchangeServer -auto ; %targPS1% }") ELSE ((ECHO MISSING %targPS1%) & (PAUSE))

::: ***EMS7 Powershell call*** 
::IF EXIST %targPS1% (@%PSExec% -PSConsoleFile "%Ex7BinPath%\exshell.psc1" -noexit -command ". '%Ex7BinPath%Exchange.ps1' ; %targPS1%") ELSE ((ECHO MISSING %targPS1%) & (PAUSE))

::: ***STOCK POWERSHELL call (or EMS7, if call the snapin in the target ps1)***
::IF EXIST %targPS1% (@%PSExec% -command "& {%targPS1% }") ELSE ((ECHO MISSING %targPS1%) & (PAUSE))
IF EXIST %targPS1% (@%PSExec% -command "& {%targPS1% }") ELSE ((ECHO MISSING %targPS1%) & (PAUSE))

::: ***STOCK POWERSHELL ON WIN7 UNDER ExecutionPolicy RESTRICTIONS***
::: This completely bypasses ExecutionPolicy restrictions 
:::IF EXIST %targPS1% (@%PSExec% -ExecutionPolicy Bypass -NoLogo -NoProfile -File %targPS1%) ELSE ((ECHO MISSING %targPS1%) & (PAUSE))
::: works to run script from java

:::pause if needed to sustain results, or for bootalert etc
::pause

::: ADDED TO MOVE PAST CRASHES
::GOTO :RERUNIT 

