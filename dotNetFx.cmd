@echo off
:: Rebuild/Repack files inside netfx_Full.mzz
set BuildMzz=1

:: Show slipstreamed patches in "Control Panel\Programs and Features\Installed Updates"
set ShowMsp=1

%windir%\system32\reg.exe query "HKU\S-1-5-19" >nul 2>&1 || (
echo ==== ERROR ====
echo This script require administrator privileges.
echo To do so, right click on this script and select 'Run as administrator'
echo.
echo Press any key to exit...
pause >nul
goto :eof
)
cd /d "%~dp0"
if not exist "BIN\7z.exe" (
echo ==== ERROR ====
echo Required binary files are missing.
echo.
echo Press any key to exit...
pause >nul
goto :eof
)
if not exist "BIN\NDP\*.vbs" (
echo ==== ERROR ====
echo Required work files are missing.
echo.
echo Press any key to exit...
pause >nul
goto :eof
)
if not exist "*NDP*ENU*.exe" (
echo ==== ERROR ====
echo NDP*-x86-x64-ENU.exe file is not detected.
echo.
echo Press any key to exit...
pause >nul
goto :eof
)

for /f %%i in ('dir /b *NDP*-ENU*.exe') do set "ndpack=%%i"
for /f "tokens=1 delims=-" %%i in ('dir /b %ndpack%') do set "ndpver=%%i"

set msp=0
for /f %%i in ('dir /b *NDP*.exe') do (
if /i not "%%i"=="%ndpack%" call set /a msp+=1
)
setlocal EnableExtensions EnableDelayedExpansion
if %msp%==0 (
set ShowMsp=0
goto :extract
)
set _ci=0
for /f %%i in ('dir /b /od *NDP*x86*.exe') do if /i not "%%i"=="%ndpack%" (
set /a _ci+=1
set "x86exe!_ci!=%%i"
)
set _cx=0
for /f %%i in ('dir /b /od *NDP*x64*.exe') do if /i not "%%i"=="%ndpack%" (
set /a _cx+=1
set "x64exe!_cx!=%%i"
)
if not %_ci%==%_cx% (
echo ==== ERROR ====
echo Update files count is not equal for both architectures.
echo.
echo Press any key to exit...
pause >nul
goto :eof
)

:extract
echo.
echo Extract files . . .
set "_wix=%~dp0BIN"
set /a rnd=%random%
if exist "%ndpver%" ren %ndpver% %ndpver%_%rnd%
BIN\7z.exe e %ndpack% -o%ndpver%\temp netfx_Full_x64.msi netfx_Full_x86.msi netfx_Full.mzz >nul
if %msp%==0 (
goto :slim
)
for /l %%i in (1,1,%_ci%) do (
for /f "tokens=2 delims=-" %%a in ('dir /b !x86exe%%i!') do (
  set "x86kb%%i=%%a"
  set "x86msp%%i=%ndpver%-%%a-x86.msp"
  for %%b in (K,B) do set x86kb%%i=!x86kb%%i:%%b=%%b!
  )
)
for /l %%i in (1,1,%_cx%) do (
for /f "tokens=2 delims=-" %%a in ('dir /b !x64exe%%i!') do (
  set "x64kb%%i=%%a"
  set "x64msp%%i=%ndpver%-%%a-x64.msp"
  for %%b in (K,B) do set x64kb%%i=!x64kb%%i:%%b=%%b!
  )
)
for /l %%i in (1,1,%_ci%) do (
BIN\7z.exe e !x86exe%%i! -o%ndpver% *.msp >nul
for /f %%a in ('dir /b %ndpver%\*!x86kb%%i!.msp') do ren %ndpver%\%%a !x86msp%%i!
)
for /l %%i in (1,1,%_cx%) do (
BIN\7z.exe e !x64exe%%i! -o%ndpver% *.msp >nul
for /f %%a in ('dir /b %ndpver%\*!x64kb%%i!.msp') do ren %ndpver%\%%a !x64msp%%i!
)

:slim
echo.
echo Slim MSI database . . .
xcopy /criy BIN\NDP\* %ndpver%\ >nul
cd %ndpver%
cscript //B slim.vbs temp\netfx_Full_x86.msi
cscript //B slim.vbs temp\netfx_Full_x64.msi
if %BuildMzz%==1 if %msp%==0 goto :noop
for /f "tokens=2* delims== " %%a in ('cscript //NoLogo WiSumInf.vbs temp\netfx_Full_x86.msi ^| findstr /i Subject') do set name="%%b"
for /f "tokens=2* delims== " %%a in ('cscript //NoLogo WiSumInf.vbs temp\netfx_Full_x86.msi ^| findstr /i Comments') do set desc="%%b"
for /f "tokens=2* delims== " %%a in ('cscript //NoLogo WiSumInf.vbs temp\netfx_Full_x86.msi ^| findstr /i Revision') do set "guid86=%%b"
for /f "tokens=2* delims== " %%a in ('cscript //NoLogo WiSumInf.vbs temp\netfx_Full_x64.msi ^| findstr /i Revision') do set "guid64=%%b"

echo.
echo Create administrative install ^(x64^) . . .
start /wait msiexec /a temp\netfx_Full_x64.msi TARGETDIR=%cd% /quiet
if not %msp%==0 (
  for /l %%i in (1,1,%_cx%) do (
  start /wait msiexec /a %cd%\netfx_Full_x64.msi PATCH=%cd%\!x64msp%%i! TARGETDIR=%cd% /quiet
  )
)

echo.
echo Create administrative install ^(x86^) . . .
start /wait msiexec /a temp\netfx_Full_x86.msi TARGETDIR=%cd% /quiet
if not %msp%==0 (
  for /l %%i in (1,1,%_ci%) do (
  start /wait msiexec /a %cd%\netfx_Full_x86.msi PATCH=%cd%\!x86msp%%i! TARGETDIR=%cd% /quiet
  )
)

rd /s /q temp\

echo.
echo Adjust MSI properties . . .
if not %msp%==0 (
cscript //B WiFilVer.vbs netfx_Full_x86.msi /u
cscript //B WiFilVer.vbs netfx_Full_x64.msi /u
)
cscript //B WiSumInf.vbs netfx_Full_x86.msi Subject=%name% Comments=%desc% Revision=%guid86% Words=4
cscript //B WiSumInf.vbs netfx_Full_x64.msi Subject=%name% Comments=%desc% Revision=%guid64% Words=4

if %BuildMzz%==0 goto :skip
if not exist "%_wix%\heat.exe" goto :skip
echo.
echo Rebuild netfx_Full.mzz . . .
cscript //B WiMakCab.vbs netfx_Full_x64.msi netfx
mkdir SourceDir
cd SourceDir
for /f "tokens=* delims=" %%i in (..\netfx.ddf) do move /y>nul %%i
"%_wix%\heat.exe" dir . -nologo -g1 -gg -suid -scom -sreg -srd -sfrag -svb6 -indent 1 -t ..\netfx.xsl -template product -out ..\netfx.wxs
cd ..
"%_wix%\candle.exe" netfx.wxs -nologo -sw1074 >nul
"%_wix%\light.exe" netfx.wixobj -nologo -spdb -sice:ICE21 -dcl:none >nul
ren product.cab netfx_Full.mzz
rd /s /q ProgramFilesFolder\ Windows\ SourceDir\
cscript //B WiSumInf.vbs netfx_Full_x86.msi Words=0
cscript //B WiSumInf.vbs netfx_Full_x64.msi Words=0

:skip
if %ShowMsp%==0 goto :end
echo.
echo Show slipstreamed updates . . .
for /l %%i in (1,1,%_ci%) do (
  (
  echo.  QueryDatabase^("UPDATE `Component` SET Condition = '!x86kb%%i!.Classification' WHERE `Component` = '!x86kb%%i!.ARP.Add'"^)
  echo.  QueryDatabase^("UPDATE `Component` SET Condition = '!x86kb%%i!.AllowRemoval <> 1' WHERE `Component` = '!x86kb%%i!.ARP.NoRemove'"^)
  echo.  QueryDatabase^("UPDATE `Component` SET Condition = '!x86kb%%i!v2.Classification' WHERE `Component` = '!x86kb%%i!v2.ARP.Add'"^)
  echo.  QueryDatabase^("UPDATE `Component` SET Condition = '!x86kb%%i!v2.AllowRemoval <> 1' WHERE `Component` = '!x86kb%%i!v2.ARP.NoRemove'"^)
  echo.  QueryDatabase^("UPDATE `Component` SET Condition = '!x86kb%%i!v3.Classification' WHERE `Component` = '!x86kb%%i!v3.ARP.Add'"^)
  echo.  QueryDatabase^("UPDATE `Component` SET Condition = '!x86kb%%i!v3.AllowRemoval <> 1' WHERE `Component` = '!x86kb%%i!v3.ARP.NoRemove'"^)
  echo.  QueryDatabase^("UPDATE `Component` SET Condition = '!x86kb%%i!v4.Classification' WHERE `Component` = '!x86kb%%i!v4.ARP.Add'"^)
  echo.  QueryDatabase^("UPDATE `Component` SET Condition = '!x86kb%%i!v4.AllowRemoval <> 1' WHERE `Component` = '!x86kb%%i!v4.ARP.NoRemove'"^)
  )>>showmsp.vbs
)
  (
  echo.  Set db ^= nothing
  echo End Function
  )>>showmsp.vbs
cscript //B showmsp.vbs

:end
echo.
echo Cleanup . . .
cscript //B esu.vbs
del /f /q netfx.* *.vbs *.msp *.ico >nul
echo.
echo Done.
echo Press any key to exit.
pause >nul
goto :eof

:noop
del /f /q netfx.* *.vbs *.ico
robocopy temp\ .\ /MOVE >nul
echo.
echo Done.
echo Press any key to exit.
pause >nul
goto :eof
