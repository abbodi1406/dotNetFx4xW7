@setlocal DisableDelayedExpansion
@echo off
:: Rebuild/Repack files inside netfx_Full.mzz
set BuildMzz=1

:: Create LZX high-compressed netfx_Full.mzz
set CompressMzz=0

:: Show slipstreamed patches in "Control Panel\Programs and Features\Installed Updates"
set ShowMsp=1

:: Create single-architecture repack
:: set to x64 or x86
:: leave blank to create both
set Single=

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
for /f %%i in ('dir /b netfx_Patch*.msp') do (
call set /a msp+=1
)

set do86=1
set do64=1
set solo=0
set rpck=x64/x86
if /i "%Single%"=="x64" (
set do86=0
set solo=1
set rpck=x64-only
)
if /i "%Single%"=="x86" (
set do64=0
set solo=1
set rpck=x86-only
)
echo.
echo Create %rpck% repack . . .

setlocal EnableExtensions EnableDelayedExpansion
if %msp%==0 (
set ShowMsp=0
goto :extract
)
set nkb=
set _ci=0
if %do86% equ 1 for /f %%i in ('dir /b /od *NDP*x86*.exe') do if /i not "%%i"=="%ndpack%" (
set /a _ci+=1
set "x86exe!_ci!=%%i"
)
if %do86% equ 1 for /f %%i in ('dir /b /od netfx_Patch_x86.msp') do (
set /a _ci+=1
set "x86exe!_ci!=%%i"
for /f %%# in ('cscript //NoLogo BIN\NDP\mspkb.vbs %%i') do set "nkb=%%#"
)
set _cx=0
if %do64% equ 1 for /f %%i in ('dir /b /od *NDP*x64*.exe') do if /i not "%%i"=="%ndpack%" (
set /a _cx+=1
set "x64exe!_cx!=%%i"
)
if %do64% equ 1 for /f %%i in ('dir /b /od netfx_Patch_x64.msp') do (
set /a _cx+=1
set "x64exe!_cx!=%%i"
for /f %%# in ('cscript //NoLogo BIN\NDP\mspkb.vbs %%i') do set "nkb=%%#"
)
if %solo% equ 0 if %_ci% neq %_cx% (
echo ==== ERROR ====
echo Update files count is not equal for both architectures.
echo.
echo Press any key to exit...
pause >nul
goto :eof
)
if not defined nkb goto :extract
if /i "%nkb:~-2,1%"=="v" set "nkb=%nkb:~0,-2%"

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
if %do86% equ 1 for /l %%i in (1,1,%_ci%) do (
dir /b !x86exe%%i! | findstr /i \.exe >nul && (
  for /f "tokens=2 delims=-" %%a in ('dir /b !x86exe%%i!') do (
    set "x86kb%%i=%%a"
    set "x86msp%%i=%ndpver%-%%a-x86.msp"
    for %%b in (K,B) do set x86kb%%i=!x86kb%%i:%%b=%%b!
    )
  ) || (
    set "x86kb%%i=%nkb%"
    set "x86msp%%i=%ndpver%-%nkb%-x86.msp"
  )
)
if %do64% equ 1 for /l %%i in (1,1,%_cx%) do (
dir /b !x64exe%%i! | findstr /i \.exe >nul && (
  for /f "tokens=2 delims=-" %%a in ('dir /b !x64exe%%i!') do (
    set "x64kb%%i=%%a"
    set "x64msp%%i=%ndpver%-%%a-x64.msp"
    for %%b in (K,B) do set x64kb%%i=!x64kb%%i:%%b=%%b!
    )
  ) || (
    set "x64kb%%i=%nkb%"
    set "x64msp%%i=%ndpver%-%nkb%-x64.msp"
  )
)
if %do86% equ 1 for /l %%i in (1,1,%_ci%) do (
  if exist "*!x86kb%%i!*x86*.exe" (
  BIN\7z.exe e !x86exe%%i! -o%ndpver% *.msp >nul
  for /f %%a in ('dir /b %ndpver%\*!x86kb%%i!.msp') do ren %ndpver%\%%a !x86msp%%i!
  ) else (
  copy /y !x86exe%%i! %ndpver%\!x86msp%%i! >nul
  )
)
if %do64% equ 1 for /l %%i in (1,1,%_cx%) do (
  if exist "*!x64kb%%i!*x64*.exe" (
  BIN\7z.exe e !x64exe%%i! -o%ndpver% *.msp >nul
  for /f %%a in ('dir /b %ndpver%\*!x64kb%%i!.msp') do ren %ndpver%\%%a !x64msp%%i!
  ) else (
  copy /y !x64exe%%i! %ndpver%\!x64msp%%i! >nul
  )
)

:slim
echo.
echo Slim MSI database . . .
xcopy /criy BIN\NDP\* %ndpver%\ >nul
cd %ndpver%
if %do86% equ 1 cscript //B slim.vbs temp\netfx_Full_x86.msi
if %do64% equ 1 cscript //B slim.vbs temp\netfx_Full_x64.msi
if %BuildMzz%==1 if %msp%==0 goto :noop
for /f "tokens=2* delims== " %%a in ('cscript //NoLogo WiSumInf.vbs temp\netfx_Full_x86.msi ^| findstr /i Subject') do set name="%%b"
for /f "tokens=2* delims== " %%a in ('cscript //NoLogo WiSumInf.vbs temp\netfx_Full_x86.msi ^| findstr /i Comments') do set desc="%%b"
for /f "tokens=2* delims== " %%a in ('cscript //NoLogo WiSumInf.vbs temp\netfx_Full_x86.msi ^| findstr /i Revision') do set "guid86=%%b"
for /f "tokens=2* delims== " %%a in ('cscript //NoLogo WiSumInf.vbs temp\netfx_Full_x64.msi ^| findstr /i Revision') do set "guid64=%%b"

if %do64% equ 0 goto :no64
echo.
echo Create administrative install ^(x64^) . . .
start /wait msiexec /a temp\netfx_Full_x64.msi TARGETDIR=%cd% /quiet
if not %msp%==0 (
  for /l %%i in (1,1,%_cx%) do (
  start /wait msiexec /a %cd%\netfx_Full_x64.msi PATCH=%cd%\!x64msp%%i! TARGETDIR=%cd% /quiet
  )
)

:no64
if %do86% equ 0 goto :no86
echo.
echo Create administrative install ^(x86^) . . .
start /wait msiexec /a temp\netfx_Full_x86.msi TARGETDIR=%cd% /quiet
if not %msp%==0 (
  for /l %%i in (1,1,%_ci%) do (
  start /wait msiexec /a %cd%\netfx_Full_x86.msi PATCH=%cd%\!x86msp%%i! TARGETDIR=%cd% /quiet
  )
)

:no86
rd /s /q temp\

echo.
echo Adjust MSI properties . . .
if exist "netfx_Full_x86.msi" (
if not %msp%==0 cscript //B WiFilVer.vbs netfx_Full_x86.msi /u
cscript //B WiSumInf.vbs netfx_Full_x86.msi Subject=%name% Comments=%desc% Revision=%guid86% Words=4
)
if exist "netfx_Full_x64.msi" (
if not %msp%==0 cscript //B WiFilVer.vbs netfx_Full_x64.msi /u
cscript //B WiSumInf.vbs netfx_Full_x64.msi Subject=%name% Comments=%desc% Revision=%guid64% Words=4
)

if %BuildMzz%==0 goto :skip
if not exist "%_wix%\heat.exe" goto :skip
if %CompressMzz%==0 (set _dcl=none) else (set _dcl=high)
echo.
echo Rebuild netfx_Full.mzz . . .
if exist "netfx_Full_x64.msi" (
cscript //B WiMakCab.vbs netfx_Full_x64.msi netfx
) else (
cscript //B WiMakCab.vbs netfx_Full_x86.msi netfx
)
mkdir SourceDir
cd SourceDir
for /f "tokens=* delims=" %%i in (..\netfx.ddf) do move /y>nul %%i
"%_wix%\heat.exe" dir . -nologo -g1 -gg -suid -scom -sreg -srd -sfrag -svb6 -indent 1 -t ..\netfx.xsl -template product -out ..\netfx.wxs
cd ..
"%_wix%\candle.exe" netfx.wxs -nologo -sw1074 >nul
"%_wix%\light.exe" netfx.wixobj -nologo -spdb -sice:ICE21 -dcl:%_dcl% >nul
ren product.cab netfx_Full.mzz
rd /s /q ProgramFilesFolder\ Windows\ SourceDir\
if exist "netfx_Full_x86.msi" (
cscript //B WiSumInf.vbs netfx_Full_x86.msi Words=0
)
if exist "netfx_Full_x64.msi" (
cscript //B WiSumInf.vbs netfx_Full_x64.msi Words=0
)

:skip
if %ShowMsp%==0 goto :end
if %_ci% gtr 0 (
set _cz=%_ci%
set _ca=x86
) else (
set _cz=%_cx%
set _ca=x64
)
echo.
echo Show slipstreamed updates . . .
for /l %%i in (1,1,%_cz%) do (
  (
  echo.  QueryDatabase^("UPDATE `Component` SET Condition = '!%_ca%kb%%i!.Classification' WHERE `Component` = '!%_ca%kb%%i!.ARP.Add'"^)
  echo.  QueryDatabase^("UPDATE `Component` SET Condition = '!%_ca%kb%%i!.AllowRemoval <> 1' WHERE `Component` = '!%_ca%kb%%i!.ARP.NoRemove'"^)
  echo.  QueryDatabase^("UPDATE `Component` SET Condition = '!%_ca%kb%%i!v2.Classification' WHERE `Component` = '!%_ca%kb%%i!v2.ARP.Add'"^)
  echo.  QueryDatabase^("UPDATE `Component` SET Condition = '!%_ca%kb%%i!v2.AllowRemoval <> 1' WHERE `Component` = '!%_ca%kb%%i!v2.ARP.NoRemove'"^)
  echo.  QueryDatabase^("UPDATE `Component` SET Condition = '!%_ca%kb%%i!v3.Classification' WHERE `Component` = '!%_ca%kb%%i!v3.ARP.Add'"^)
  echo.  QueryDatabase^("UPDATE `Component` SET Condition = '!%_ca%kb%%i!v3.AllowRemoval <> 1' WHERE `Component` = '!%_ca%kb%%i!v3.ARP.NoRemove'"^)
  echo.  QueryDatabase^("UPDATE `Component` SET Condition = '!%_ca%kb%%i!v4.Classification' WHERE `Component` = '!%_ca%kb%%i!v4.ARP.Add'"^)
  echo.  QueryDatabase^("UPDATE `Component` SET Condition = '!%_ca%kb%%i!v4.AllowRemoval <> 1' WHERE `Component` = '!%_ca%kb%%i!v4.ARP.NoRemove'"^)
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
