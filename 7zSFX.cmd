@echo off
cd /d "%~dp0"
if not exist "BIN\7z.exe" (
echo ==== ERROR ====
echo Required binary files are missing.
echo.
echo Press any key to exit...
pause >nul
goto :eof
)
if not exist "BIN\7zSFX\*.sfx" (
echo ==== ERROR ====
echo Required SFX module is missing.
echo.
echo Press any key to exit...
pause >nul
goto :eof
)
for /f %%a in ('dir /b /ad .\NDP4* 2^>nul') do if exist "%%a\*.msi" (
set "_src=%%a"
)
if not defined _src (
echo ==== ERROR ====
echo NDP directory is not detected.
echo.
echo Press any key to exit...
pause >nul
goto :eof
)
echo.
echo Create 7z archive . . .
echo.
attrib -A %_src%\* /S /D
BIN\7z.exe a %_src%.7z .\%_src%\* -mqs -mx -m0=BCJ2 -m1=LZMA:d26 -m2=LZMA:d19 -m3=LZMA:d19 -mb0:1 -mb0s1:2 -mb0s2:3 -bso0
echo.
echo Create 7z SFX . . .
echo.
if /i "%_src:~-2%"=="LP" (
copy /b BIN\7zSFXLP\7zSDm.sfx + BIN\7zSFXLP\config.txt + %_src%.7z %_src%-Slim-x86-x64-INTL.exe >nul
) else (
copy /b BIN\7zSFX\7zSD.sfx + BIN\7zSFX\config.txt + %_src%.7z %_src%-Slim-x86-x64-ENU.exe >nul
)
del /f /q %_src%.7z >nul
echo.
echo Done.
echo Press any key to exit.
pause >nul
goto :eof
