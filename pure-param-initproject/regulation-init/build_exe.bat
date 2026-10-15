@echo off
setlocal
pushd "%~dp0"

set "EXIT_CODE=0"

echo Installation/verification de PyInstaller...
py -m pip install pyinstaller
if errorlevel 1 (
	set "EXIT_CODE=1"
	goto :END
)

echo Build executable...
py -m PyInstaller --onefile --name regulation-init src\main.py
if errorlevel 1 (
	set "EXIT_CODE=1"
	goto :END
)

echo.
echo Executable genere dans dist\regulation-init.exe

:END
echo.
if "%EXIT_CODE%"=="0" (
	echo Script termine avec succes.
) else (
	echo Script termine avec erreur. Code=%EXIT_CODE%
)
pause
popd
exit /b %EXIT_CODE%
