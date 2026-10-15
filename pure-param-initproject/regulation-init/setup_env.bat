@echo off
setlocal
pushd "%~dp0"

set "EXIT_CODE=0"

echo [1/2] Verification de Python Launcher...
py --version >nul 2>&1
if errorlevel 1 (
  echo ERREUR: Python Launcher 'py' est introuvable.
  echo Installe Python depuis python.org puis relance ce script.
  set "EXIT_CODE=1"
  goto :END
)

echo [2/2] Installation des dependances...
py -m pip install --upgrade pip
if errorlevel 1 (
  set "EXIT_CODE=1"
  goto :END
)
py -m pip install -r requirements.txt
if errorlevel 1 (
  set "EXIT_CODE=1"
  goto :END
)

echo.
echo OK: environnement pret.
echo Prochaine etape: lance run_local.bat ou run_interactive_gitlab.bat

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
