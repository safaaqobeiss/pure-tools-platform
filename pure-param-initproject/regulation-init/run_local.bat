@echo off
setlocal
pushd "%~dp0"

set "EXIT_CODE=0"

py src\main.py
set "EXIT_CODE=%errorlevel%"

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
