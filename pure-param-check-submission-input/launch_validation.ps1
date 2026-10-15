# launch_validation.ps1 – Right-click → "Run with PowerShell"
Set-Location $PSScriptRoot
$venv = Join-Path $PSScriptRoot ".venv\Scripts\Activate.ps1"
if (Test-Path $venv) { & $venv }
python launch_validation.py
Read-Host "Press Enter to exit"
