# excel_to_csv.ps1 – Right-click → "Run with PowerShell"
Set-Location $PSScriptRoot
$venv = Join-Path $PSScriptRoot ".venv\Scripts\Activate.ps1"
if (Test-Path $venv) { & $venv }
python excel_to_csv.py --interactive
Read-Host "Press Enter to exit"
