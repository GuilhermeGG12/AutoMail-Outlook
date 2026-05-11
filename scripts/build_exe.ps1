$ErrorActionPreference = "Stop"

$Python = Join-Path $PSScriptRoot "..\.venv\Scripts\python.exe"
if (-not (Test-Path $Python)) {
  $Python = "python"
}

& $Python -m PyInstaller `
  --noconfirm `
  --windowed `
  --name "OutlookMailMergeAssistant" `
  --paths "src" `
  --add-data "src\mailmerge_assistant\assets;mailmerge_assistant\assets" `
  "src\mailmerge_assistant\app.py"
