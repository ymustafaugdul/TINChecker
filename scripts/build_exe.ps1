$ErrorActionPreference = "Stop"

python -m PyInstaller --clean .\tin_checker.spec

Write-Host ""
Write-Host "Build tamamlandı: dist\TINChecker.exe"
