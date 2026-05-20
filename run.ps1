Set-Location $PSScriptRoot
$python = Join-Path $PSScriptRoot "venv\Scripts\python.exe"

# Chrome CDP check
try {
    Invoke-RestMethod http://localhost:9222/json -TimeoutSec 2 | Out-Null
    Write-Host "Chrome uzhe zapushen." -ForegroundColor Green
} catch {
    Write-Host "Zapuskayu Chrome..." -ForegroundColor Yellow
    Start-Process "chrome.exe" "--remote-debugging-port=9222 --user-data-dir=`"$env:APPDATA\ChromeGastronom`""
    Start-Sleep -Seconds 4
}

# File dialog
Add-Type -AssemblyName System.Windows.Forms
$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Title = "Vyberi ZIP s receptom"
$dialog.Filter = "Arhiv ili dokument (*.zip;*.docx)|*.zip;*.docx"
$dialog.InitialDirectory = [Environment]::GetFolderPath("UserProfile") + "\Downloads"

if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "Fajl ne vybran. Vyhod." -ForegroundColor Red
    Read-Host "Enter"
    exit
}

$recipeFile = $dialog.FileName
Write-Host ""
Write-Host "Fajl: $recipeFile" -ForegroundColor Cyan
Write-Host ""

Write-Host "[1/2] Konvertiruyu recept..." -ForegroundColor Yellow
& $python docx_to_json.py "$recipeFile"
if ($LASTEXITCODE -ne 0) {
    Write-Host "Oshibka pri konvertacii!" -ForegroundColor Red
    Read-Host "Enter"
    exit
}

Write-Host ""
Write-Host "[2/2] Zagruzhayu na sajt..." -ForegroundColor Yellow
& $python upload_recipes.py

Write-Host ""
Write-Host "Gotovo." -ForegroundColor Green
Read-Host "Enter"
