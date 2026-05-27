Write-Host "=== WVP Channel Manager ===" -ForegroundColor Cyan
try {
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
} catch {}
try {
    import requests, openpyxl
    Write-Host "Starting..." -ForegroundColor Green
    python main.py
} catch {
    Write-Host "Installing dependencies..." -ForegroundColor Yellow
    pip install -r requirements.txt
    Write-Host "Starting..." -ForegroundColor Green
    python main.py
}
Read-Host "Press Enter to exit"
