# AI Resume Shortlisting System - Startup Script
Write-Host "=" * 70 -ForegroundColor Cyan
Write-Host "  AI-POWERED RESUME SHORTLISTING SYSTEM" -ForegroundColor Green
Write-Host "=" * 70 -ForegroundColor Cyan
Write-Host ""

$projectRoot = "C:\Users\aryan\Downloads\ML\MAJOR Project 2"

# Function to start a service in a new window
function Start-Service {
    param($Name, $Path, $Command, $Color)
    
    Write-Host "Starting $Name..." -ForegroundColor $Color
    Start-Process powershell -ArgumentList "-NoExit", "-Command", "cd '$Path'; Write-Host '=== $Name ===' -ForegroundColor $Color; $Command"
    Start-Sleep -Seconds 2
}

# Check if backend is ready
if (-not (Test-Path "$projectRoot\backend\node_modules")) {
    Write-Host "Installing backend dependencies..." -ForegroundColor Yellow
    cd "$projectRoot\backend"
    npm install
}

# Check if frontend is ready
if (-not (Test-Path "$projectRoot\frontend\node_modules")) {
    Write-Host "Installing frontend dependencies..." -ForegroundColor Yellow
    cd "$projectRoot\frontend"
    npm install
}

# Check Python dependencies
Write-Host "Checking Python ML Pipeline..." -ForegroundColor Yellow
cd "$projectRoot\ml-pipeline"
if (-not (Test-Path "venv")) {
    Write-Host "Creating Python virtual environment..." -ForegroundColor Yellow
    python -m venv venv
}

Write-Host ""
Write-Host "=" * 70 -ForegroundColor Green
Write-Host "STARTING ALL SERVICES..." -ForegroundColor Green
Write-Host "=" * 70 -ForegroundColor Green
Write-Host ""

# Start Backend
Start-Service -Name "BACKEND (Port 5000)" `
              -Path "$projectRoot\backend" `
              -Command "npm start" `
              -Color "Green"

# Start ML Pipeline
Start-Service -Name "ML PIPELINE (Port 5001)" `
              -Path "$projectRoot\ml-pipeline" `
              -Command "python src/app.py" `
              -Color "Magenta"

# Start Frontend
Start-Service -Name "FRONTEND (Port 3000)" `
              -Path "$projectRoot\frontend" `
              -Command "npm start" `
              -Color "Cyan"

Write-Host ""
Write-Host "=" * 70 -ForegroundColor Green
Write-Host "ALL SERVICES STARTED!" -ForegroundColor Green
Write-Host "=" * 70 -ForegroundColor Green
Write-Host ""
Write-Host "Access the application at:" -ForegroundColor Yellow
Write-Host "  Frontend:    http://localhost:3000" -ForegroundColor Cyan
Write-Host "  Backend API: http://localhost:5000" -ForegroundColor Green
Write-Host "  ML Pipeline: http://localhost:5001" -ForegroundColor Magenta
Write-Host ""
Write-Host "Press any key to exit this window (services will keep running)..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
