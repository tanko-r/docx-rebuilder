# DOCX Rebuilder - Windows Context Menu Installer (PowerShell)
# Run this script to install the right-click menu option for .docx files

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "DOCX Rebuilder - Context Menu Installer" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Get script directory
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectDir = Split-Path -Parent $ScriptDir

# Check for Python
try {
    $pythonVersion = python --version 2>&1
    Write-Host "Found: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Python is not installed or not in PATH." -ForegroundColor Red
    Write-Host "Please install Python from https://python.org" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

# Get Python paths
$pythonExe = (Get-Command python).Source
$pythonwExe = $pythonExe -replace "python\.exe$", "pythonw.exe"
$guiScript = Join-Path $ProjectDir "docx_rebuilder\gui.py"

Write-Host ""
Write-Host "Python: $pythonExe"
Write-Host "GUI Script: $guiScript"

# Install dependencies
Write-Host ""
Write-Host "Installing dependencies..." -ForegroundColor Yellow
pip install -r "$ProjectDir\requirements.txt" 2>&1 | Out-Null

# Create registry entries
Write-Host ""
Write-Host "Installing context menu..." -ForegroundColor Yellow

$command = "`"$pythonwExe`" `"$guiScript`" `"%1`""

# Registry paths
$registryPaths = @(
    "HKCU:\SOFTWARE\Classes\.docx\shell\DocxRebuilder",
    "HKCU:\SOFTWARE\Classes\Word.Document.12\shell\DocxRebuilder"
)

foreach ($regPath in $registryPaths) {
    try {
        # Create main key
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }
        Set-ItemProperty -Path $regPath -Name "(Default)" -Value "Rebuild with DOCX Rebuilder"
        Set-ItemProperty -Path $regPath -Name "Icon" -Value "$pythonExe,0"

        # Create command subkey
        $cmdPath = "$regPath\command"
        if (-not (Test-Path $cmdPath)) {
            New-Item -Path $cmdPath -Force | Out-Null
        }
        Set-ItemProperty -Path $cmdPath -Name "(Default)" -Value $command

        Write-Host "  Added: $regPath" -ForegroundColor Green
    } catch {
        Write-Host "  Warning: Could not create $regPath" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "================================================" -ForegroundColor Green
Write-Host "Installation complete!" -ForegroundColor Green
Write-Host ""
Write-Host "Right-click any .docx file and select" -ForegroundColor White
Write-Host "'Rebuild with DOCX Rebuilder' to use." -ForegroundColor White
Write-Host "================================================" -ForegroundColor Green
Write-Host ""

Read-Host "Press Enter to exit"
