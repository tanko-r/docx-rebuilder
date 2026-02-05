# DOCX Rebuilder - Context Menu Uninstaller (PowerShell)

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "DOCX Rebuilder - Context Menu Uninstaller" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Registry paths to remove
$registryPaths = @(
    "HKCU:\SOFTWARE\Classes\.docx\shell\DocxRebuilder",
    "HKCU:\SOFTWARE\Classes\Word.Document.12\shell\DocxRebuilder"
)

foreach ($regPath in $registryPaths) {
    try {
        if (Test-Path $regPath) {
            Remove-Item -Path $regPath -Recurse -Force
            Write-Host "  Removed: $regPath" -ForegroundColor Green
        } else {
            Write-Host "  Not found: $regPath" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "  Error removing $regPath : $_" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "Context menu uninstalled successfully!" -ForegroundColor Green
Write-Host ""

Read-Host "Press Enter to exit"
