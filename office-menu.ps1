# =====================================================
# Office-Menu.ps1 (FIXED)
# Safe master menu for Remove / Install Office
# =====================================================

# ---------- Admin check ----------
$IsAdmin = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin) {
    Write-Host "Please run PowerShell as Administrator" -ForegroundColor Red
    pause
    exit 1
}

# ---------- Resolve script directory safely ----------
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

$RemoveScript  = Join-Path $ScriptDir "Remove-Office.ps1"
$InstallScript = Join-Path $ScriptDir "Install-Office.ps1"

# ---------- Validate script paths ----------
if (!(Test-Path $RemoveScript)) {
    Write-Host "Remove-Office.ps1 not found in $ScriptDir" -ForegroundColor Red
    pause
    exit 1
}

if (!(Test-Path $InstallScript)) {
    Write-Host "Install-Office.ps1 not found in $ScriptDir" -ForegroundColor Red
    pause
    exit 1
}

function Show-Menu {
    Clear-Host
    Write-Host "==============================================="
    Write-Host " Microsoft Office Deployment Menu"
    Write-Host "==============================================="
    Write-Host ""
    Write-Host "  [1] Remove Microsoft Office"
    Write-Host "  [2] Install / Upgrade Microsoft Office"
    Write-Host ""
    Write-Host "  [0] Exit"
    Write-Host ""
}

do {
    Show-Menu
    $choice = Read-Host "Enter your choice (0, 1, or 2)"

    switch ($choice) {
        "1" {
            Write-Host ""
            Write-Host "Launching Office Removal..." -ForegroundColor Yellow
            & "$RemoveScript"
            Write-Host ""
            pause
        }
        "2" {
            Write-Host ""
            Write-Host "Launching Office Installation..." -ForegroundColor Yellow
            & "$InstallScript"
            Write-Host ""
            pause
        }
        "0" {
            Write-Host "Exiting..." -ForegroundColor Cyan
        }
        default {
