# =====================================================
# Install-Office.ps1
# Probe BEFORE + Verify AFTER install
# Numeric confirmation (0 / 1)
# =====================================================

# ---------- Admin check ----------
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin) {
    Write-Host "Please run PowerShell as Administrator" -ForegroundColor Red
    exit 1
}

Write-Host "Probing existing Office installation..." -ForegroundColor Cyan

# ---------- Detection (BEFORE) ----------
$C2RPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
$OfficeDetected = Test-Path $C2RPath

if ($OfficeDetected) {
    $pre = Get-ItemProperty $C2RPath

    Write-Host ""
    Write-Host "Existing Office detected:" -ForegroundColor Yellow
    Write-Host "  Product : $($pre.ProductReleaseIds)"
    Write-Host "  Version : $($pre.VersionToReport)"
    Write-Host "  Channel : $($pre.UpdateChannel)"
    Write-Host ""
    Write-Host "Choose an option:"
    Write-Host "  [0] Exit"
    Write-Host "  [1] Upgrade / Reinstall Microsoft 365 Apps"
}
else {
    Write-Host ""
    Write-Host "No existing Office detected." -ForegroundColor Green
    Write-Host "Choose an option:"
    Write-Host "  [0] Exit"
    Write-Host "  [1] Install Microsoft 365 Apps"
}

# ---------- User input (numeric validation) ----------
do {
    $choice = Read-Host "Enter your choice (0 or 1)"
} until ($choice -in @("0","1"))

if ($choice -eq "0") {
