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
    Write-Host "Operation cancelled by user." -ForegroundColor Cyan
    exit 0
}

# ---------- Paths ----------
$Temp = "C:\Temp\OfficeInstall"
New-Item -ItemType Directory -Path $Temp -Force | Out-Null

$ODT = "$Temp\setup.exe"
$Cfg = "$Temp\config.xml"

# ---------- Download ODT + config ----------
Write-Host "Downloading Office Deployment Tool..." -ForegroundColor Cyan
Invoke-WebRequest "https://officecdn.microsoft.com/pr/wsus/setup.exe" -OutFile $ODT

Write-Host "Downloading Office configuration..." -ForegroundColor Cyan
Invoke-WebRequest "https://raw.githubusercontent.com/<USERNAME>/office-deploy-script/main/config.xml" -OutFile $Cfg

# ---------- Install ----------
Write-Host "Installing Microsoft 365 Apps..." -ForegroundColor Cyan
Start-Process $ODT "/configure `"$Cfg`"" -Wait

# ---------- Verification (AFTER) ----------
Write-Host "Verifying installation..." -ForegroundColor Cyan
Start-Sleep -Seconds 5

if (Test-Path $C2RPath) {
    $post = Get-ItemProperty $C2RPath

    Write-Host ""
    Write-Host "✅ Office installation verified successfully:" -ForegroundColor Green
    Write-Host "  Product : $($post.ProductReleaseIds)"
    Write-Host "  Version : $($post.VersionToReport)"
    Write-Host "  Channel : $($post.UpdateChannel)"
    Write-Host "  Platform: $($post.Platform)"

    exit 0
}
else {
    Write-Host "❌ Office installation NOT detected after setup." -ForegroundColor Red
    exit 2
}
