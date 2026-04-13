# =====================================================
# Install-Office.ps1 (FIX 0-2048)
# Pre-cleanup + Install + Verify
# =====================================================

# ---------- Admin check ----------
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin) {
    Write-Host "Please run PowerShell as Administrator" -ForegroundColor Red
    exit 1
}

Write-Host "Preparing system for Office installation..." -ForegroundColor Cyan

# =====================================================
# PRE-INSTALL CLEANUP (CRITICAL)
# =====================================================

# Kill running Office processes
Get-Process winword,excel,powerpnt,outlook,officeclicktorun -ErrorAction SilentlyContinue |
Stop-Process -Force

# Stop Click-to-Run service if exists
Get-Service ClickToRunSvc -ErrorAction SilentlyContinue |
Where-Object {$_.Status -ne "Stopped"} |
Stop-Service -Force

# Wait for MSI locks to release
Start-Sleep -Seconds 10

# =====================================================
# INSTALL USING ODT
# =====================================================

$Temp = "C:\Temp\OfficeInstall"
New-Item -ItemType Directory -Path $Temp -Force | Out-Null

$ODT = Join-Path $Temp "setup.exe"
$Cfg = Join-Path $Temp "config.xml"

Write-Host "Downloading Office Deployment Tool..." -ForegroundColor Cyan
Invoke-WebRequest "https://officecdn.microsoft.com/pr/wsus/setup.exe" -OutFile $ODT -ErrorAction Stop

Write-Host "Downloading Office configuration..." -ForegroundColor Cyan
Invoke-WebRequest "https://raw.githubusercontent.com/pntrivt06/office-deploy-script/main/config.xml" -OutFile $Cfg -ErrorAction Stop

Write-Host "Installing Microsoft 365 Apps..." -ForegroundColor Cyan
Start-Process -FilePath $ODT `
    -ArgumentList "/configure `"$Cfg`"" `
    -Wait

# =====================================================
# VERIFICATION
# =====================================================

$C2RPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"

Start-Sleep -Seconds 5

if (Test-Path $C2RPath) {
    $post = Get-ItemProperty $C2RPath
    Write-Host "✅ Office installation verified:" -ForegroundColor Green
    Write-Host "  Product : $($post.ProductReleaseIds)"
    Write-Host "  Version : $($post.VersionToReport)"
    Write-Host "  Channel : $($post.UpdateChannel)"
    exit 0
}
else {
    Write-Host "❌ Office installation failed – registry not found." -ForegroundColor Red
    exit 2
}
