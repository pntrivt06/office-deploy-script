# Install-Office.ps1
# Detect + Confirm + Install Office

# ---- Admin check ----
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) { Write-Host "Run PowerShell as Administrator" -ForegroundColor Red; exit 1 }

# ---- Detect existing Office ----
$C2RPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"

if (Test-Path $C2RPath) {
    $info = Get-ItemProperty $C2RPath
    Write-Host "Detected installed Office:" -ForegroundColor Cyan
    Write-Host "  Version : $($info.VersionToReport)"
    Write-Host "  Channel : $($info.UpdateChannel)"
    $confirm = Read-Host "Office is already installed. Do you want to REINSTALL / UPGRADE? (Y/N)"
    if ($confirm -notmatch '^[Yy]') {
        Write-Host "Install cancelled by user." -ForegroundColor Cyan
        exit 0
    }
}
else {
    $confirm = Read-Host "No Office detected. Do you want to INSTALL Microsoft 365 Apps? (Y/N)"
    if ($confirm -notmatch '^[Yy]') {
        Write-Host "Install cancelled by user." -ForegroundColor Cyan
        exit 0
    }
}

$Temp = "C:\\Temp\\OfficeInstall"
New-Item -ItemType Directory -Path $Temp -Force | Out-Null

$ODT = "$Temp\\setup.exe"
$Cfg = "$Temp\\config.xml"

Invoke-WebRequest "https://officecdn.microsoft.com/pr/wsus/setup.exe" -OutFile $ODT
Invoke-WebRequest "https://raw.githubusercontent.com/USERNAME/office-deploy-script/main/config.xml" -OutFile $Cfg

Write-Host "Installing Microsoft 365 Apps..." -ForegroundColor Cyan
Start-Process $ODT "/configure `"$Cfg`"" -Wait

# ---- Post-install detection ----
if (Test-Path $C2RPath) {
    $new = Get-ItemProperty $C2RPath
    Write-Host "✅ Office installed successfully:" -ForegroundColor Green
    Write-Host "  Version : $($new.VersionToReport)"
    Write-Host "  Channel : $($new.UpdateChannel)"
}
else {
    Write-Host "❌ Office installation not detected." -ForegroundColor Red
}

exit 0
