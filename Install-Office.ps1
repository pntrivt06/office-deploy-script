# =====================================================
# Install-Office.ps1
# Install / Upgrade Microsoft 365 Apps
# SAFE FOR MENU CALL
# =====================================================

# ---------- Admin check ----------
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin) {
    Write-Host "Please run PowerShell as Administrator" -ForegroundColor Red
    Read-Host "Press ENTER to exit"
    exit 1
}

Write-Host "Checking existing Office installation..." -ForegroundColor Cyan

# ---------- Detect existing Office ----------
$C2RPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
$OfficeDetected = Test-Path $C2RPath

if ($OfficeDetected) {
    $pre = Get-ItemProperty $C2RPath
    Write-Host ""
    Write-Host "Existing Office detected:" -ForegroundColor Yellow
    Write-Host "  Product : $($pre.ProductReleaseIds)"
    Write-Host "  Version : $($pre.VersionToReport)"
    Write-Host "  Channel : $($pre.UpdateChannel)"
}
else {
    Write-Host "No existing Office detected." -ForegroundColor Green
}

Write-Host ""
Write-Host "Select an option:"
Write-Host "  [0] Exit"
Write-Host "  [1] Install / Upgrade Microsoft Office"

do {
    $choice = Read-Host "Enter your choice (0 or 1)"
} until ($choice -in @("0","1"))

if ($choice -eq "0") {
    Write-Host "Operation cancelled by user." -ForegroundColor Cyan
    exit 0
}

# =====================================================
# DOWNLOAD ODT + CONFIG
# =====================================================

$Temp = "C:\Temp\OfficeInstall"
New-Item -ItemType Directory -Path $Temp -Force | Out-Null

$ODT = Join-Path $Temp "setup.exe"
$Cfg = Join-Path $Temp "config.xml"

Write-Host "Downloading Office Deployment Tool..." -ForegroundColor Cyan
Invoke-WebRequest "https://officecdn.microsoft.com/pr/wsus/setup.exe" `
    -OutFile $ODT `
    -ErrorAction Stop

Write-Host "Downloading Office configuration..." -ForegroundColor Cyan
Invoke-WebRequest "https://raw.githubusercontent.com/pntrivt06/office-deploy-script/main/config.xml" `
    -OutFile $Cfg `
    -ErrorAction Stop

# ---------- Validate paths ----------
if (!(Test-Path $ODT)) {
    Write-Host "ODT not found: $ODT" -ForegroundColor Red
    exit 1
}
if (!(Test-Path $Cfg)) {
    Write-Host "Config not found: $Cfg" -ForegroundColor Red
    exit 1
}

# =====================================================
# INSTALL OFFICE (NO BACKTICK, NO QUOTE ISSUE)
# =====================================================

Write-Host "Installing Microsoft Office..." -ForegroundColor Cyan

$arguments = @(
    "/configure",
    $Cfg
)

Start-Process -FilePath $ODT -ArgumentList $arguments -Wait

# =====================================================
# VERIFY INSTALL
# =====================================================

Start-Sleep -Seconds 5

if (Test-Path $C2RPath) {
    $post = Get-ItemProperty $C2RPath
    Write-Host ""
    Write-Host "✅ Office installed successfully:" -ForegroundColor Green
    Write-Host "  Product : $($post.ProductReleaseIds)"
    Write-Host "  Version : $($post.VersionToReport)"
    Write-Host "  Channel : $($post.UpdateChannel)"
    exit 0
}
else {
    Write-Host "❌ Office installation not detected." -ForegroundColor Red
    exit 2
}
