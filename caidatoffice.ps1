# =====================================================
# Install-Office.ps1
# Install / Upgrade Microsoft 365 Apps
# FIXED VERSION - Reliable Detection
# SAFE FOR MENU CALL / INTUNE / SCCM
# =====================================================

# ---------- Force TLS 1.2 ----------
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ---------- Admin check ----------
$IsAdmin = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin) {
    Write-Host "Please run PowerShell as Administrator" -ForegroundColor Red
    Read-Host "Press ENTER to exit"
    exit 1
}

Write-Host "Checking existing Office installation..." -ForegroundColor Cyan

# ---------- Detect existing Office ----------
$C2RMain = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
$C2RWOW  = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"

$OfficeDetected = (Test-Path $C2RMain) -or (Test-Path $C2RWOW)

if ($OfficeDetected) {
    $pre = Get-ItemProperty $C2RMain -ErrorAction SilentlyContinue
    if (-not $pre) {
        $pre = Get-ItemProperty $C2RWOW -ErrorAction SilentlyContinue
    }

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
$Log = Join-Path $Temp "OfficeInstall.log"

Write-Host "Downloading Office Deployment Tool..." -ForegroundColor Cyan
Invoke-WebRequest "https://officecdn.microsoft.com/pr/wsus/setup.exe" `
    -OutFile $ODT -ErrorAction Stop

Write-Host "Downloading Office configuration..." -ForegroundColor Cyan
Invoke-WebRequest "https://raw.githubusercontent.com/pntrivt06/office-deploy-script/main/config.xml" `
    -OutFile $Cfg -ErrorAction Stop

if (!(Test-Path $ODT)) {
    Write-Host "ODT not found: $ODT" -ForegroundColor Red
    exit 1
}
if (!(Test-Path $Cfg)) {
    Write-Host "Config not found: $Cfg" -ForegroundColor Red
    exit 1
}

# =====================================================
# INSTALL OFFICE
# =====================================================

Write-Host "Installing Microsoft Office..." -ForegroundColor Cyan
Write-Host "Log: $Log" -ForegroundColor DarkGray

$arguments = @(
    "/configure",
    $Cfg,
    "/log",
    $Log
)

Start-Process -FilePath $ODT -ArgumentList $arguments -Wait

# =====================================================
# VERIFY INSTALL (ROBUST DETECTION)
# =====================================================

Write-Host ""
Write-Host "Verifying Office installation..." -ForegroundColor Cyan

$Detected = $false
for ($i = 1; $i -le 12; $i++) {
    Start-Sleep -Seconds 10

    if (Test-Path $C2RMain -or Test-Path $C2RWOW) {
        $Detected = $true
        break
    }

    Write-Host "  Waiting for Click-to-Run registry ($i/12)..."
}

if ($Detected) {
    $post = Get-ItemProperty $C2RMain -ErrorAction SilentlyContinue
    if (-not $post) {
        $post = Get-ItemProperty $C2RWOW -ErrorAction SilentlyContinue
    }

    Write-Host ""
    Write-Host "✅ Office installed successfully:" -ForegroundColor Green
    Write-Host "  Product : $($post.ProductReleaseIds)"
    Write-Host "  Version : $($post.VersionToReport)"
    Write-Host "  Channel : $($post.UpdateChannel)"
    exit 0
}
else {
    Write-Host ""
    Write-Host "❌ Office installation not detected after waiting." -ForegroundColor Red
    Write-Host "   Office may still be installing in background." -ForegroundColor Yellow
    Write-Host "   Check log: $Log" -ForegroundColor Yellow
    exit 2
}
