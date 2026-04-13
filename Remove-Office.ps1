# =====================================================
# Remove-Office.ps1 (FIXED)
# Safe removal: MSI via SaRA, C2R via ODT (conditional)
# =====================================================

# ---------- Admin check ----------
$IsAdmin = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin) {
    Write-Host "Please run PowerShell as Administrator" -ForegroundColor Red
    exit 1
}

Write-Host "Detecting existing Office installations..." -ForegroundColor Cyan

# ---------- Detect Office 2019 MSI ----------
$Office2019MSI = Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall |
Get-ItemProperty |
Where-Object {
    $_.DisplayName -match "Office 2019" -and
    $_.DisplayName -notmatch "Click-to-Run"
}

# ---------- Detect Click-to-Run ----------
$C2RPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
$HasC2R  = Test-Path $C2RPath

if (-not $Office2019MSI -and -not $HasC2R) {
    Write-Host "✅ No Office detected. Nothing to remove." -ForegroundColor Green
    exit 0
}

if ($Office2019MSI) {
    Write-Host "Detected: Office 2019 ProPlus (MSI)" -ForegroundColor Yellow
}

if ($HasC2R) {
    $c2r = Get-ItemProperty $C2RPath
    Write-Host "Detected: Click-to-Run Office $($c2r.VersionToReport)" -ForegroundColor Yellow
}

# ---------- User confirmation ----------
$confirm = Read-Host "Do you want to UNINSTALL detected Office products? (Y/N)"
if ($confirm -notmatch '^[Yy]') {
    Write-Host "Uninstall cancelled by user." -ForegroundColor Cyan
    exit 0
}

# ---------- Working folder ----------
$Temp = "C:\Temp\OfficeRemove"
New-Item -ItemType Directory -Path $Temp -Force | Out-Null

# =====================================================
# REMOVE OFFICE 2019 MSI (SaRA)
# =====================================================
if ($Office2019MSI) {
    Write-Host "Removing Office 2019 MSI using Microsoft SaRA..." -ForegroundColor Cyan

    $SaraUrl = "https://aka.ms/SaRA_OfficeUninstallFromPC"
    $SaraExe = "$Temp\SaRA.exe"

    Invoke-WebRequest $SaraUrl -OutFile $SaraExe
    Start-Process $SaraExe -ArgumentList "-OfficeVersion 2019 -Quiet" -Wait

    Write-Host "SaRA completed for Office 2019 MSI." -ForegroundColor Green
}
else {
    Write-Host "No Office 2019 MSI detected → skip SaRA" -ForegroundColor Green
}

# =====================================================
# REMOVE CLICK-TO-RUN (ODT) – ONLY IF EXISTS
# =====================================================
if ($HasC2R) {

    Write-Host "Removing Click-to-Run Office using ODT..." -ForegroundColor Cyan

    $ODT = "$Temp\setup.exe"
    $RemoveXML = "$Temp\Remove.xml"

@"
<Configuration>
  <Remove All="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
"@ | Out-File $RemoveXML -Encoding UTF8 -Force

    Invoke-WebRequest "https://officecdn.microsoft.com/pr/wsus/setup.exe" -OutFile $ODT
    Start-Process $ODT "/configure `"$RemoveXML`"" -Wait

    Write-Host "Click-to-Run Office removed." -ForegroundColor Green
}
else {
