# Remove-Office.ps1
# Detect + Confirm + Remove Office

# ---- Admin check ----
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) { Write-Host "Run PowerShell as Administrator" -ForegroundColor Red; exit 1 }

Write-Host "Detecting existing Office installations..." -ForegroundColor Cyan

# ---- Detect MSI Office 2019 ----
$Office2019MSI = Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall |
Get-ItemProperty |
Where-Object { $_.DisplayName -match "Office 2019" -and $_.DisplayName -notmatch "Click-to-Run" }

# ---- Detect Click-to-Run Office ----
$C2R = Test-Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"

if (-not $Office2019MSI -and -not $C2R) {
    Write-Host "✅ No Office installation detected. Nothing to remove." -ForegroundColor Green
    exit 0
}

if ($Office2019MSI) {
    Write-Host "Detected: Office 2019 ProPlus (MSI)" -ForegroundColor Yellow
}

if ($C2R) {
    $c2rInfo = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    Write-Host "Detected: Click-to-Run Office $($c2rInfo.VersionToReport)" -ForegroundColor Yellow
}

$confirm = Read-Host "Do you want to UNINSTALL detected Office products? (Y/N)"
if ($confirm -notmatch '^[Yy]') {
    Write-Host "Uninstall cancelled by user." -ForegroundColor Cyan
    exit 0
}

$Temp = "C:\\Temp\\OfficeRemove"
New-Item -ItemType Directory -Path $Temp -Force | Out-Null

# ---- Remove MSI Office 2019 using SaRA ----
if ($Office2019MSI) {
    Write-Host "Removing Office 2019 MSI using Microsoft SaRA..." -ForegroundColor Cyan
    $SaraUrl = "https://aka.ms/SaRA_OfficeUninstallFromPC"
    $SaraExe = "$Temp\\SaRA.exe"
    Invoke-WebRequest $SaraUrl -OutFile $SaraExe
    Start-Process $SaraExe -ArgumentList "-OfficeVersion 2019 -Quiet" -Wait
}

# ---- Remove Click-to-Run Office ----
if ($C2R) {
    Write-Host "Removing Click-to-Run Office..." -ForegroundColor Cyan
    $ODT = "$Temp\\setup.exe"
    $RemoveXML = "$Temp\\Remove.xml"

@"
<Configuration>
  <Remove All=\"TRUE\">
    <RemoveMSI />
  </Remove>
  <Display Level=\"None\" AcceptEULA=\"TRUE\" />
</Configuration>
"@ | Out-File $RemoveXML -Encoding UTF8

    Invoke-WebRequest "https://officecdn.microsoft.com/pr/wsus/setup.exe" -OutFile $ODT
    Start-Process $ODT "/configure `"$RemoveXML`"" -Wait
}

Write-Host "✅ Office uninstall completed." -ForegroundColor Green
exit 0
