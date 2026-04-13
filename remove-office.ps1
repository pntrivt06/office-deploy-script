# =====================================================
# Remove-Office.ps1
# MENU 0/1 – Full Office removal
# + BLOCK Outlook AppX from reinstalling
# =====================================================

# ---------- Admin check ----------
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin) {
    Write-Host "Please run PowerShell as Administrator" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Detecting installed Microsoft Office components..." -ForegroundColor Cyan

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
    Write-Host "✅ No Microsoft Office detected." -ForegroundColor Green
}

if ($Office2019MSI) {
    Write-Host "Detected: Office 2019 ProPlus (MSI)" -ForegroundColor Yellow
}

if ($HasC2R) {
    $c2r = Get-ItemProperty $C2RPath
    Write-Host "Detected: Click-to-Run Office $($c2r.VersionToReport)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Select an option:"
Write-Host "  [0] Exit (do nothing)"
Write-Host "  [1] REMOVE Microsoft Office completely"

# ---------- Numeric menu ----------
do {
    $choice = Read-Host "Enter your choice (0 or 1)"
} until ($choice -in @("0","1"))

if ($choice -eq "0") {
    Write-Host "Operation cancelled by user." -ForegroundColor Cyan
    exit 0
}

# =====================================================
# FULL REMOVAL
# =====================================================

Write-Host ""
Write-Host "Starting FULL Office removal..." -ForegroundColor Red

# ---------- Kill Office processes ----------
Get-Process winword,excel,powerpnt,outlook,lync,skype,officeclicktorun -ErrorAction SilentlyContinue |
Stop-Process -Force

# =====================================================
# REMOVE OUTLOOK APPX ONLY (KEEP Mail & Calendar)
# =====================================================
Write-Host ""
Write-Host "Checking for Outlook AppX (New Outlook)..." -ForegroundColor Cyan

$OutlookAppx = Get-AppxPackage |
Where-Object {
    $_.Name -eq "Microsoft.OutlookForWindows" -or
    $_.Name -like "*Outlook*"
}

if ($OutlookAppx) {
    foreach ($app in $OutlookAppx) {
        Write-Host "Removing Outlook AppX: $($app.Name)" -ForegroundColor Yellow
        Remove-AppxPackage -Package $app.PackageFullName -ErrorAction SilentlyContinue
    }
    Write-Host "✅ Outlook AppX removed." -ForegroundColor Green
}
else {
    Write-Host "No Outlook AppX detected." -ForegroundColor Green
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

    Write-Host "✅ Office 2019 MSI removed." -ForegroundColor Green
}

# =====================================================
# REMOVE CLICK-TO-RUN OFFICE (ALL APPS)
# =====================================================
if ($HasC2R) {

    Write-Host "Removing Click-to-Run Office (all Win32 apps)..." -ForegroundColor Cyan

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

    Write-Host "✅ Click-to-Run Office removed." -ForegroundColor Green
}

# =====================================================
# BLOCK OUTLOOK APPX FROM REINSTALLING
# =====================================================
Write-Host ""
Write-Host "Blocking Outlook AppX from reinstalling..." -ForegroundColor Cyan

# --- Block New Outlook migration (Microsoft supported) ---
$regOutlook = "HKLM:\SOFTWARE\Microsoft\Office\Outlook\Preferences"
if (!(Test-Path $regOutlook)) {
    New-Item -Path $regOutlook -Force | Out-Null
}

Set-ItemProperty -Path $regOutlook `
    -Name "NewOutlookMigrationUserSetting" `
    -Type DWord `
    -Value 0

# --- Disable Microsoft Store auto-update ---
$regStore = "HKLM:\SOFTWARE\Policies\Microsoft\WindowsStore"
if (!(Test-Path $regStore)) {
    New-Item -Path $regStore -Force | Out-Null
}

Set-ItemProperty -Path $regStore `
    -Name "AutoDownload" `
    -Type DWord `
    -Value 2

Write-Host "✅ Outlook AppX blocked from reinstalling." -ForegroundColor Green

Write-Host ""
Write-Host "✅ OFFICE REMOVAL COMPLETED SUCCESSFULLY." -ForegroundColor Green
Write-Host "⚠️ Recommended: REBOOT the computer before installing Office again." -ForegroundColor Yellow
exit 0
``
