# ===============================
# Deploy.ps1 - Office Deployment
# Includes SaRA for MSI 2019
# ===============================

# -------- Admin check --------
$IsAdmin = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin) {
    Write-Host "Please run PowerShell as Administrator" -ForegroundColor Red
    exit 1
}

# -------- Paths --------
$Temp = "C:\Temp\OfficeDeploy"
New-Item -ItemType Directory -Path $Temp -Force | Out-Null

# -------- Detect Office 2019 MSI --------
Write-Host "Detecting Office 2019 MSI..." -ForegroundColor Cyan

$Office2019MSI = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" |
Get-ItemProperty |
Where-Object {
    $_.DisplayName -match "Office 2019" -and
    $_.DisplayName -notmatch "Click-to-Run"
}

if ($Office2019MSI) {

    Write-Host "Office 2019 MSI detected. Starting SaRA removal..." -ForegroundColor Yellow

    # -------- Download SaRA --------
    $SaraUrl = "https://aka.ms/SaRA_OfficeUninstallFromPC"
    $SaraExe = "$Temp\SaraSetup.exe"

    Invoke-WebRequest $SaraUrl -OutFile $SaraExe

    # -------- Run SaRA silent --------
    Start-Process $SaraExe `
        -ArgumentList "-OfficeVersion 2019 -Quiet" `
        -Wait

    Write-Host "SaRA completed. Verifying removal..." -ForegroundColor Cyan

    Start-Sleep -Seconds 10

    # -------- Verify removal --------
    $Verify = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" |
    Get-ItemProperty |
    Where-Object {
        $_.DisplayName -match "Office 2019" -and
        $_.DisplayName -notmatch "Click-to-Run"
    }

    if ($Verify) {
        Write-Host "❌ Office 2019 MSI still detected. Manual intervention required." -ForegroundColor Red
        exit 2
    }

    Write-Host "✅ Office 2019 MSI successfully removed." -ForegroundColor Green
}
else {
    Write-Host "No Office 2019 MSI detected." -ForegroundColor Green
}

# -------- Deploy Microsoft 365 Apps --------
Write-Host "Starting Microsoft 365 Apps installation..." -ForegroundColor Cyan

$ODT  = "$Temp\setup.exe"
$Cfg  = "$Temp\config.xml"

Invoke-WebRequest "https://officecdn.microsoft.com/pr/wsus/setup.exe" -OutFile $ODT
Invoke-WebRequest "https://raw.githubusercontent.com/pntrivt06/office-deploy-script/main/config.xml" -OutFile $Cfg

Start-Process $ODT "/configure `"$Cfg`"" -Wait

Write-Host "✅ Office deployment completed successfully" -ForegroundColor Green
