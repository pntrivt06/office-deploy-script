# Install-Office.ps1
# Run PowerShell as Administrator

If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Please run this script as Administrator" -ForegroundColor Red
    exit 1
}

$WorkingDir = "$PSScriptRoot\OfficeInstall"
$ODTExe = "$WorkingDir\setup.exe"
$SetupURL = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
$ConfigXML = "$PSScriptRoot\config.xml"

If (!(Test-Path $WorkingDir)) { New-Item -ItemType Directory -Path $WorkingDir | Out-Null }

Write-Host "Downloading Office Deployment Tool..." -ForegroundColor Cyan
Invoke-WebRequest -Uri $SetupURL -OutFile $ODTExe

Write-Host "Removing old Office and installing Microsoft 365 Apps..." -ForegroundColor Cyan
Start-Process -FilePath $ODTExe -ArgumentList "/configure `"$ConfigXML`"" -Wait

Write-Host "Office installation completed successfully." -ForegroundColor Green
