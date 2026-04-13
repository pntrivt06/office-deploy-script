# Install-Office.ps1

$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) { Write-Host "Run as Administrator"; exit 1 }

$Temp = "C:\\Temp\\OfficeInstall"
New-Item -ItemType Directory -Path $Temp -Force | Out-Null

$ODT = "$Temp\\setup.exe"
$Cfg = "$Temp\\config.xml"

Invoke-WebRequest "https://officecdn.microsoft.com/pr/wsus/setup.exe" -OutFile $ODT
Invoke-WebRequest "https://raw.githubusercontent.com/pntrivt06/office-deploy-script/main/config.xml" -OutFile $Cfg

Start-Process $ODT "/configure `"$Cfg`"" -Wait

Write-Host "Office installed"
