# Deploy.ps1 - main logic

# ---- Admin check (FIXED) ----
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin) {
    Write-Host "Please run PowerShell as Administrator" -ForegroundColor Red
    exit 1
}

# ---- Detect Office ----
$reg = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"

if (Test-Path $reg) {
    $r = Get-ItemProperty $reg
    Write-Host "Detected Office:" -ForegroundColor Cyan
    Write-Host "  Product :" $r.ProductReleaseIds
    Write-Host "  Version :" $r.VersionToReport
    Write-Host "  Channel :" $r.UpdateChannel
} else {
    Write-Host "No Office detected" -ForegroundColor Cyan
}

# ---- Paths ----
$temp = "C:\Temp\OfficeDeploy"
New-Item -ItemType Directory -Path $temp -Force | Out-Null

# ---- Download ODT & config ----
Invoke-WebRequest "https://officecdn.microsoft.com/pr/wsus/setup.exe" -OutFile "$temp\setup.exe"
Invoke-WebRequest "https://raw.githubusercontent.com/pntrivt06/office-deploy-script/main/config.xml" -OutFile "$temp\config.xml"

# ---- Install Office ----
Start-Process "$temp\setup.exe" "/configure `"$temp\config.xml`"" -Wait

Write-Host "✅ Office deployment completed" -ForegroundColor Green
