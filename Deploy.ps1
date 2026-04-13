# Deploy.ps1 - main logic

$IsAdmin = (
    [Security.Principal.WindowsPrincipal]
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin) {
    Write-Host "Run PowerShell as Administrator" -ForegroundColor Red
    exit 1
}

$reg = "HKLM:\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration"
if (Test-Path $reg) {
    $r = Get-ItemProperty $reg
    Write-Host "Detected Office:" $r.ProductReleaseIds $r.VersionToReport
} else {
    Write-Host "No Office detected"
}

$temp = "C:\\Temp\\OfficeDeploy"
New-Item -ItemType Directory -Path $temp -Force | Out-Null

Invoke-WebRequest "https://officecdn.microsoft.com/pr/wsus/setup.exe" -OutFile "$temp\\setup.exe"
Invoke-WebRequest "https://raw.githubusercontent.com/pntrivt06/office-deploy-script/main/config.xml" -OutFile "$temp\\config.xml"

Start-Process "$temp\\setup.exe" "/configure `"$temp\\config.xml`"" -Wait

Write-Host "Office deployment completed"
