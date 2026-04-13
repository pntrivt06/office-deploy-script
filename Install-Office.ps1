# ==========================================
# Install-Office.ps1
# Run directly from GitHub RAW
# ==========================================
#Detect Office Version
function Get-OfficeInfo {
    $regPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"

    if (Test-Path $regPath) {
        $reg = Get-ItemProperty $regPath

        return [PSCustomObject]@{
            Installed       = $true
            ProductRelease  = $reg.ProductReleaseIds
            Version         = $reg.VersionToReport
            Channel         = $reg.UpdateChannel
            Platform        = $reg.Platform
            InstallPath     = $reg.InstallRoot
        }
    }
    else {
        return [PSCustomObject]@{
            Installed      = $false
        }
    }
}
# ---- Check Admin ----
If (-NOT ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Please run PowerShell as Administrator" -ForegroundColor Red
    exit 1
}

# ---- Variables ----
$RepoOwner  = "pntrivt06"
$RepoName   = "office-deploy-script"
$Branch     = "main"

$RawBaseUrl = "https://raw.githubusercontent.com/$RepoOwner/$RepoName/$Branch"

$TempDir    = "C:\Temp\OfficeDeploy"
$ODTExe     = "$TempDir\setup.exe"
$ConfigXML  = "$TempDir\config.xml"
$SetupURL   = "https://officecdn.microsoft.com/pr/wsus/setup.exe"

# ---- Create Temp Folder ----
If (!(Test-Path $TempDir)) {
    New-Item -ItemType Directory -Path $TempDir | Out-Null
}

Write-Host "Working directory: $TempDir" -ForegroundColor Cyan

# ---- Download config.xml from GitHub ----
Write-Host "Downloading config.xml from GitHub..." -ForegroundColor Cyan
Invoke-WebRequest `
    -Uri "$RawBaseUrl/config.xml" `
    -OutFile $ConfigXML `
    -UseBasicParsing

# ---- Download Office Deployment Tool ----
Write-Host "Downloading Office Deployment Tool..." -ForegroundColor Cyan
Invoke-WebRequest `
    -Uri $SetupURL `
    -OutFile $ODTExe `
    -UseBasicParsing

# ---- Install Office ----
Write-Host "Removing old Office and installing Microsoft 365 Apps..." -ForegroundColor Yellow
Start-Process `
    -FilePath $ODTExe `
    -ArgumentList "/configure `"$ConfigXML`"" `
    -Wait

Write-Host "✅ Office installation completed successfully." -ForegroundColor Green
