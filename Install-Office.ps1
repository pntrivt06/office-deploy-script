# ==========================================
# Install-Office.ps1 (Interactive Version)
# Office Deploy from GitHub with User Choice
# ==========================================

# ---- Check Admin ----
If (-NOT ([Security.Principal.WindowsPrincipal] \
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Please run PowerShell as Administrator" -ForegroundColor Red
    exit 1
}

# ---- Detect Office ----
function Get-OfficeInfo {
    $regPath = "HKLM:\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration"
    if (Test-Path $regPath) {
        $reg = Get-ItemProperty $regPath
        return [PSCustomObject]@{
            Installed      = $true
            Product        = $reg.ProductReleaseIds
            Version        = $reg.VersionToReport
            Channel        = $reg.UpdateChannel
            Platform       = $reg.Platform
        }
    }
    return [PSCustomObject]@{ Installed = $false }
}

$office = Get-OfficeInfo

if ($office.Installed) {
    Write-Host "Detected existing Office installation:" -ForegroundColor Cyan
    Write-Host "  Product : $($office.Product)"
    Write-Host "  Version : $($office.Version)"
    Write-Host "  Channel : $($office.Channel)"
    Write-Host "  Platform: $($office.Platform)"

    $choice = Read-Host "Do you want to REMOVE existing Office before installing new one? (Y/N)"
} else {
    Write-Host "No existing Office installation detected." -ForegroundColor Green
    $choice = "N"
}

# ---- GitHub Variables ----
$RepoOwner  = "pntrivt06"
$RepoName   = "office-deploy-script"
$Branch     = "main"
$RawBaseUrl = "https://raw.githubusercontent.com/$RepoOwner/$RepoName/$Branch"

# ---- Paths ----
$TempDir   = "C:\\Temp\\OfficeDeploy"
$ODTExe    = "$TempDir\\setup.exe"
$ConfigXML = "$TempDir\\config.xml"

# ---- Create Temp Folder ----
If (!(Test-Path $TempDir)) {
    New-Item -ItemType Directory -Path $TempDir | Out-Null
}

# ---- Download files ----
Write-Host "Downloading config.xml from GitHub..." -ForegroundColor Cyan
Invoke-WebRequest -Uri "$RawBaseUrl/config.xml" -OutFile $ConfigXML -UseBasicParsing

Write-Host "Downloading Office Deployment Tool..." -ForegroundColor Cyan
Invoke-WebRequest -Uri "https://officecdn.microsoft.com/pr/wsus/setup.exe" -OutFile $ODTExe -UseBasicParsing

# ---- Run ODT ----
if ($choice -match '^[Yy]') {
    Write-Host "Running REMOVE + INSTALL..." -ForegroundColor Yellow
} else {
    Write-Host "Running INSTALL / UPGRADE without forced removal..." -ForegroundColor Yellow
}

Start-Process -FilePath $ODTExe -ArgumentList "/configure `"$ConfigXML`"" -Wait

Write-Host "✅ Operation completed." -ForegroundColor Green
