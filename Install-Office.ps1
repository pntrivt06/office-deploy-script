# ==========================================
# Install-Office.ps1 (Interactive / GitHub)
# ==========================================

# ---- Check Admin (SAFE for iex) ----
$IsAdmin = (
    [Security.Principal.WindowsPrincipal]
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator
)

if (-not $IsAdmin) {
    Write-Host "Please run PowerShell as Administrator" -ForegroundColor Red
    exit 1
}

# ---- Detect Office ----
function Get-OfficeInfo {
    $regPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"

    if (Test-Path $regPath) {
        $reg = Get-ItemProperty $regPath
        return [PSCustomObject]@{
            Installed = $true
            Product   = $reg.ProductReleaseIds
            Version   = $reg.VersionToReport
            Channel   = $reg.UpdateChannel
            Platform  = $reg.Platform
        }
    }

    return [PSCustomObject]@{ Installed = $false }
}

$Office = Get-OfficeInfo

if ($Office.Installed) {
    Write-Host "Detected Office installation:" -ForegroundColor Cyan
    Write-Host "  Product : $($Office.Product)"
    Write-Host "  Version : $($Office.Version)"
    Write-Host "  Channel : $($Office.Channel)"
    Write-Host "  Platform: $($Office.Platform)"

    $Choice = Read-Host "Do you want to REMOVE existing Office before installing new one? (Y/N)"
} else {
    Write-Host "No Office detected." -ForegroundColor Green
    $Choice = "N"
}

# ---- GitHub settings ----
$RepoOwner = "YOUR_GITHUB_USERNAME"
$RepoName  = "office-deploy-script"
$Branch    = "main"

$RawBaseUrl = "https://raw.githubusercontent.com/$RepoOwner/$RepoName/$Branch"

# ---- Paths ----
$TempDir   = "C:\Temp\OfficeDeploy"
$ODTExe    = "$TempDir\setup.exe"
$ConfigXML = "$TempDir\config.xml"

if (!(Test-Path $TempDir)) {
    New-Item -ItemType Directory -Path $TempDir | Out-Null
}

# ---- Download files ----
Write-Host "Downloading config.xml..." -ForegroundColor Cyan
Invoke-WebRequest "$RawBaseUrl/config.xml" -OutFile $ConfigXML

Write-Host "Downloading Office Deployment Tool..." -ForegroundColor Cyan
Invoke-WebRequest "https://officecdn.microsoft.com/pr/wsus/setup.exe" -OutFile $ODTExe

# ---- Run install ----
if ($Choice -match '^[Yy]') {
    Write-Host "REMOVE + INSTALL selected" -ForegroundColor Yellow
} else {
    Write-Host "INSTALL / UPGRADE only" -ForegroundColor Yellow
}

Start-Process $ODTExe "/configure `"$ConfigXML`"" -Wait

Write-Host "✅ Completed" -ForegroundColor Green
