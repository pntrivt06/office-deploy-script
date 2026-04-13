# Remove-Office.ps1

$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) { Write-Host "Run as Administrator"; exit 1 }

$Temp = "C:\\Temp\\OfficeRemove"
New-Item -ItemType Directory -Path $Temp -Force | Out-Null

# Detect Office 2019 MSI
$Office2019MSI = Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall |
Get-ItemProperty |
Where-Object { $_.DisplayName -match "Office 2019" -and $_.DisplayName -notmatch "Click-to-Run" }

if ($Office2019MSI) {
    $SaraUrl = "https://aka.ms/SaRA_OfficeUninstallFromPC"
    $SaraExe = "$Temp\\SaRA.exe"
    Invoke-WebRequest $SaraUrl -OutFile $SaraExe
    Start-Process $SaraExe -ArgumentList "-OfficeVersion 2019 -Quiet" -Wait
}

# Remove Click-to-Run Office
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

Write-Host "Office removed"
