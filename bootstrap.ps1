$base = "https://raw.githubusercontent.com/pntrivt06/office-deploy-script/main"
$dir  = "$env:TEMP\OfficeDeploy"

New-Item $dir -ItemType Directory -Force | Out-Null

Invoke-WebRequest "$base/Office-Menu.ps1"   -OutFile "$dir\Office-Menu.ps1"
Invoke-WebRequest "$base/Remove-Office.ps1" -OutFile "$dir\Remove-Office.ps1"
Invoke-WebRequest "$base/Install-Office.ps1"-OutFile "$dir\Install-Office.ps1"

powershell -ExecutionPolicy Bypass -File "$dir\Office-Menu.ps1"
