$u = "https://raw.githubusercontent.com/pntrivt06/office-deploy-script/main/Office-Menu.ps1"
$f = "$env:TEMP\Office-Menu.ps1"

Invoke-WebRequest $u -OutFile $f

powershell -ExecutionPolicy Bypass -File $f
