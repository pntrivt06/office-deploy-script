$u  = "https://raw.githubusercontent.com/<USERNAME>/office-deploy-script/main/office-menu.ps1"
$f  = "$env:TEMP\office-menu.ps1"

Invoke-WebRequest $u -OutFile $f

powershell -ExecutionPolicy Bypass -File $f
``
