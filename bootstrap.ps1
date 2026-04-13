$u="https://raw.githubusercontent.com/USERNAME/office-deploy-script/main/Deploy.ps1"
$f="$env:TEMP\Deploy.ps1"
iwr $u -OutFile $f
powershell -ExecutionPolicy Bypass -File $f
