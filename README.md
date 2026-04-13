# Interactive Office Deployment Script

This PowerShell script detects existing Office installations and asks the user whether to remove them before installing the latest Microsoft 365 Apps.

## Features
- Detect Office version/channel
- User decides Remove or Install only
- Current Channel
- Language: en-GB
- Excludes Outlook & Visio

## Run
Run PowerShell as Administrator:
```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
irm https://raw.githubusercontent.com/YOUR_GITHUB_USERNAME/office-deploy-script/main/Install-Office.ps1 | iex
```
