The scripts in this folder are part of the Google Forms series.

# Frontend PowerShell with Google Forms: Part 1 Getting your data
(Watch the video for this content)[https://www.youtube.com/watch?v=ZqAShden9qA]

## Get-GoogleFormData.ps1
This script demonstrates how you can get data from a Google Sheet that is attached to a Google Form. And how you can update the sheet to mark lines that have been processed.

## Requirements 
This script uses the UMN-Google module to work with the Google API. And the Secret Management modules to securely store your credentials
```powershell
Install-Module -Name UMN-Google -Force
Install-Module -Name Microsoft.PowerShell.SecretStore -Force
Install-Module -Name Microsoft.PowerShell.SecretManagement -Force
```

## Set Up the Secret Management
Below is an example of the code run in the video that added the credential JSON to a local secret store.
```powershell
Get-SecretStoreConfiguration
Register-SecretVault -ModuleName Microsoft.PowerShell.SecretStore -Name GoogleForms
$JsonContent = Get-Content "C:\Users\You\Downloads\youcredential.json" -Raw
Set-Secret -Name 'your-secret' -Secret $JsonContent -Vault GoogleForms
```