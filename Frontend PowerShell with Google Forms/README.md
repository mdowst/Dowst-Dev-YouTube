The scripts in this folder are part of the Google Forms series.

# Frontend PowerShell with Google Forms: Part 1 Getting your data
(Watch the video for this content)[https://www.youtube.com/watch?v=ZqAShden9qA]

## Get-GoogleFormData.ps1
This script demonstrates how you can get data from a Google Sheet that is attached to a Google Form. And how you can update the sheet to mark lines that have been processed.

## Requirements 
This script uses the UMN-Google module to work with the Google API. And the Secret Management modules to securely store your credentials
```powershell
Install-Module -Name UMN-Google
Install-Module -Name Microsoft.PowerShell.SecretStore
Install-Module -Name Microsoft.PowerShell.SecretManagement
```

## Set Up the Secret Management
Below is an example of the code run in the video that added the credential JSON to a local secret store.
```powershell
Get-SecretStoreConfiguration
Register-SecretVault -ModuleName Microsoft.PowerShell.SecretStore -Name GoogleForms
$JsonContent = Get-Content "C:\Users\You\Downloads\youcredential.json" -Raw
Set-Secret -Name 'your-secret' -Secret $JsonContent -Vault GoogleForms
```
----

# Frontend PowerShell with Google Forms: Part 2 Replying with Gmail
(Watch the video for this content)[https://youtu.be/5gCtq6ZlwOs]

## Send-EmailFromForm.ps1
This script builds upon the Get-GoogleFormData.ps1 script to add custom email replies from a Google form.

## Requirements 
This script uses the UMN-Google and the Secret Management modules and Mailozaurr to send emails.
```powershell
Install-Module -Name UMN-Google
Install-Module -Name Microsoft.PowerShell.SecretStore
Install-Module -Name Microsoft.PowerShell.SecretManagement
Install-Module -Name Mailozaurr
```

## Setting Client Secrets 
The ClientID and CLientSecret come from the OAuth client in the Google project.
```powershell
Set-Secret -Name 'Gmail-ClientID' -Vault GoogleForms -Secret (Get-Clipboard)
Set-Secret -Name 'Gmail-ClientSecret' -Vault GoogleForms -Secret (Get-Clipboard)
```

## Sending test email
```powershell
Import-Module -Name Mailozaurr
$GmailAccount = '<yourgmail>'
$To = '<toaddress>'

$ClientID = Get-Secret -Name 'Gmail-ClientID' -Vault GoogleForms -AsPlainText
$ClientSecret = Get-Secret -Name 'Gmail-ClientSecret' -Vault GoogleForms -AsPlainText

$ConnectoAuthGoogleParam = @{
	ClientID     = $ClientID
	ClientSecret = $ClientSecret
	GmailAccount = $GmailAccount
}
$CredentialOAuth2 = Connect-oAuthGoogle @ConnectoAuthGoogleParam#[pause]

$SendEmailMessageParam = @{
	From                = @{ Name = 'The Wizard' ; Email = $GmailAccount }
	To                  = $To
	Server              = 'smtp.gmail.com'
	Text                = 'This is a test email'
	Subject             = 'Test from PowerShell'
	SecureSocketOptions = 'Auto'
	Credential          = $CredentialOAuth2
	oAuth               = $true
}
Send-EmailMessage @SendEmailMessageParam
```