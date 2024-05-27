$fileName = 'Your sheet name'
$sheetName = 'Form Responses 1'
$ToProcessHeader = 'API_processed'
$ActionHeader = 'Action'
$GmailAccount = '<Your Gmail Account>'
$ClientID = Get-Secret -Name 'Gmail-ClientID' -Vault GoogleForms -AsPlainText
$ClientSecret = Get-Secret -Name 'Gmail-ClientSecret' -Vault GoogleForms -AsPlainText

Import-Module -Name UMN-Google
Import-Module -Name Mailozaurr

Function Get-ColumnLetter {
    param(
        [Parameter(Mandatory = $true)]
        [int]$ColumnNumber
    )

    [string]$columnLetter = ''
    # Get the multiple of 26
    [int]$prefix = [math]::Floor($ColumnNumber / 26)
    if ($prefix -gt 0) {
        # Add prefix column
        $columnLetter += [char]$($prefix + 64)
        $ColumnNumber = $ColumnNumber - $($prefix * 26) + 65
    }
    else {
        $ColumnNumber += 65
    }
    # Get column letter
    $columnLetter += [char]$ColumnNumber
    $columnLetter
}

Function Add-SheetColumn {
    [cmdletbinding()]
    param(
        $accessToken,
        $spreadSheetID,
        $sheetName,
        $ColumnHeader,
        $DefaultValue
    )
    $GetGSheetDataParam = @{
        accessToken   = $accessToken
        cell          = 'AllData'
        sheetName     = $sheetName
        spreadSheetID = $spreadSheetID
    }
    $sheetData = Get-GSheetData @GetGSheetDataParam

    # Get the first entry
    $entry = $sheetData | Select-Object -First 1

    $columnNumber = @($entry.psobject.Properties.Name).IndexOf($ColumnHeader)
    if ($columnNumber -eq -1) {
        $column = Get-ColumnLetter -Column @($entry.psobject.Properties.Name).Count
        $SetGSheetDataColumn = @{
            accessToken   = $accessToken
            rangeA1       = "$($column)1"
            sheetName     = $sheetName
            spreadSheetID = $spreadSheetID
            values        = @(@($ColumnHeader), @())
        }
        try {
            Set-GSheetData @SetGSheetDataColumn -ErrorAction Stop
        }
        catch {
            $errorMessage = $_.ErrorDetails.Message | ConvertFrom-Json
            if ($errorMessage.error.message -match 'exceeds grid limits') {
                # Calculat the number of columns to add
                $max = [regex]::Match($errorMessage.error.message, '(?<=max columns: )([0-9]+)').Value
                $add = @($entry.psobject.Properties.Name).Count + 1 - $max

                # Set the variables for adding sheet data
                $GetGSheetSheetIDParam = @{
                    accessToken   = $accessToken
                    sheetName     = $sheetName
                    spreadSheetID = $spreadSheetID
                }
                $sheetId = Get-GSheetSheetID @GetGSheetSheetIDParam
                # reference: https://developers.google.com/sheets/api/samples/rowcolumn#append-rows-columns
                $json = @{
                    "requests" = @(
                        @{
                            "appendDimension" = @{
                                "sheetId"   = $sheetId
                                "dimension" = "COLUMNS"
                                "length"    = $add
                            }
                        }
                    )
                } | ConvertTo-Json -Depth 3
                # Send Request to insert new column
                $uri = "https://sheets.googleapis.com/v4/spreadsheets/$($spreadSheetID):batchUpdate"
                $contenttype = 'application/json'
                Invoke-RestMethod -Method 'POST' -Uri $uri -Body $json -ContentType $contenttype -Headers @{"Authorization" = "Bearer $accessToken" }
                # Try again to insert the column header
                Set-GSheetData @SetGSheetDataColumn -ErrorAction Stop
            }
            else {
                throw $_
            }


        }

        # Set the default value for each existing row
        foreach ($entry in $sheetData) {
            $row = $sheetData.IndexOf($entry) + 2
            $SetGSheetDataParam = @{
                accessToken   = $accessToken
                sheetName     = $sheetName
                spreadSheetID = $spreadSheetID
                values        = @(@($DefaultValue), @())
                rangeA1       = "$($column)$($row)"
            }
            Set-GSheetData @SetGSheetDataParam
        }
    }
    else {
        $column = Get-ColumnLetter -Column $columnNumber
    }
    $column
}

Function Get-SheetColumn {
    [cmdletbinding()]
    param(
        $Entry,
        $ColumnHeader
    )

    $columnNumber = @($entry.psobject.Properties.Name).IndexOf($ColumnHeader)
    if ($columnNumber -eq -1) {
        throw "Column '$ColumnHeader' was not found. Run Add-SheetColumn to create the column first."
    }
    Get-ColumnLetter -Column $columnNumber
}

# Get Access Token
$GetGOAuthTokenServiceParam = @{
    scope      = "https://www.googleapis.com/auth/drive"
    iss        = "posh-760@formdevelopment.iam.gserviceaccount.com"
    jsonString = (Get-Secret -Name 'posh-760' -Vault GoogleForms -AsPlainText)
}
$accessToken = Get-GOAuthTokenService @GetGOAuthTokenServiceParam | Where-Object { $_ }

$spreadSheetID = Get-GSheetSpreadSheetID -accessToken $accessToken -fileName $fileName

# Confirm ToProcessHeader is set
$AddSheetColumn = @{
    accessToken   = $accessToken
    ColumnHeader  = $ToProcessHeader
    sheetName     = $sheetName
    spreadSheetID = $spreadSheetID
    DefaultValue  = 'New'
}
$column = Add-SheetColumn @AddSheetColumn
# Confirm ActionHeader is set
$AddSheetColumn['ColumnHeader'] = $ActionHeader
$AddSheetColumn['DefaultValue'] = ''
$column = Add-SheetColumn @AddSheetColumn

$GetGSheetDataParam = @{
    accessToken   = $accessToken
    cell          = 'AllData'
    sheetName     = $sheetName
    spreadSheetID = $spreadSheetID
}
$sheetData = Get-GSheetData @GetGSheetDataParam

foreach ($entry in $sheetData | Where-Object { $_.$ToProcessHeader -eq 'New' }) {
    Write-Host "$($entry.'What is your name?') " -NoNewline
    if ($entry.'What is the air-speed velocity of an unladen swallow?' -ne '20.1 MPH') {
        $ActionToTake = "Throw off cliff!!!"
        Write-Host $ActionToTake -ForegroundColor Red
    }
    else {
        $ActionToTake = "Go on. Off you go."
        Write-Host $ActionToTake -ForegroundColor Green
    }

    $row = $sheetData.IndexOf($entry) + 2
    $ToProcessColumn = Get-SheetColumn -Entry $entry -ColumnHeader $ToProcessHeader
    $ActionColumn = Get-SheetColumn -Entry $entry -ColumnHeader $ActionHeader
    $SetGSheetDataParam = @{
        accessToken   = $accessToken
        sheetName     = $sheetName
        spreadSheetID = $spreadSheetID
        values        = @(@('Email', $ActionToTake), @())
        rangeA1       = "$($ToProcessColumn)$($row):$($ActionColumn)$($row)"
    }
    Set-GSheetData @SetGSheetDataParam | Out-Null
}

# Get Sheet with updated column data
$sheetData = Get-GSheetData @GetGSheetDataParam
# Get Gmail Authorization
$CredentialOAuth2 = Connect-oAuthGoogle -ClientID $ClientID -ClientSecret $ClientSecret -GmailAccount $GmailAccount

# Process requests with the status of Email
foreach ($entry in $sheetData | Where-Object { $_.$ToProcessHeader -eq 'Email' }) {
    $SendEmailMessageParam = @{
        From                = @{ Name = 'The Wizard' ; Email = $GmailAccount }
        To                  = $entry.'Email Address'
        Server              = 'smtp.gmail.com'
        Text                = $entry.Action
        Subject             = "Survey Results for $($entry.'What is your name?')"
        SecureSocketOptions = 'Auto'
        Credential          = $CredentialOAuth2
        oAuth               = $true
    }
    try {
        $message = Send-EmailMessage @SendEmailMessageParam -ErrorAction Stop
    }
    catch {
        $message = $null
    }

    if ($message.Status) {
        $row = $sheetData.IndexOf($entry) + 2
        $ToProcessColumn = Get-SheetColumn -Entry $entry -ColumnHeader $ToProcessHeader
        $SetGSheetDataParam = @{
            accessToken   = $accessToken
            sheetName     = $sheetName
            spreadSheetID = $spreadSheetID
            values        = @(@('Done'), @())
            rangeA1       = "$($ToProcessColumn)$($row)"
        }
        Set-GSheetData @SetGSheetDataParam | Out-Null
    }
}