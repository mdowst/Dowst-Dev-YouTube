<#
This script demonstrates how you can get data from a Google Sheet that is attached to a Google Form. 
And how you can update the sheet to mark lines that have been processed.
#>
$fileName = 'Your sheet name'
$sheetName = 'Form Responses 1'

Import-Module -Name UMN-Google

Function Get-ColumnLetter{
    param(
        [Parameter(Mandatory = $true)]
        [int]$ColumnNumber
    )

    [string]$columnLetter = ''
    # Get the multiple of 26
    [int]$prefix = [math]::Floor($ColumnNumber / 26)
    if($prefix -gt 0){
        # Add prefix column
        $columnLetter += [char]$($prefix + 64)
        $ColumnNumber = $ColumnNumber - $($prefix * 26) + 65
    }
    else{
        $ColumnNumber += 65
    }
    # Get column letter
    $columnLetter += [char]$ColumnNumber
    $columnLetter
}

# Get Access Token
$GetGOAuthTokenServiceParam = @{
	scope    = "https://www.googleapis.com/auth/drive"
	iss      = "<your-service-account"
	jsonString = (Get-Secret -Name 'your-secret' -Vault GoogleForms -AsPlainText)
}
$accessToken = Get-GOAuthTokenService @GetGOAuthTokenServiceParam | Where-Object{ $_ }

$spreadSheetID = Get-GSheetSpreadSheetID -accessToken $accessToken -fileName $fileName

$GetGSheetDataParam = @{
	accessToken   = $accessToken
	cell          = 'AllData'
	sheetName     = $sheetName
	spreadSheetID = $spreadSheetID
}
$sheetData = Get-GSheetData @GetGSheetDataParam

foreach($entry in $sheetData | Where-Object{ $_.API_processed -ne 'True' }){
    Write-Host "$($entry.'What is your name?') " -NoNewline
    if($entry.'What is the air-speed velocity of an unladen swallow?' -ne '20.1 MPH'){
        Write-Host "Throw off cliff!!!" -ForegroundColor Red
    }
    else{
        Write-Host "Go on. Off you go." -ForegroundColor Green
    }

    $row = $sheetData.IndexOf($entry) + 2

    $columnNumber = @($entry.psobject.Properties.Name).IndexOf('API_processed')
    if($columnNumber -eq -1){
        $column = Get-ColumnLetter -Column @($entry.psobject.Properties.Name).Count
        $SetGSheetDataColumn = @{
            accessToken   = $accessToken
            rangeA1       = "$($column)1"
            sheetName     = $sheetName
            spreadSheetID = $spreadSheetID
            values        = @(@('API_processed'),@())
        }
        Set-GSheetData @SetGSheetDataColumn
    }
    else{
        $column = Get-ColumnLetter -Column $columnNumber
    }

    $SetGSheetDataParam = @{
        accessToken   = $accessToken
        sheetName     = $sheetName
        spreadSheetID = $spreadSheetID
        values        = @(@('True'),@())
        rangeA1       = "$($column)$($row)"
    }
    Set-GSheetData @SetGSheetDataParam
}