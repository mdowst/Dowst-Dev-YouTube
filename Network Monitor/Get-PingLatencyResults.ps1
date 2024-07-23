# From the video Diagnose Network Latency with PowerShell and SQLite - https://youtu.be/Do5Ve6M9mmU
# Set to the path of your SQLite database from Invoke-PingLatencyTest.ps1
$dbPath = ".\YourDatabaseFile.sqlite"

# Compare multiple tables
$tables = Get-MySQLiteTable -Path $dbPath
$AllResults = $tables.Name | Where-Object{ $_ -match '^Host_' } | ForEach-Object{
    Invoke-MySQLiteQuery -Path $dbPath -query "Select * from $($_)"
}

# Convert Datetime string back to Datetime object
$AllResults | ForEach-Object{
    $_.Datetime = Get-Date $_.Datetime
}

# Group results on Address
$AllResults | Group-Object -Property Address

# Group results on 1 minute intervals and 10 minute intervals
$AllResults | Group-Object -Property {$_.Datetime.ToString('yyyy-MM-dd HH:mm')}
$AllResults | Group-Object -Property {$_.Datetime.ToString('yyyy-MM-dd HH:mm').Substring(0,15)}


# Parse grouped results to object
$Averages = $AllResults | Group-Object -Property {$_.Datetime.ToString('yyyy-MM-dd HH:mm')} | ForEach-Object{
    # Create hashtable for each entry
    $hashtable = [ordered]@{
        'Datetime' = $_.Name
    }
    # Get each address in the time group
    $_.Group | Group-Object -Property Address | ForEach-Object{
        # Get the average for the address in this time group
        $measure = $_.Group | Where-Object{ $_.Status -eq 'Success'} | Measure-Object -Property RoundtripTime -Average
        # Add results to the hashtable
        $hashtable.Add($_.Name, $measure.Average)
    }
    # Convert the hashtable to a PowerShell object and output
    [pscustomobject]$hashtable
}
$Averages | Format-Table
