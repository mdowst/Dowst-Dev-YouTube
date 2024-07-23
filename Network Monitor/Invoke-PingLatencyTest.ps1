# From the video Diagnose Network Latency with PowerShell and SQLite - https://youtu.be/Do5Ve6M9mmU
    <#
.SYNOPSIS
This script will ping multiple remote addresses and write the results and roundtrip time to a SQLite database

.PARAMETER TestIpAddresses
A string array of multiple IP addresses

.PARAMETER DatabaseFolder
The folder to store the database in. Can be local or network share.

.PARAMETER RunMinutes
The number of minutes to run the data collection for.

.EXAMPLE
$TestIpAddresses = @( 
    '192.168.1.1'
    '172.31.48.1' 
    '172.31.54.34' 
) 
.\Invoke-PingLatencyTest.ps1 -TestIpAddresses $TestIpAddresses -DatabaseFolder 'C:\Scripts\PingTests' -RunMinutes 60 

#>
param(
    [string[]]$TestIpAddresses, 
    [string]$DatabaseFolder,
    [int]$RunMinutes
) 

# Set the name of the database file to the name of the system running this script
$dbPath = Join-Path $DatabaseFolder "$([system.environment]::MachineName).sqlite"
# Only create the database file if it does not already exist 
if (-not(Test-Path $dbPath)) { 
    New-MySQLiteDB -Path $dbPath 
} 


Function New-AddressTable{ 
    <#
.SYNOPSIS
Repeatable function to set the database table for recording ping results

.PARAMETER Address
The address that is being pinged

.PARAMETER dbPath
The path to the database file

#>
    param( 
        $Address, 
        $dbPath 
    ) 

    # Set the columns and their data types
    $ColumnProperties = [ordered]@{ 
        Datetime = "text" 
        Address = "text" 
        RoundtripTime = "int" 
        Status = "text"  
    }
    # Set the address to confirm to SQLite requirements
    $addrTable = "Host_$($Address.Replace('.','_'))" 

    # Create the table if it does not already exist
    $tables = Get-MySQLiteTable -Path $dbPath 
    if ($tables.name -notcontains $addrTable) { 
        New-MySQLiteDBTable -Path $dbPath -TableName $addrTable -ColumnProperties $ColumnProperties 
    }

    $addrTable
}

# Create an object with the address to ping and table name for that address
$TestObjects = $TestIpAddresses | ForEach-Object { 
    $table = New-AddressTable -dbPath $dbPath -Address $_ 
    [PSCustomObject]@{ 
        Address = $_ 
        Table = $table 
    } 
} 

# RUn the pings in parallel and write results to the SQLite database
$TestObjects | ForEach-Object -Parallel { 
    # Create a stopwatch to ensure the while loop runs for the number of minutes specified
    $timer = [system.diagnostics.stopwatch]::StartNew() 
    while ($timer.Elapsed.TotalMinutes -lt $using:RunMinutes) { 
        # Create a stopwatch to ensure we pause for a full 5 seconds
        $runtimer = [system.diagnostics.stopwatch]::StartNew() 

        # Ping the remote address
        $ping = New-Object System.Net.NetworkInformation.Ping 
        $asyncPing = $ping.SendPingAsync($_.Address) 
        $reply = $asyncPing.GetAwaiter().GetResult() 

        # Convert the current time to ISO 8601 string
        $Date = (Get-Date).ToUniversalTime().ToString('u')

        # Write results to the database
        Invoke-MySQLiteQuery -Path $using:dbPath -query "INSERT INTO $($_.Table) (Datetime, Address, RoundtripTime, Status) VALUES ('$($Date)','$($_.Address)',$($reply.RoundtripTime),'$($reply.Status)');" 

        # pause if it hasn't been 5 seconds
        while ($runtimer.Elapsed.Seconds -lt 5) {} 
    }
} 