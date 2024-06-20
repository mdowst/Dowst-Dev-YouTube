# From the video Monitor Your Network with PowerShell - https://youtu.be/NsCWmYjP9F4
# The Directory to save CSV files
$CsvPath = ''

# An array of IP addresses to test
$TestIpAddresses = @(
    'RouterIP'
    'ModemIP'
    'GatewayIP'
    'ExternalID'
)

# The max time to check each IP
$timeout = 700

# Build the initial object for each IP Address
$pingTests = $TestIpAddresses | ForEach-Object {
    [pscustomobject]@{
        Datetime      = (Get-Date)
        IpAddress     = $_
        RoundtripTime = 0
        Status        = 'Start'
        LastStatus    = $null
    }
}

# Loop for 24 hours
$timer = [system.diagnostics.stopwatch]::StartNew()
while ($timer.Elapsed.TotalHours -lt 24) {
    # Create a stopwatch to ensure we pause for a full second
    $runtimer = [system.diagnostics.stopwatch]::StartNew()

    # Ping each IP address in parallel and record the results
    $pingTests = $pingTests | ForEach-Object -Parallel {
        $ping = New-Object System.Net.NetworkInformation.Ping
        $asyncPing = $ping.SendPingAsync($_.IpAddress, $using:timeout)
        $reply = $asyncPing.GetAwaiter().GetResult()
    
        if ($reply.Status -ne $_.Status) {
            Write-Verbose "$($reply.Status) -ne $($_.Status)"
        }
        $_.Datetime = (Get-Date)
        $_.RoundtripTime = $reply.RoundtripTime
        $_.LastStatus = $_.Status
        $_.Status = $reply.Status
        $_
    }

    # Only export status changes
    $pingTests | Where-Object { $_.Status -ne $_.LastStatus } | Foreach-Object {
        Export-Csv -InputObject $_ -Path (Join-Path $CsvPath "$($_.IpAddress).csv") -Append 
    }

    # Pause this loop until 1 full second has passed
    while ($runtimer.Elapsed.Seconds -lt 1) {}
    Write-Verbose $timer.Elapsed.TotalSeconds
}