# Import the CSV from Invoke-PingTest.ps1
$import = Import-Csv "<Your CSV File>"

# Filter out to failure to failure status changes
$pingTimes = $import | Where-Object{ $_.Status -eq 'Success' -or $_.LastStatus -in 'Success','Start' } 

# loop through each entry to find the up and down times
$lastTime = $null
foreach($item in $pingTimes){
    if($item.LastStatus -eq 'Start'){
        $lastTime = $null
    }
    elseif($lastTime){
        # Combine the time span with the item object to get a consolidated object
        New-TimeSpan -Start $lastTime -End $item.Datetime | Select-Object -Property @{l='Status';e={$item.Status}}, 
        @{l='Date';e={$item.Datetime}}, TotalMinutes
    }
    $lastTime = $item.Datetime
}