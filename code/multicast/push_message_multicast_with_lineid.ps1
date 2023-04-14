# File Selection Dialog Display
Add-Type -assemblyName System.Windows.Forms

# Various Variables
# channelAccessToken of the official line account you want to use
$channelAccessToken = "[Please enter your accessToken]"
# message that you want to push
$message = "sample message"
# url and headers used to excute multicast
$url = "https://api.line.me/v2/bot/message/multicast"
$headers = @{ "Content-Type" = "application/json"
              "Authorization" = "Bearer $channelAccessToken"
 }

# File Selection Dialog Display
$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Filter = "csvfiles|*.csv"
$dialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
$dialog.Title = "Please select a file"

# index to count number of executions for multicast
$index = 1
# lineid to Push
$lineIdToPush = @()
if ($dialog.ShowDialog() -eq "OK") {
    # If file is selected

    # Import selected files
    $users = Import-Csv $dialog.FileName -Encoding UTF8
    
    foreach($user in $users){
        
        if ($lineIdToPush.Count -lt 500) {
            # If number of lineid stored for push is less than 500
            $lineIdToPush += $user.lineId
        } else {
            try {
                # execute multicast
                $body = @{
                    "to" = $lineIdToPush; 
                    "messages" = @(@{
                            "type" = "text";
                            "text" = "$message"
                    })
                } | ConvertTo-Json
                $encodedBody= [System.Text.Encoding]::UTF8.GetBytes($body)
                $res = Invoke-WebRequest $url -Method 'POST' -Headers $headers -Body $encodedBody
                # increase index
                $index = $index + 1
                # initialize lineIdToPush and add next lineid to push
                $lineIdToPush = @()
                $lineIdToPush += $user.lineId
            } catch {
                # If executed multicast failed
                $successUserCount = $index * 500
                Write-Host ("failedLineId:$lineIdToPush")
                Write-Error -Message "Sending message failed.Please confirme failedLineId." -ErrorAction Stop
            }
        }
    }
    
    if ($lineIdToPush.Count -ne 0) {
        # If after Foreach There is lineid that was not received message
        try {
            $body = @{
                "to" = $lineIdToPush; 
                "messages" = @(@{
                        "type" = "text";
                        "text" = "$message"
                })
            } | ConvertTo-Json
            $encodedBody= [System.Text.Encoding]::UTF8.GetBytes($body)
            $res = Invoke-WebRequest $url -Method 'POST' -Headers $headers -Body $encodedBody
        } catch {
            # If executed multicast failed
            $successUserCount = $index * 500
            Write-Host ("failedLineId:$lineIdToPush")
            Write-Error -Message "Sending message failed.Please confirme failedLineId." -ErrorAction Stop
        }
    }
    Write-Host ("All process finished without failed")
} else {
    # If no file is selected
    Write-Host ("No files were selected")
}
