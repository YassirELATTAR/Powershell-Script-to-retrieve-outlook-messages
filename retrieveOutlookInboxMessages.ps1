#Edit your credentials here:
$login = "USER@TENANT.onmicrosoft.com"
$Psswd = 'PASSWORD'
$securString = ConvertTo-SecureString $Psswd  -AsPlainText -Force
$UserCredential = New-Object System.Management.Automation.PSCredential ($login, $securString)
Connect-ExchangeOnline -Credential $UserCredential
Connect-AzureAD -Credential $UserCredential
Connect-MgGraph -Scopes "Application.ReadWrite.All","User.ReadWrite.All","Directory.ReadWrite.All"

# Replace these with your own values
$client_id = "[client-id]"
$client_secret = "[client_secret_key]"
$tenant_id = "[tenant_id]"
$user_id = "[used_id]"

# Get an access token
$body = @{
    client_id     = $client_id
    client_secret = $client_secret
    scope         = "https://graph.microsoft.com/.default"
    grant_type    = "client_credentials"
}


$token_response = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenant_id/oauth2/v2.0/token" -Method POST -Body $body

# Use the access token to call Microsoft Graph API
$headers = @{
    Authorization = "Bearer $($token_response.access_token)"
}




#$url= https://graph.microsoft.com/v1.0/users/1463e072-9308-485a-997b-dd8028c1e059/messages?%24top=10&%24skip=47970
#$firsturl = https://graph.microsoft.com/v1.0/users/1463e072-9308-485a-997b-dd8028c1e059/messages?%24top=10&%24skip=180780


#Skipin npthing at first:
$skipnum=0

#Basic and Initial URL:
$url = "https://graph.microsoft.com/v1.0/users/$user_id/messages"
$counter=0
# Loop until there are no more pages of results
while ($url -ne $null) {

    #Reset the Authorization tokens in case they get blocked (could be changed to Try-Catch):
    if ($counter % 5000 -eq 0) {
        Start-Sleep -Seconds 5
        $counter=0
        $token_response = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenant_id/oauth2/v2.0/token" -Method POST -Body $body
        $headers = @{
            Authorization = "Bearer $($token_response.access_token)"
        }
        Start-Sleep -Seconds 5
    }

    #Skip the already retrived 500 items:
    $skipnum=$skipnum+500
    
    #customize the URL in case it returns to the beginning:
    $url= "https://graph.microsoft.com/v1.0/users/$user_id/messages?%24top=500&%24skip=$skipnum"
    # Get the next page of results
    $response = Invoke-RestMethod -Uri $url -Headers $headers

    
    # Process the messages
    foreach ($message in $response.value) {
        $counter++
        # Save the message body to a file
        $filename = "$($message.id)_$($counter).txt"
        Set-Content -Path $filename -Value $message.body.content
        
        Write-Host -NoNewline " Downloaded: $skipnum should have been retrieved by now"
        Write-Host -NoNewline "`r"
    }

    # Get the URL for the next page so that we can check if it is null and break the loop:
    $url = $response.'@odata.nextLink'
}

