# Install the Microsoft Authentication Library (MSAL) for PowerShell
Install-Module -Name MSAL.PS -Scope CurrentUser

# Tenant and client details
$tenantId = "xxxx"
$clientId = "xxxx"
$scope = "https://analysis.windows.net/powerbi/api/.default"

# Authenticate and obtain an access token
$authResult = Get-MsalToken -ClientId $clientId -TenantId $tenantId -Scopes $scope

# Extract the access token
$accessToken = $authResult.AccessToken

# Use the access token to make API requests
$datasetId = "xxxx"
$url = "https://api.powerbi.com/v1.0/myorg/datasets/$datasetId/Default.UpdateParameters"

# Define the JSON body
$body = @{
    updateDetails = @(
        @{
            name = "Parameter_xxxx"
            newValue = "xxxx"
        }
    )
} | ConvertTo-Json -Depth 10 -Compress

# Send the POST request
try {
    $response = Invoke-RestMethod -Method Post -Uri $url -Headers @{
        Authorization = "Bearer $accessToken"
        "Content-Type" = "application/json"
    } -Body $body

    Write-Host "Power Query parameter successfully updated."
}
catch {
    Write-Host "Error occurred:" $_.Exception.Message
}
