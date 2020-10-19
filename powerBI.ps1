# Parameters - fill these in before running the script!
# =====================================================

$groupID = "REDACTED" # the ID of the group that hosts the dataset. Use "me" if this is your My Workspace
$datasetID = "REDACTED"
$clientId = "REDACTED"

# End Parameters =======================================

# Calls the Active Directory Authentication Library (ADAL) to authenticate against AAD
function GetAuthToken
{
       if(-not (Get-Module AzureRm.Profile)) {
         Import-Module AzureRm.Profile
       }
 
       $redirectUri = "urn:ietf:wg:oauth:2.0:oob"
 
       $resourceAppIdURI = "https://analysis.windows.net/powerbi/api"
 
       $authority = "https://login.microsoftonline.com/common/oauth2/authorize";
 
       $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
 
       $authResult = $authContext.AcquireToken($resourceAppIdURI, $clientId, $redirectUri, "Auto")
 
       return $authResult
}

# Get the auth token from AAD
$token = GetAuthToken
Write-Host $token
# Building Rest API header with authorization token
$authHeader = @{
   'Content-Type'='application/json'
   'Authorization'=$token.CreateAuthorizationHeader()
}

# properly format groups path
$groupsPath = ""
if ($groupID -eq "me") {
    $groupsPath = "myorg"
} else {
    $groupsPath = "myorg/groups/$groupID"
}

# Check the refresh history
# Uncomment + '?$top=1' to get the most recent refresh
$uri = "https://api.powerbi.com/v1.0/$groupsPath/datasets/$datasetID/refreshes"# + '?$top=1'
$refreshHistory = Invoke-RestMethod -Uri $uri –Headers $authHeader –Method GET –Verbose | Select-Object -ExpandProperty value
# Remove $refreshHistory in final
$refreshHistory