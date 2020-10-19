#sharepoint online powershell permissions report
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

#SPO Client Object Model Context
$siteURL = "https://foundationriskpartners.sharepoint.com/sites/bidash"
$ReportFile="C:\Users\CheikhMoctar\Documents\FRP_SQL_project\permission\group_Users.csv"  
$userId = "cmoctar@foundationrp.com"  
$pwd = Read-Host -Prompt "Enter password" -AsSecureString  
$creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userId, $pwd)  
$Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)  
$Ctx.credentials = $creds     

# geting all groups and memebers of each group
Try {
   
    #Get all Groups
    $Groups=$Ctx.Web.SiteGroups
    $Ctx.Load($Groups)
    $Ctx.ExecuteQuery()
    "$("Title") `t $("Email") `t $("LoginName")" | Out-File $ReportFile -Append
    #Get Each member from the Group
    Foreach($Group in $Groups)
    {
        Write-Host "--- $($Group.Title) --- "
 
        #Getting the members
        $SiteUsers=$Group.Users
        $Ctx.Load($SiteUsers)
        $Ctx.ExecuteQuery()
        Foreach($User in $SiteUsers)
        {
            Write-Host "$($User.Title), $($User.Email), $($User.LoginName)"
            "$($Group.Title) `t $($User.Title) `t $($User.Email) `t $($User.LoginName)" | Out-File $ReportFile -Append
        }
    }
}
Catch {
    write-host -f Red "Error getting groups and users!" $_.Exception.Message
}


