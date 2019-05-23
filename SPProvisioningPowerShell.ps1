[CmdletBinding()]
Param(
    [string]$sitetitle,
    [string]$siteurl,
    [string]$siteowner,
    [string]$sitetype,
    [string]$sitealias,
    [string]$sitedetails
)

#create site functions
function Comm-CreateSite {
    Write-Output ("Creating a Communication Site")

    #create site
    New-PnPSite -Type CommunicationSite -Title $sitetitle -Url $siteurl -Lcid 1033  

    #wait for site to be created
    do {
        Start-Sleep -Seconds 1
        Write-Output ("Checking for site to finish creating")
        Connect-PnPOnline -Url $siteurl -Credentials $cred -ErrorAction SilentlyContinue
        $site = Get-PnPSite -ErrorAction SilentlyContinue
    } while ($site -eq $null)

    #extra pause
    Write-Output ("Pausing while site is being creating")
    Start-Sleep -Seconds 5

    #add theme
    $sitedesign = Get-PnPSiteDesign | where {$_.Title -eq "Multicolor Theme"}
    Invoke-PnPSiteDesign -Identity $sitedesign.Id  

    #add owner to group
    $ownergroup = Get-PnPGroup | where {$_.Title -like "*Owners*"}
    Add-PnPUserToGroup -LoginName $siteowner -Identity $ownergroup.Id

    #add all company if required
    if($sitedetails -like "*AllCompany*"){
        Write-Output ("Adding everyone except external users to comm site")
        $visitorgroup = Get-PnPGroup | where {$_.Title -like "*Visitor*"}
        Add-PnPUserToGroup -LoginName "c:0-.f|rolemanager|spo-grid-all-users/fa17dd8f-73cb-4300-9dfd-265b06fd8901" -Identity $visitorgroup.Id
    }
    
    
}

function Partner-CreateSite {
    Write-Output ("Creating a Partner Collaboration Site")  
    New-PnPTenantSite -Title $sitetitle -Url $siteurl -Owner $siteowner -Lcid 1033 -Template "STS#3" -TimeZone 11 -Wait

}

function Team-CreateSite {
    Write-Output ("Creating a Team Collaboration Site")  
    New-PnPSite -Type TeamSite -Title $sitetitle -Url $siteurl -Lcid 1033

}

function PM-CreateSite {
    Write-Output ("Creating a Project Management Site")  
    New-PnPSite -Type TeamSite -Alias $sitealias -Title $sitetitle

    #wait for site to be created
    Write-Output ("Checking for site to finish creating")
    do {
        Start-Sleep -Seconds 1
        Connect-PnPOnline -Url $siteurl -Credentials $cred -ErrorAction SilentlyContinue
        $site = Get-PnPSite -ErrorAction SilentlyContinue
    } while ($site -eq $null)

    #extra pause
    Write-Output ("Pausing while site is being creating")
    Start-Sleep -Seconds 5

    #add site design
    Write-Output ("Adding site design")
    $sitedesign = Get-PnPSiteDesign | where {$_.Title -eq "Project Site"}
    Invoke-PnPSiteDesign -Identity $sitedesign.Id  

    #add owner to group
    Write-Output ("Updating owners")
    $ownergroup = Get-PnPGroup | where {$_.Title -like "*Owners*"}
    Add-PnPUserToGroup -LoginName $siteowner -Identity $ownergroup.Id
    Add-PnPUserToGroup -LoginName "vlad@globomantics.org" -Identity $ownergroup.Id

    #apply pnp provisioning 
    Write-Output ("Applying PnP template")
    $filename = "pmtemplate.xml"
    #$logoname  = "pmcolorlogo.png"
    $connString = Get-AutomationVariable -Name 'StorageConnString'
    $containerName = Get-AutomationVariable -Name 'StorageContainer'

    Write-Output ("Connecting to Azure Storage")
    $storageAccount = New-AzureStorageContext -ConnectionString $connString
    Write-Output ("[[ Connected to Azure Storage")

    Write-Output ("Getting file '" + $fileName + "' from Azure Blob Store")
    Get-AzureStorageBlobContent -Blob $fileName -Container $containerName -Destination ("c:\temp\" + $fileName) -Context $storageAccount
    Write-Output ("[[ File '" + $fileName + "' saved ]]")

    Write-Output ("Applying Provisioning Template")
    Apply-PnPProvisioningTemplate -Path ("c:\temp\" + $fileName)
    Write-Output ("[[ Provisioning Template Applied ]]")

    #add to project top nav
    Connect-PnPOnline -Url "https://globomanticsorg.sharepoint.com/sites/ProjectCentral" -Credentials $cred
    Add-PnPNavigationNode -Location TopNavigationBar -Title $sitetitle -Url $siteurl -Parent 2006

    #set logo
    #$appId = Get-AutomationVariable -Name 'GroupAppId'
    #$appSecret = Get-AutomationVariable -Name 'GroupAppSecret'
    #Write-Output ("Connecting to MS Graph")
    #Connect-PnPOnline -AppId $appId -AppSecret $appSecret -AADDomain 'globomantics.org'

    #Write-Output ("Getting file '" + $logoname + "' from Azure Blob Store")
    #Get-AzureStorageBlobContent -Blob $logoname -Container $containerName -Destination ("c:\temp\" + $logoname) -Context $storageAccount
    #Write-Output ("[[ File '" + $logoname + "' saved ]]")

}
function DocRep-CreateSite {
    Write-Output ("Creating a Document Repository Site")  
    New-PnPTenantSite -Title $sitetitle -Url $siteurl -Owner $siteowner -Lcid 1033 -Template "STS#3" -TimeZone 11 -Wait

    #wait for site to be created
    do {
        Start-Sleep -Seconds 1
        Write-Output ("Checking for site to finish creating")
        Connect-PnPOnline -Url $siteurl -Credentials $cred -ErrorAction SilentlyContinue
        $site = Get-PnPSite -ErrorAction SilentlyContinue
    } while ($site -eq $null)

    #get site design for comm sites
    $sitedesign = Get-PnPSiteDesign | where {$_.Title -eq "Multicolor Theme"}
    Invoke-PnPSiteDesign -Identity $sitedesign.Id  

}

#set fields
$domain = "https://globomanticsorg.sharepoint.com"
$status = "Connecting"

#get creds
$cred = Get-AutomationPSCredential -Name "SharePoint Login"
try {
    #connect to spo
    Connect-PnPOnline "https://globomanticsorg-admin.sharepoint.com" -Credentials $cred

    #check connection
    $context = Get-PnPContext
    if($context){
        Write-Output ("Connected to SharePoint Online - Checking if site exists")
        $status = "Connected"
        $siteurl = $domain + "/sites/" + $siteurl

        #check if site exists
        $site = Get-PnPTenantSite $siteurl -ErrorAction SilentlyContinue
        if(!$site){
            $status = "Creating site"
            Write-Output ("Creating a new site for $siteurl")
            if ($sitetype -eq "Communication") {
                Comm-CreateSite
            } elseif ($sitetype -eq "Project Management") {
                PM-CreateSite
            } elseif ($sitetype -eq "Document Repository") {
                DocRep-CreateSite
            } elseif ($sitetype -eq "Partner Collaboration") {
                Partner-CreateSite
            } elseif  ($sitetype -eq "Team Collaboration") {
                Team-CreateSite
            } else {
                Write-Output ("No site for the site type: $sitetype")
            }
        } else {
            Write-Output ("$siteurl already exists")
            $status = "Site already exists"
        }

    } else {
        Write-Output ("Issue connecting to SharePoint Online")
        $status = "Error connecting to SharePoint Online"
    }
}
catch
{
    #issue with script
    $status = "Ran into an issue: $($PSItem.ToString())"
    Write-Output $status
}





#Write-Output ("Creating News Page")
#$newPage = Add-PnPClientSidePage -Name "News" -LayoutType Article -Publish
#$newText = Add-PnPClientSideText -Page "News" -Text "This is where you can find all of your news"
#$newWP = Add-PnPClientSideWebPart -Page "News" -DefaultWebPartType NewsReel

#$newsUrl = $web.Url + "/SitePages/News.aspx"
#$newNode = Add-PnPNavigationNode -Location QuickLaunch -Title "News" -Url $newsUrl
#Write-Output ("[[ News Page Created ]]")





Connect-PnPOnline -AppId "f0825d75-962a-4426-bb3c-7d7650e2b98e" -AppSecret "c7vTIwQDWxM*glvst1SGCAk:7V9X*Mp_" -AADDomain 'globomanticsorg.onmicrosoft.com'
$group = Get-PnPUnifiedGroup -Identity "WhiskeyProject"
Set-PnPUnifiedGroup -Identity $group.GroupId -Owners "drew@globomantics.org"
Set-PnPUnifiedGroup -Identity $group.GroupId -Owners "drew@globomantics.org","vlad@globomantics.org"

Set-PnPUnifiedGroup -Identity $group.GroupId -GroupLogoPath ".\preferences-color.png"


$team = New-Team -MailNickName "TestTeam" -displayname "Test Teams" -Visibility "private" -Description "Test Description" -Owner "drew@globomantics.org"
New-TeamChannel -GroupId $team.GroupId -DisplayName "Announcements 📢"
New-TeamChannel -GroupId $team.GroupId -DisplayName "Training 🏋️"
New-TeamChannel -GroupId $team.GroupId -DisplayName "Planning 📅"

$accesstoken = Get-PnPAccessToken
$createplan = @'
{
    "owner": "e3511c8a-869d-47a2-b19f-c82c9c56c72f",
    "title": "title-value"
  }
'@
$createplanuri = "https://graph.microsoft.com/v1.0/planner/plans"
Invoke-RestMethod -Uri $createplanuri -Body $createplan -ContentType "application/json" -Headers @{Authorization = "Bearer $accesstoken"} -Method Post

POST https://graph.microsoft.com/v1.0/planner/plans
Content-type: application/json
Content-length: 381

{
  "owner": "ebf3b108-5234-4e22-b93d-656d7dae5874",
  "title": "title-value"
}




#---Connect to SPO
$creds = Get-Credential
Connect-SPOService https://globomanticsorg-admin.sharepoint.com -Credential $creds

#---Get themes
Get-SPOTheme

#---Get all site scripts
Get-SPOSiteScript | select Title,Description, Id

#---Get all site designs
Get-SPOSiteDesign | select Title, Description, SiteScriptIds

#---View all site scripts for a site design
$sitedesignsitescripts = Get-SPOSiteDesign | where {$_.Title -eq "Project Site"} | select -ExpandProperty SiteScriptIds
foreach($ss in $sitedesignsitescripts){Get-SPOSiteScript -Identity $ss | select Title, Id}

#---Apply a site design (LARGE, USE THIS ONE)---
$sitedesign = Get-SPOSiteDesign | where {$_.Title -eq "Multicolor Theme"}
Add-SPOSiteDesignTask -SiteDesignId $sitedesign.Id -WebUrl "https://globomanticsorg.sharepoint.com/sites/modelproject"

#---View site designs ran on a site
Get-SPOSiteDesignRun -WebUrl "https://globomanticsorg.sharepoint.com/sites/modelproject"

#---Extract pnp template
Connect-PnPOnline https://globomanticsorg.sharepoint.com/sites/modelproject -Credentials $creds

Get-PnPProvisioningTemplate -Out "c:\temp\modelproject.xml"