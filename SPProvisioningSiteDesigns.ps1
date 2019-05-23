#---Connect to SPO
$creds = Get-Credential
Connect-SPOService https://tenant-admin.sharepoint.com -Credential $creds

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
Add-SPOSiteDesignTask -SiteDesignId $sitedesign.Id -WebUrl "https://tenant.sharepoint.com/sites/modelproject"

#---View site designs ran on a site
Get-SPOSiteDesignRun -WebUrl "https://tenant.sharepoint.com/sites/modelproject"

#---Extract pnp template
Connect-PnPOnline https://tenant.sharepoint.com/sites/modelproject -Credentials $creds
Get-PnPProvisioningTemplate -Out "c:\temp\modelproject.xml"