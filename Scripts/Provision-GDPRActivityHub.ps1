<#
.SYNOPSIS
Provisions a GDPR Activity Hub site

.EXAMPLE
PS C:\> .\Provision-GDPRActivityHub.ps1 -SiteName "GDPRActivityHub" -SiteDescription "My GDPR Activity Hub" -Credentials $credentials

.EXAMPLE
PS C:\> .\Provision-GDPRActivityHub.ps1 -SiteName "GDPRActivityHub" -SiteDescription "My GDPR Activity Hub" -ConfigureCDN -CDNSiteName "CDN" -Credentials $credentials

#>
[CmdletBinding()]
param
(
    [Parameter(Mandatory = $true, HelpMessage="The URL of the already created Modern Site")]
    [String]
    $GroupSiteUrl,

    [Parameter(ParameterSetName = "CDN", Mandatory = $true, HelpMessage="Declares whether to create and configure a CDN in the target Office 365 tenant")]
    [Switch]
    $ConfigureCDN,

    [Parameter(ParameterSetName = "CDN", Mandatory = $true, HelpMessage="The name of the Team Site that will be created to support the CDN, e.g. ""CDN""")]
    [String]
    $CDNSiteName,

    [Parameter(ParameterSetName = "CDN", Mandatory = $true, HelpMessage="The name of the Docuemnt Library that will be created to support the CDN, e.g. ""CDNFiles""")]
    [String]
    $CDNLibraryName,

    [Parameter(ParameterSetName = "CDN", Mandatory = $true, HelpMessage="The path where the SPFx client-side web parts assets are located after build")]
    [String]
    $SPFxAssetsPath,

	[Parameter(Mandatory = $false, HelpMessage="Optional tenant administration credentials")]
	[PSCredential]
	$Credentials
)

try
{
    # **********************************************
    # Provision SharePoint Online artifacts
    # **********************************************

    Write-Host "Creating artifacts on site" $GroupSiteUrl

    # Get credentials to connect to SharePoint Online, if they are missing
    if($Credentials -eq $null)
    {
	    $Credentials = Get-Credential -Message "Enter Tenant Admin Credentials"
    }

    # Connect to the target site
    Connect-PnPOnline $GroupSiteUrl -Credentials $Credentials

    # Provision taxonomy items, fields, content types, and lists
    Apply-PnPProvisioningTemplate -Path .\GDPR-Activity-Hub-Information-Architecture-Full.xml -Handlers Fields,ContentTypes,Lists,TermGroups

    # Provision workflows
    Apply-PnPProvisioningTemplate -Path .\GDPR-Activity-Hub-Workflows.xml -Handlers Workflows

    # **********************************************
    # Configure the Office 365 CDN, if requested
    # **********************************************

    if ($ConfigureCDN.IsPresent)
    {
        Write-Host "Configuring the Office 365 CDN Settings"
        $CDNDescription = "Content Delivery Network"

        $spoAdminCenterUrl = $GroupUrl.replace(".sharepoint", "-admin.sharepoint")
        $spoAdminCenterUrl = $spoAdminCenterUrl.substring(0, $spoAdminCenterUrl.IndexOf("sharepoint.com/") + 15)

        $spoRootSiteUrl = $GroupUrl.substring(0, $GroupUrl.IndexOf("sharepoint.com/") + 15)

        # Create the CDN Site
        Connect-PnPOnline $spoAdminCenterUrl -Credentials $Credentials

        # Determine the current username
        $web = Get-PnPWeb
        $context = Get-PnPContext
        $user = $web.CurrentUser
        $context.Load($user)
        Execute-PnPQuery

        $currentUser = $user.Email

        # Create a new Site Collection
        $cdnSiteURL = $spoRootSiteUrl + "sites/" + $CDNSiteName
        New-PnPTenantSite -Title $CDNDescription -Url $cdnSiteURL -Description $CDNDescription -Owner $currentUser -Lcid 1033 -Template STS#0 -TimeZone 0 -RemoveDeletedSite -Wait

        # Create the CDN Files library in the CDN site
        Connect-PnPOnline $cdnSiteURL -Credentials $Credentials
        New-PnPList -Title $CDNLibraryName -Url $CDNLibraryName -Template DocumentLibrary
    
        # Create a folder in the CDNFiles document library
        $cdnFilesLibrary = Get-PnPList -Identity $CDNLibraryName
        $packageFolder = $cdnFilesLibrary.RootFolder.Folders.Add("GDPRActivityHub")
        $context = Get-PnPContext
        $context.Load($packageFolder)
        Execute-PnPQuery

        foreach ($file in (dir $SPFxAssetsPath -File)) 
        {
            $fileStream = New-Object IO.FileStream($file.FullName, [System.IO.FileMode]::Open)
            $fileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
            $fileCreationInfo.Overwrite = $true
            $fileCreationInfo.ContentStream = $fileStream
            $fileCreationInfo.URL = $file
            $upload = $packageFolder.Files.Add($fileCreationInfo)
            $context.Load($upload)
            Execute-PnPQuery
        }

        # Configure the CDN at the tenant level
        Connect-SPOService -Url $spoAdminCenterUrl -Credential $Credentials
        Set-SPOTenantCdnEnabled -CdnType Public -Confirm:$false

        Add-SPOTenantCdnOrigin -CdnType Public -OriginUrl sites/$CDNSiteName/$CDNLibraryName -Confirm:$false
    }

    Write-Host -ForegroundColor Green "All the automatic steps are now completed!"
    Write-Host "Please proceed with the manual steps documented on the Setup Guide!"
}
catch 
{
    Write-Host -ForegroundColor Red "Exception occurred!" 
    Write-Host -ForegroundColor Red "Exception Type: $($_.Exception.GetType().FullName)"
    Write-Host -ForegroundColor Red "Exception Message: $($_.Exception.Message)"
}
