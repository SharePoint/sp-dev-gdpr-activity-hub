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

	[Parameter(Mandatory = $false, HelpMessage="Optional tenant administration credentials")]
	[PSCredential]
	$Credentials
)

try
{
    # **********************************************
    # Prompt for Disclaimer
    # **********************************************

    $wscript = New-Object -comobject wscript.shell 
    $disclaimer = $wscript.popup("This GDPR Activity Hub is intended to assist organizations with their GDPR compliance progress.  This GDPR Activity Hub should not be relied upon to determine how GDPR applies to an organization or an organization’s compliance with GDPR.  This GDPR Activity Hub does not constitute legal advice, nor does it provide any certifications or guarantees regarding GDPR compliance.  Instead, we hope the GDPR Activity Hub identifies steps that organizations can implement to simplify their GDPR compliance efforts.  The application of GDPR is highly fact-specific. We encourage all organizations using this GDPR Activity Hub to work with a legally qualified professional to discuss GDPR, how it applies specifically to their organization, and how best to ensure compliance.

    MICROSOFT MAKES NO WARRANTIES, EXPRESS, IMPLIED, OR STATUTORY, AS TO THE INFORMATION IN THIS GDPR Activity Hub. Microsoft disclaims any conditions, express or implied, or other terms that use of the Microsoft products or services will ensure the organization’s compliance with the GDPR.  This GDPR Activity Hub is provided ""as-is.""  Information and recommendations expressed in this GDPR Activity Hub may change without notice.

    This GDPR Activity Hub does not provide the user with any legal rights to any intellectual property in any Microsoft product or service.  Organizations may use this GDPR Activity Hub for internal, reference purposes only.

    © 2017 Microsoft.  All rights reserved.", ` 
    60, "DISCLAIMER", 1 + 64) 

    If ($disclaimer -eq 1) { 

        # **********************************************
        # Provision SharePoint Online artifacts
        # **********************************************

        Write-Host "Creating artifacts on target site" $GroupSiteUrl

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

            $spoAdminCenterUrl = $GroupSiteUrl.replace(".sharepoint", "-admin.sharepoint")
            $spoAdminCenterUrl = $spoAdminCenterUrl.substring(0, $spoAdminCenterUrl.IndexOf("sharepoint.com/") + 15)

            $spoRootSiteUrl = $GroupSiteUrl.substring(0, $GroupSiteUrl.IndexOf("sharepoint.com/") + 15)
            $spoTenantName = $spoRootSiteUrl.Substring(8, $spoRootSiteUrl.LastIndexOf("/") - 8)

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
            Write-Host "Creating CDN Site Collection"
            $cdnSiteURL = $spoRootSiteUrl + "sites/" + $CDNSiteName
            New-PnPTenantSite -Title $CDNDescription -Url $cdnSiteURL -Description $CDNDescription -Owner $currentUser -Lcid 1033 -Template STS#0 -TimeZone 0 -RemoveDeletedSite -Wait

            # Create the CDN Files library in the CDN site
            Connect-PnPOnline $cdnSiteURL -Credentials $Credentials
            New-PnPList -Title $CDNLibraryName -Url $CDNLibraryName -Template DocumentLibrary

            # Build and package the solution
            Write-Host "Building SPFx package and bundling"
            Push-Location ..\GDPRStarterKit

            $cdnSiteAssetsFullUrl = "https://publiccdn.sharepointonline.com/" + $spoTenantName + "/sites/" + $CDNSiteName + "/" + $CDNLibraryName + "/GDPRActivityHub"
            & npm install --save
            & gulp update-manifest --cdnpath "$cdnSiteAssetsFullUrl"
            & gulp clean
            & gulp bundle --ship
            & gulp package-solution --ship

            Pop-Location
    
            # Create a folder in the CDNFiles document library
            Write-Host "Uploading SPFx assets to the CDN"
            $cdnFilesLibrary = Get-PnPList -Identity $CDNLibraryName
            $packageFolder = $cdnFilesLibrary.RootFolder.Folders.Add("GDPRActivityHub")
            $context = Get-PnPContext
            $context.Load($packageFolder)
            Execute-PnPQuery

            $cdnSiteAssetsUploadUrl = $CDNLibraryName + "/GDPRActivityHub"
            foreach ($file in (dir ..\GDPRStarterKit\temp\deploy -File)) 
            {
                $uploadedFile = Add-PnPFile -Path $file.FullName -Folder $cdnSiteAssetsUploadUrl
            }

            # Configure the CDN at the tenant level
            Connect-SPOService -Url $spoAdminCenterUrl -Credential $Credentials
            Set-SPOTenantCdnEnabled -CdnType Public -Confirm:$false

            Add-SPOTenantCdnOrigin -CdnType Public -OriginUrl sites/$CDNSiteName/$CDNLibraryName -Confirm:$false
        }

        Write-Host -ForegroundColor Green "All the automatic steps are now completed!"
        Write-Host "Please proceed with the manual steps documented on the Setup Guide!"

        $sppkgPath = (Get-Item -Path "..\GDPRStarterKit\sharepoint\solution\gdpr-starter-kit.sppkg" -Verbose).FullName
        Write-Host "You can find the .SPPKG file at the following path:" $sppkgPath


    } else { 
        Write-Host "The operation has been cancelled."
    } 
}
catch 
{
    Write-Host -ForegroundColor Red "Exception occurred!" 
    Write-Host -ForegroundColor Red "Exception Type: $($_.Exception.GetType().FullName)"
    Write-Host -ForegroundColor Red "Exception Message: $($_.Exception.Message)"
}
