$Credentials = $null
$GroupSiteUrl="https://tronoxglobal.sharepoint.com/sites/gdprtest" 
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

       

        Write-Host -ForegroundColor Green "All the automatic steps are now completed!"
        Write-Host "Please proceed with the manual steps documented on the Setup Guide!"

        $sppkgPath = (Get-Item -Path "..\GDPRStarterKit\sharepoint\solution\gdpr-starter-kit.sppkg" -Verbose).FullName
        Write-Host "You can find the .SPPKG file at the following path:" $sppkgPath

