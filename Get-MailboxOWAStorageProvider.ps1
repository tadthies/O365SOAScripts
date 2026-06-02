#requires -Modules Microsoft.Graph.Authentication
##############################################################################################
#This sample script is not supported under any Microsoft standard support program or service.
#Microsoft further disclaims all implied warranties including, without limitation, any implied
#warranties of merchantability or of fitness for a particular purpose. The entire risk arising
#out of the use or performance of the sample script and documentation remains with you. In no
#event shall Microsoft, its authors, or anyone else involved in the creation, production, or
#delivery of the scripts be liable for any damages whatsoever (including, without limitation,
#damages for loss of business profits, business interruption, loss of business information,
#or other pecuniary loss) arising out of the use of or inability to use the sample script or
#documentation, even if Microsoft has been advised of the possibility of such damages.
##############################################################################################

<#
	.Synopsis
		Get the configured third-party storage providers for a given mailbox.
	.Description
		Using Microsoft Graph, get any storage providers configured for a mailbox (which can be
		configured via Outlook on the web) and output third-party providers and the account name
		configured for a given provider.
		
		Requires an app registration with the application permission MailboxConfigItem.Read.
	.Parameter UserId
		UPN or Entra object ID of the mailbox. Supports pipeline input
		of UPNs, email addresses (if they are the same as the UPN), object IDs, or objects with
		an Id or UserPrincipalName property, such as with Get-MgUser and Get-Mailbox.
	.Parameter CloudEnvironment
		Office 365 cloud environment for the tenant. Valid values are Commercial, USGovGCC,
		USGovGCCHigh, USGovDoD, and China. Default value is Commercial.
	.Parameter ApplicationId
		Application (client) ID of the app registration to use for connecting to Microsoft Graph.
		This can be stored in the script or provided at runtime.
	.Parameter ClientSecret
		PSCredential object containing a valid client secret for the app registration. If not provided and
		a certificate thumbprint has not been stored in the script, the script will prompt for the client secret.
		(If used, it will be used instead of the certificate thumbprint stored in the script. Cannot be used
		if CertificateThumbprint parameter is used.)
	.Parameter CertificateThumbprint
		Thumbprint of a certificate for the app registration to use for authentication. This can be stored in the script
		If not used or a thumbprint is not stored in the script, client secret authentication will be used.
		(Cannot be used if ClientSecret parameter is used.)
	.Parameter TenantId
		Tenant ID or tenant domain of the tenant. This can be stored in the script or provided at runtime.
	.Example
		Get-MailboxOWAStorageProvider johndoe@contoso.com
		Get-MailboxOWAStorageProvider johndoe@contoso.com -ClientSecret (Get-Credential -UserName "DoesNotMatter")
		Get-Mailbox | Get-MailboxOWAStorageProvider

	.Notes
		Version: 2.0
		Date: June 2, 2026
#>

[CmdletBinding(DefaultParameterSetName='CS')]
param (
	[Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelinebyPropertyName=$true,Position=0)][Alias('UserPrincipalName','Id')][string]$UserId,
	[ValidateSet('Commercial','USGovGCC','USGovGCCHigh','USGovDoD','China')][string]$CloudEnvironment = 'Commercial',
	[string]$ApplicationId, # Can set a default value here so it does not need to be provided every time the script is run
	[string]$TenantId, # Can set a default value here so it does not need to be provided every time the script is run
	[Parameter(ParameterSetName='CS')][PSCredential]$ClientSecret,
	[Parameter(ParameterSetName='CT')][string]$CertificateThumbprint # If you want to set a default value, do not set it here, but in the Begin block
)

begin {
	$certThumbprint = '' # If using cert auth, you can specify the certificate thumbprint here if you want to use it every time without having to provide it when running the script.

	$connectionState = Get-MgContext
	if ($connectionState.AuthType -eq 'AppOnly' -and $connectionState.Scopes -contains 'MailboxConfigItem.Read') {
		# Already connected with app authentication that has the required permission
		Write-Verbose -Message "Will use existing Graph connection because it meets requirements."
		return
	}

	# ApplicationId and TenantId are not required parameters so the script can use an existing connection without requiring them to be provided.
	# If existing connection does not exist and they were not provided or default not set, prompt for them.
	while (-not($ApplicationId)) {
		$ApplicationId = Read-Host -Prompt "Enter the application (client) ID of the app registration to use"
	}
	while (-not($TenantId)) {
		$TenantId = Read-Host -Prompt "Enter the tenant ID or tenant domain of the tenant"
	}

	# If no auth provided and no cert stored, prompt for secret. Otherwise, certificate will be used
	if ($PSCmdlet.ParameterSetName -eq 'CS' -and $certThumbprint -eq '') {
        while ($null -eq $ClientSecret -or $ClientSecret.Password.Length -eq 0) {
            # UserName is a required parameter for Get-Credential but the value is not used
            $ClientSecret = (Get-Credential -Message "Enter a valid client secret for the app registration in the password field." -UserName "DoesNotMatter")
        }
	}
	
	switch ($CloudEnvironment) {
		"Commercial"   {$cloud = "Global"}
		"USGovGCC"     {$cloud = "Global"}
		"USGovGCCHigh" {$cloud = "USGov"}
		"USGovDoD"     {$cloud = "USGovDoD"}
		"China"        {$cloud = "China"}            
	}

	# Set the arguments to be used with Connect-MgGraph
	if ($ClientSecret) {
    	$connectionArguments = @{
			Environment = $cloud
			ContextScope = "Process"
			ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($ApplicationId, $ClientSecret.Password)
			TenantId = $TenantId
		}
	} else {
		$connectionArguments = @{
			Environment = $cloud
			ContextScope = "Process"
			CertificateThumbprint = $(if ($CertificateThumbprint) { $CertificateThumbprint } else { $certThumbprint })
			ClientId = $ApplicationId
			TenantId = $TenantId
		}
	}

    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    #Write-Host -ForegroundColor Green "$(Get-Date) Connecting to Microsoft Graph with application authentication..."
    Connect-MgGraph @connectionArguments | Out-Null
	if (-not(Get-MgContext)) {
		Write-Error -Message "Failed to connect to Microsoft Graph. Check the error message and ensure that the script's app registration details and credential information (secret or certificate) are correct."
		throw
	} elseif ((Get-MgContext).Scopes -notcontains 'MailboxConfigItem.Read') {
		Write-Error -Message "The app registration does not have the MailboxConfigItem.Read permission. Ensure that admin consent has been granted for the permission in the app registration."
		throw
	}

}

process {
	Write-Progress -Activity 'Getting OWA storage provider settings' -CurrentOperation "Mailbox for $UserId"
	$uriSegment = "/beta/users/$UserId/mailFolders/root/userConfigurations/OWA.AttachmentDataProvider"
	# Get IPM.Configuration message for OWA attachment provider (which has a roaming dictionary property that contains the settings)
	try {
		$userConfig = Invoke-MgGraphRequest -Method GET -Uri $uriSegment -OutputType PSObject
		if ($userConfig.xmlData) {
			# Convert Base64 byte array (binary) to string
			[xml]$xmlString = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($userConfig.xmlData))
			foreach ($sp in $xmlString.AttachmentDataProvider.entry) {
				# Include in output only third-party providers
				if ($sp.isThirdPartyProvider -eq $true) {	
					New-Object -TypeName psobject -Property @{
						Id = $UserId
						ProviderName = $sp.DisplayName
						ProviderAccount = $sp.associatedDataProviderAccountId
					}
				}
			}
		}
	}
	catch {
		# Parse the response message that contains the OData error response in JSON
		$responseError = (($_.ErrorDetails.Message -split "`r?`n`r?`n", 2)[1] | ConvertFrom-Json).error
		if ($responseError.message -like "*specified object was not found in the store*") {
			# Mailbox is valid but does not have any provider object, likely from the user never opening the mailbox or using OWA
		
		} else {
			Write-Warning -Message "$UserId`: $($responseError.message)"
		}
	}
}
end {
	Write-Progress -Activity 'Getting OWA storage provider settings' -Completed
}