#requires -Version 7.0

<#
.SYNOPSIS
Retrieves Entra ID users and publishes them to a custom Log Analytics table via the
Azure Monitor Log Ingestion API.

.DESCRIPTION
Authenticates with Azure using managed identity (Azure Automation runbook context),
retrieves Entra users via Microsoft Graph, transforms the result into a stable schema,
and sends records to a custom Log Analytics table through a Data Collection Rule (DCR)
and Data Collection Endpoint (DCE).

The user retrieval logic is aligned with the existing AAD export runbook:
- userType == 'Member'
- accountEnabled == true
- only users synchronized from on-premises AD (OnPremisesDistinguishedName is not empty)

If user retrieval fails or returns no valid users, ingestion is skipped.

.OUTPUTS
None. Records are published directly to Log Analytics via the Log Ingestion API.

.EXAMPLE
.\Publish-EntraIDUsersToLAW.ps1

.NOTES
Author: GitHub Copilot
Date: 2026-04-10
Version: 1.0.0
#>

[CmdletBinding()]
param ()

Connect-AzAccount -Identity | Out-Null

#region Configuration

[bool]$EnableVerbose = $true

[string]$SubscriptionId = "<your-subscription-id>"
[string]$TenantIdentifier = "<your-azure-ad-tenant-id>"
[string]$AzureKeyVaultName = "<your-key-vault-name>"

[string]$GraphResource = "https://graph.microsoft.com"
[string]$ManagedIdentityApiVersion = "2019-08-01"

[string]$LogAnalyticsWorkspaceResourceId = "/subscriptions/<your-subscription-id>/resourceGroups/<your-resource-group>/providers/Microsoft.OperationalInsights/workspaces/<your-law-name>"
[string]$AzDcrResourceGroup = "<your-dcr-resource-group>"
[string]$AzDceName = "<your-data-collection-endpoint-name>"
[string]$TableName = "entraidusers"
[string]$AzDcrName = "dcr-prd-" + $TableName + "_CL"

[bool]$AzDcrSetLogIngestApiAppPermissionsDcrLevel = $false
[array]$AzLogDcrTableCreateFromReferenceMachine = @()
[bool]$AzLogDcrTableCreateFromAnyMachine = $true

[int]$MaxIngestionPayloadBytes = 500000
[int]$IngestionRetryMaxAttempts = 4
[int]$IngestionRetryInitialDelaySeconds = 2

# Retrieve the Log Ingestion application secret from Key Vault.
[object]$LogIngestKeyVaultSecret = Get-AzKeyVaultSecret -VaultName $AzureKeyVaultName -Name "<your-log-ingestion-app-secret-name>"
[System.Security.SecureString]$LogIngestSecretValue = $LogIngestKeyVaultSecret.SecretValue
[string]$LogIngestAppSecret = $LogIngestSecretValue | ConvertFrom-SecureString -AsPlainText

[string]$LogIngestAppId = "<your-log-ingestion-app-client-id>"
[string]$AzDcrLogIngestServicePrincipalObjectId = "<your-log-ingestion-app-service-principal-object-id>"

#endregion Configuration

#region Class Definitions

class EntraUserRecord {
	[string]$Id
	[string]$DisplayName
	[string]$UserPrincipalName
	[string]$UserType
	[string]$LastInteractiveSignIn
	[string]$GivenName
	[string]$Surname
	[string]$Mail
	[string]$Department
	[string]$City
	[string]$PostalCode
	[string]$State
	[string]$Country
	[string]$CompanyName
	[string]$OnPremisesSamAccountName
	[string]$OnPremisesDistinguishedName
	[string]$OnPremisesExtensionAttribute1
	[string]$OnPremisesExtensionAttribute2
	[string]$OnPremisesExtensionAttribute5
	[string]$OnPremisesExtensionAttribute10
	[string]$OnPremisesExtensionAttribute11
	[string]$OnPremisesExtensionAttribute12
	[string]$ErrorMessage

	EntraUserRecord() {
		$this.Id = [string]::Empty
		$this.DisplayName = [string]::Empty
		$this.UserPrincipalName = [string]::Empty
		$this.UserType = [string]::Empty
		$this.LastInteractiveSignIn = [string]::Empty
		$this.GivenName = [string]::Empty
		$this.Surname = [string]::Empty
		$this.Mail = [string]::Empty
		$this.Department = [string]::Empty
		$this.City = [string]::Empty
		$this.PostalCode = [string]::Empty
		$this.State = [string]::Empty
		$this.Country = [string]::Empty
		$this.CompanyName = [string]::Empty
		$this.OnPremisesSamAccountName = [string]::Empty
		$this.OnPremisesDistinguishedName = [string]::Empty
		$this.OnPremisesExtensionAttribute1 = [string]::Empty
		$this.OnPremisesExtensionAttribute2 = [string]::Empty
		$this.OnPremisesExtensionAttribute5 = [string]::Empty
		$this.OnPremisesExtensionAttribute10 = [string]::Empty
		$this.OnPremisesExtensionAttribute11 = [string]::Empty
		$this.OnPremisesExtensionAttribute12 = [string]::Empty
		$this.ErrorMessage = [string]::Empty
	}
}

#endregion Class Definitions

#region Functions

function Split-RecordsByJsonByteSize {
	[CmdletBinding()]
	[OutputType([object[]])]
	param (
		[Parameter(Mandatory = $true)]
		[AllowEmptyCollection()]
		[array]$Records,

		[Parameter(Mandatory = $false)]
		[ValidateRange(1, 1048576)]
		[int]$MaxPayloadBytes = 950000,

		[Parameter(Mandatory = $false)]
		[ValidateRange(2, 100)]
		[int]$JsonDepth = 10
	)

	if ($null -eq $Records -or $Records.Count -eq 0) {
		return @()
	}

	[System.Collections.Generic.List[object[]]]$Batches = [System.Collections.Generic.List[object[]]]::new()
	[System.Collections.Generic.List[object]]$CurrentBatch = [System.Collections.Generic.List[object]]::new()
	[int]$CurrentBatchBytes = 2

	foreach ($Record in $Records) {
		[string]$RecordJson = $Record | ConvertTo-Json -Compress -Depth $JsonDepth
		[int]$RecordBytes = [System.Text.Encoding]::UTF8.GetByteCount($RecordJson)
		[int]$AddedBytes = $RecordBytes + $(if ($CurrentBatch.Count -gt 0) { 1 } else { 0 })

		if ($RecordBytes + 2 -gt $MaxPayloadBytes) {
			if ($CurrentBatch.Count -gt 0) {
				$Batches.Add($CurrentBatch.ToArray())
				$CurrentBatch.Clear()
				$CurrentBatchBytes = 2
			}

			$Batches.Add(@($Record))
			continue
		}

		if ($CurrentBatchBytes + $AddedBytes -le $MaxPayloadBytes) {
			$CurrentBatch.Add($Record)
			$CurrentBatchBytes += $AddedBytes
			continue
		}

		$Batches.Add($CurrentBatch.ToArray())
		$CurrentBatch.Clear()
		$CurrentBatch.Add($Record)
		$CurrentBatchBytes = 2 + $RecordBytes
	}

	if ($CurrentBatch.Count -gt 0) {
		$Batches.Add($CurrentBatch.ToArray())
	}

	return $Batches.ToArray()
}

function Invoke-LogIngestionBatchWithRetry {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[array]$DataBatch,

		[Parameter(Mandatory = $true)]
		[string]$DceName,

		[Parameter(Mandatory = $true)]
		[string]$DcrName,

		[Parameter(Mandatory = $true)]
		[string]$TableName,

		[Parameter(Mandatory = $true)]
		[string]$AzAppId,

		[Parameter(Mandatory = $true)]
		[string]$AzAppSecret,

		[Parameter(Mandatory = $true)]
		[string]$TenantId,

		[Parameter(Mandatory = $false)]
		[ValidateRange(1, 10)]
		[int]$MaxAttempts = 4,

		[Parameter(Mandatory = $false)]
		[ValidateRange(1, 60)]
		[int]$InitialDelaySeconds = 2,

		[Parameter(Mandatory = $false)]
		[bool]$EnableVerboseLogging = $false
	)

	for ([int]$Attempt = 1; $Attempt -le $MaxAttempts; $Attempt++) {
		try {
			$null = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output `
				-DceName     $DceName `
				-DcrName     $DcrName `
				-Data        $DataBatch `
				-TableName   $TableName `
				-AzAppId     $AzAppId `
				-AzAppSecret $AzAppSecret `
				-TenantId    $TenantId `
				-BatchAmount $DataBatch.Count `
				-Verbose:$EnableVerboseLogging

			return
		}
		catch {
			[int]$StatusCode = 0
			if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
				$StatusCode = [int]$_.Exception.Response.StatusCode
			}

			[bool]$IsTransient = ($StatusCode -in @(0, 408, 429, 500, 502, 503, 504))
			[bool]$ShouldRetry = ($IsTransient -and $Attempt -lt $MaxAttempts)

			if (-not $ShouldRetry) {
				throw
			}

			[int]$DelaySeconds = [int]([Math]::Pow(2, $Attempt - 1) * $InitialDelaySeconds)
			Write-Warning "Transient ingestion failure (attempt $Attempt/$MaxAttempts, status code: $StatusCode). Retrying in $DelaySeconds second(s)."
			Start-Sleep -Seconds $DelaySeconds
		}
	}
}

function New-DummyEntraUserRecord {
	[CmdletBinding()]
	[OutputType([EntraUserRecord[]])]
	param (
		[Parameter(Mandatory = $true)]
		[string]$ErrorMessage
	)

	[EntraUserRecord]$DummyRecord = [EntraUserRecord]::new()
	$DummyRecord.ErrorMessage = $ErrorMessage

	return @($DummyRecord)
}

function Get-GraphAccessToken {
	[CmdletBinding()]
	[OutputType([string])]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$ResourceUri,

		[Parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()]
		[string]$ApiVersion = '2019-08-01'
	)

	# Primary method: Az.Accounts token acquisition is the most reliable in Automation runtime environments.
	try {
		$tokenResult = Get-AzAccessToken -ResourceUrl $ResourceUri -ErrorAction Stop
		if ($tokenResult -and $null -ne $tokenResult.Token) {
			[string]$tokenText = [string]::Empty

			if ($tokenResult.Token -is [System.Security.SecureString]) {
				$tokenText = ConvertFrom-SecureString -SecureString $tokenResult.Token -AsPlainText
			}
			else {
				$tokenText = [string]$tokenResult.Token
			}

			if (-not [string]::IsNullOrWhiteSpace($tokenText)) {
				return $tokenText
			}
		}
	}
	catch {
		Write-Verbose "Get-AzAccessToken failed for '$ResourceUri'. Falling back to managed identity endpoint. Details: $($_.Exception.Message)"
	}

	# Fallback method: query the managed identity endpoint directly.
	[string]$identityEndpoint = $env:IDENTITY_ENDPOINT
	[string]$identityHeader = $env:IDENTITY_HEADER

	if ([string]::IsNullOrWhiteSpace($identityEndpoint) -or [string]::IsNullOrWhiteSpace($identityHeader)) {
		throw "Unable to acquire Graph token. Get-AzAccessToken failed and managed identity endpoint variables are unavailable."
	}

	[string]$encodedResource = [System.Uri]::EscapeDataString($ResourceUri)
	[string]$separator = $(if ($identityEndpoint.Contains('?')) { '&' } else { '?' })
	[string]$tokenUri = $identityEndpoint + $separator + "api-version=$ApiVersion&resource=$encodedResource"

	$headers = @{ 'X-Identity-Header' = $identityHeader }
	$response = Invoke-RestMethod -Method Get -Uri $tokenUri -Headers $headers -ErrorAction Stop

	if ($null -eq $response -or [string]::IsNullOrWhiteSpace([string]$response.access_token)) {
		throw "Managed identity token endpoint did not return an access token for resource '$ResourceUri'."
	}

	return [string]$response.access_token
}

function Get-EntraUserRecord {
	[CmdletBinding()]
	[OutputType([EntraUserRecord[]])]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$GraphAccessToken
	)

	try {
		Import-Module -Name Microsoft.Graph.Users -ErrorAction Stop
		[System.Security.SecureString]$SecureToken = ConvertTo-SecureString $GraphAccessToken -AsPlainText -Force
		Connect-MgGraph -AccessToken $SecureToken -NoWelcome -ErrorAction Stop | Out-Null

		$properties = @(
			'id',
			'displayName',
			'userPrincipalName',
			'accountEnabled',
			'mail',
			'department',
			'city',
			'postalCode',
			'state',
			'country',
			'companyName',
			'onPremisesSamAccountName',
			'givenName',
			'surname',
			'onPremisesDistinguishedName',
			'onPremisesExtensionAttributes',
			'userType',
			'signInActivity'
		)

		$filter = "accountEnabled eq true and userType eq 'Member'"
		$queryUsers = $null
		try {
			$queryUsers = Get-MgUser -Filter $filter -Property $properties -All -ErrorAction Stop
		}
		catch {
			if ($_.Exception.Message -match 'AuditLog.Read.All|Authentication_MSGraphPermissionMissing') {
				Write-Warning "Missing Graph permission for signInActivity. Continuing without last interactive sign-in data."
				$propertiesWithoutSignIn = $properties | Where-Object { $_ -ne 'signInActivity' }
				$queryUsers = Get-MgUser -Filter $filter -Property $propertiesWithoutSignIn -All -ErrorAction Stop
			}
			else {
				throw
			}
		}

		[System.Collections.Generic.List[EntraUserRecord]]$Results = [System.Collections.Generic.List[EntraUserRecord]]::new()

		foreach ($User in $queryUsers) {
			if ([string]::IsNullOrWhiteSpace([string]$User.OnPremisesDistinguishedName)) {
				continue
			}

			[EntraUserRecord]$Record = [EntraUserRecord]::new()
			$Record.Id = [string]$User.Id
			$Record.DisplayName = ([string]$User.DisplayName).Replace(',', '')
			$Record.UserPrincipalName = ([string]$User.UserPrincipalName).Replace(',', '')
			$Record.UserType = ([string]$User.UserType).Replace(',', '')
			if ($User.SignInActivity -and $User.SignInActivity.LastSignInDateTime) {
				$Record.LastInteractiveSignIn = [string]$User.SignInActivity.LastSignInDateTime
			}
			$Record.GivenName = ([string]$User.GivenName).Replace(',', '')
			$Record.Surname = ([string]$User.Surname).Replace(',', '')
			$Record.Mail = ([string]$User.Mail).Replace(',', '')
			$Record.Department = ([string]$User.Department).Replace(',', '')
			$Record.City = ([string]$User.City).Replace(',', '')
			$Record.PostalCode = ([string]$User.PostalCode).Replace(',', '')
			$Record.State = ([string]$User.State).Replace(',', '')
			$Record.Country = ([string]$User.Country).Replace(',', '')
			$Record.CompanyName = ([string]$User.CompanyName).Replace(',', '')
			$Record.OnPremisesSamAccountName = ([string]$User.OnPremisesSamAccountName).Replace(',', '')
			$Record.OnPremisesDistinguishedName = ([string]$User.OnPremisesDistinguishedName).Replace(',', '')

			if ($User.OnPremisesExtensionAttributes) {
				$Record.OnPremisesExtensionAttribute1 = [string]$User.OnPremisesExtensionAttributes.ExtensionAttribute1
				$Record.OnPremisesExtensionAttribute2 = [string]$User.OnPremisesExtensionAttributes.ExtensionAttribute2
				$Record.OnPremisesExtensionAttribute5 = [string]$User.OnPremisesExtensionAttributes.ExtensionAttribute5
				$Record.OnPremisesExtensionAttribute10 = [string]$User.OnPremisesExtensionAttributes.ExtensionAttribute10
				$Record.OnPremisesExtensionAttribute11 = [string]$User.OnPremisesExtensionAttributes.ExtensionAttribute11
				$Record.OnPremisesExtensionAttribute12 = [string]$User.OnPremisesExtensionAttributes.ExtensionAttribute12
			}

			if ([string]::IsNullOrWhiteSpace($Record.OnPremisesExtensionAttribute1)) { $Record.OnPremisesExtensionAttribute1 = '0' }
			if ([string]::IsNullOrWhiteSpace($Record.OnPremisesExtensionAttribute2)) { $Record.OnPremisesExtensionAttribute2 = '0' }
			if ([string]::IsNullOrWhiteSpace($Record.OnPremisesExtensionAttribute5)) { $Record.OnPremisesExtensionAttribute5 = '0' }

			$Record.OnPremisesExtensionAttribute1 = $Record.OnPremisesExtensionAttribute1.Replace(',', '')
			$Record.OnPremisesExtensionAttribute2 = $Record.OnPremisesExtensionAttribute2.Replace(',', '')
			$Record.OnPremisesExtensionAttribute5 = $Record.OnPremisesExtensionAttribute5.Replace(',', '')
			$Record.OnPremisesExtensionAttribute10 = $Record.OnPremisesExtensionAttribute10.Replace(',', '')
			$Record.OnPremisesExtensionAttribute11 = $Record.OnPremisesExtensionAttribute11.Replace(',', '')
			$Record.OnPremisesExtensionAttribute12 = $Record.OnPremisesExtensionAttribute12.Replace(',', '')

			$Results.Add($Record)
		}

		if ($Results.Count -eq 0) {
			return New-DummyEntraUserRecord -ErrorMessage "No Entra users matched the selection criteria."
		}

		return $Results.ToArray()
	}
	catch {
		[string]$CaughtErrorMessage = "Failed to retrieve Entra users from Microsoft Graph. Details: $($_.Exception.Message)"
		Write-Error -Message $CaughtErrorMessage
		return New-DummyEntraUserRecord -ErrorMessage $CaughtErrorMessage
	}
	finally {
		Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
	}
}

#endregion Functions

#region Main Script Execution

[EntraUserRecord[]]$EntraUsers = @()

try {
	Write-Host "Starting Entra ID user retrieval..." -ForegroundColor Green

	Set-AzContext -Subscription $SubscriptionId -ErrorAction Stop | Out-Null

	if ([string]::IsNullOrWhiteSpace($LogIngestAppSecret)) {
		throw "Log ingestion app secret could not be retrieved from Key Vault. Verify secret '<your-log-ingestion-app-secret-name>' in vault '$AzureKeyVaultName'."
	}

	[string]$GraphAccessToken = Get-GraphAccessToken `
		-ResourceUri      $GraphResource `
		-ApiVersion       $ManagedIdentityApiVersion

	$EntraUsers = Get-EntraUserRecord -GraphAccessToken $GraphAccessToken

	Write-Host "Retrieved $($EntraUsers.Count) Entra user record(s)." -ForegroundColor Cyan
}
catch {
	Write-Error "Entra user retrieval failed. Details: $($_.Exception.Message)"
	exit 1
}

# Do not publish to Log Analytics when only a dummy error record was returned.
if ($EntraUsers.Count -eq 0 -or
	($EntraUsers.Count -eq 1 -and -not [string]::IsNullOrEmpty($EntraUsers[0].ErrorMessage))) {
	Write-Warning "No valid Entra user records were returned. Skipping Log Analytics ingestion."
	if ($EntraUsers.Count -gt 0) {
		Write-Warning "Reason: $($EntraUsers[0].ErrorMessage)"
	}
	exit 0
}

#endregion Main Script Execution

#region Log Analytics Ingestion

# Build global variable with all Data Collection Endpoints visible to the Log Ingestion application.
try {
	$global:AzDceDetails = Get-AzDceListAll -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantIdentifier -Verbose:$EnableVerbose
}
catch {
	Write-Error "Failed to retrieve Data Collection Endpoints. Details: $($_.Exception.Message)"
}

# Build global variable with all Data Collection Rules visible to the Log Ingestion application.
try {
	$global:AzDcrDetails = Get-AzDcrListAll -AzAppId $LogIngestAppId -AzAppSecret $LogIngestAppSecret -TenantId $TenantIdentifier -Verbose:$EnableVerbose
}
catch {
	Write-Error "Failed to retrieve Data Collection Rules. Details: $($_.Exception.Message)"
}

[System.Collections.Generic.List[PSCustomObject]]$TransformedDataList = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($User in $EntraUsers) {
	[PSCustomObject]$DataObject = [PSCustomObject]@{
		Id                             = [string]$User.Id
		DisplayName                    = [string]$User.DisplayName
		UserPrincipalName              = [string]$User.UserPrincipalName
		UserType                       = [string]$User.UserType
		LastInteractiveSignIn          = [string]$User.LastInteractiveSignIn
		GivenName                      = [string]$User.GivenName
		Surname                        = [string]$User.Surname
		Mail                           = [string]$User.Mail
		Department                     = [string]$User.Department
		City                           = [string]$User.City
		PostalCode                     = [string]$User.PostalCode
		State                          = [string]$User.State
		Country                        = [string]$User.Country
		CompanyName                    = [string]$User.CompanyName
		OnPremisesSamAccountName       = [string]$User.OnPremisesSamAccountName
		OnPremisesDistinguishedName    = [string]$User.OnPremisesDistinguishedName
		OnPremisesExtensionAttribute1  = [string]$User.OnPremisesExtensionAttribute1
		OnPremisesExtensionAttribute2  = [string]$User.OnPremisesExtensionAttribute2
		OnPremisesExtensionAttribute5  = [string]$User.OnPremisesExtensionAttribute5
		OnPremisesExtensionAttribute10 = [string]$User.OnPremisesExtensionAttribute10
		OnPremisesExtensionAttribute11 = [string]$User.OnPremisesExtensionAttribute11
		OnPremisesExtensionAttribute12 = [string]$User.OnPremisesExtensionAttribute12
	}
	$TransformedDataList.Add($DataObject)
}

[array]$DataVariable = $TransformedDataList.ToArray()

$DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$EnableVerbose
$DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName
$DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$EnableVerbose
$DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$EnableVerbose

if ($DataVariable) {
	try {
		$null = CheckCreateUpdate-TableDcr-Structure `
			-AzLogWorkspaceResourceId                   $LogAnalyticsWorkspaceResourceId `
			-AzAppId                                    $LogIngestAppId `
			-AzAppSecret                                $LogIngestAppSecret `
			-TenantId                                   $TenantIdentifier `
			-DceName                                    $AzDceName `
			-DcrName                                    $AzDcrName `
			-DcrResourceGroup                           $AzDcrResourceGroup `
			-TableName                                  $TableName `
			-Data                                       $DataVariable `
			-LogIngestServicePricipleObjectId           $AzDcrLogIngestServicePrincipalObjectId `
			-AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
			-AzLogDcrTableCreateFromAnyMachine          $AzLogDcrTableCreateFromAnyMachine `
			-AzLogDcrTableCreateFromReferenceMachine    $AzLogDcrTableCreateFromReferenceMachine
	}
	catch {
		Write-Error "Failed to check or update the Table and Data Collection Rule structure. Details: $($_.Exception.Message)"
	}
}

if ($DataVariable) {
	try {
		[array]$DataBatches = Split-RecordsByJsonByteSize -Records $DataVariable -MaxPayloadBytes $MaxIngestionPayloadBytes

		[int]$PostedBatchCount = 0
		[int]$SkippedBatchCount = 0

		Write-Host "Posting $($DataVariable.Count) record(s) to Log Analytics in $($DataBatches.Count) payload batch(es)." -ForegroundColor Cyan

		# Release the full dataset before posting so the library has more heap available per batch.
		$DataVariable = $null
		[System.GC]::Collect()

		foreach ($DataBatch in $DataBatches) {
			[string]$BatchJson = $DataBatch | ConvertTo-Json -Compress -Depth 10
			[int]$BatchBytes = [System.Text.Encoding]::UTF8.GetByteCount($BatchJson)

			if ($BatchBytes -gt 1048576) {
				Write-Warning "Skipping one oversized payload batch ($BatchBytes bytes) because it exceeds 1 MB."
				$SkippedBatchCount++
				continue
			}

			Invoke-LogIngestionBatchWithRetry `
				-DataBatch            $DataBatch `
				-DceName              $AzDceName `
				-DcrName              $AzDcrName `
				-TableName            $TableName `
				-AzAppId              $LogIngestAppId `
				-AzAppSecret          $LogIngestAppSecret `
				-TenantId             $TenantIdentifier `
				-MaxAttempts          $IngestionRetryMaxAttempts `
				-InitialDelaySeconds  $IngestionRetryInitialDelaySeconds `
				-EnableVerboseLogging $EnableVerbose

			$PostedBatchCount++
		}

		Write-Host "Ingestion completed. Posted batches: $PostedBatchCount. Skipped oversized batches: $SkippedBatchCount." -ForegroundColor Green
	}
	catch {
		Write-Error "Failed to post data to Log Analytics via the Log Ingestion API. Details: $($_.Exception.Message)"
	}
}

if ($Error.Count -gt 0) {
	Write-Warning "The following non-terminating errors occurred during script execution:"
	foreach ($ErrorRecord in $Error) {
		Write-Warning $ErrorRecord.ToString()
	}
	$Error.Clear()
}

#endregion Log Analytics Ingestion
