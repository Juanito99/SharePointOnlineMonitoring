#requires -Version 7.0
<#
.SYNOPSIS
Reads SharePoint Online site usage activity from Purview audit logs for the last 1 hour.
.DESCRIPTION
Authenticates with Microsoft Entra ID using client credentials, queries the
Microsoft Purview Management Activity API for SharePoint audit content in the
selected window, filters for site visit, download, and upload events, and
stores the result in a System.Collections.ArrayList.

The script excludes OneDrive and personal site activity by filtering out URLs
containing "-my.sharepoint.com" or personal site patterns.

The script writes the collected records to the variable $SharePointUsageArrayList
and returns the same ArrayList as output.

Optimizations include:
- Parallel content blob fetching using ForEach-Object -Parallel
- HashSet for O(1) operation lookups instead of O(n) iteration
- Pre-compiled regex patterns for OneDrive filtering
- Exponential backoff retry logic for API reliability
- Connection timeout handling
.PARAMETER TenantId
The Microsoft Entra tenant identifier (GUID).
.PARAMETER ClientId
The Microsoft Entra application (client) identifier (GUID).
.PARAMETER ClientSecret
The client secret for the Microsoft Entra application.
.PARAMETER HoursBack
The number of hours to look back from current UTC time. This script uses 1.
.PARAMETER MaxRetryCount
Maximum number of retry attempts for transient API failures.
.PARAMETER ParallelThrottleLimit
Maximum concurrent requests when fetching content blobs in parallel.
.PARAMETER RequestTimeoutSeconds
Timeout in seconds for individual HTTP requests.
.OUTPUTS
System.Collections.ArrayList containing SharePointUsageRecord objects.
.EXAMPLE
.\Get-SharePointSiteUsageInformation.ps1 -TenantId "00000000-0000-0000-0000-000000000000" -ClientId "11111111-1111-1111-1111-111111111111" -ClientSecret "<secret>"
.NOTES
Author: GitHub Copilot
Date: 2026-03-25
Version: 2.0.0
.VERSIONPREFERENCE
PowerShell 7.x
#>

[string]$M365TenantName     = '<your-m365-tenant-name>'
[string]$TenantId           = '<your-azure-ad-tenant-id>'
[string]$ClientId           = '<your-graph-app-client-id>'
[string]$ClientSecret       = [string](Get-ItemProperty -Path "HKLM:\SOFTWARE\<your-company>\AzurePowerShellReadAccess" -Name "ClientSecret" -ErrorAction "Stop").ClientSecret
[int]$HoursBack             = 1
[bool]$BypassProxy          = $true
[int]$MaxRetryCount         = 3
[int]$ParallelThrottleLimit = 5
[int]$RequestTimeoutSeconds = 60

[string]$LogIngestAppId                           = "<your-log-ingestion-app-client-id>"
[string]$LogIngestAppSecret                       = [string](Get-ItemProperty -Path "HKLM:\SOFTWARE\<your-company>\AzureLogIngestion" -Name "ClientSecret" -ErrorAction "Stop").ClientSecret
[string]$AzDcrLogIngestServicePrincipalObjectId   = "<your-log-ingestion-app-service-principal-object-id>"
[string]$LogAnalyticsWorkspaceResourceId          = "/subscriptions/<your-subscription-id>/resourceGroups/<your-resource-group>/providers/Microsoft.OperationalInsights/workspaces/<your-law-name>"
[string]$AzDcrResourceGroup                       = "<your-dcr-resource-group>"
[bool]$AzDcrSetLogIngestApiAppPermissionsDcrLevel = $false
[array]$AzLogDcrTableCreateFromReferenceMachine   = @()
[bool]$AzLogDcrTableCreateFromAnyMachine          = $true
[string]$AzDceName                                = "<your-data-collection-endpoint-name>"
[string]$TableName                                = "sharepointusageinfo"
[string]$AzDcrName                                = "dcr-prd-" + $TableName + "_CL"
[bool]$EnableVerbose                              = $true
[string]$VerbosePreference                        = "SilentlyContinue"  # "Continue"


# Strongly typed class representing a SharePoint usage record.
class SharePointUsageRecord {
	[datetime]$CreationTime
	[string]$UserId
	[string]$Operation
	[string]$SiteUrl
	[string]$SiteNameShort
	[string]$ObjectId
	[string]$FileExtension
	[string]$Parameters
	[string]$QuotaDetails

	SharePointUsageRecord() {
		$this.CreationTime         = [datetime]::MinValue
		$this.UserId               = [string]::Empty
		$this.Operation            = [string]::Empty
		$this.SiteUrl              = [string]::Empty
		$this.SiteNameShort        = [string]::Empty
		$this.ObjectId             = [string]::Empty
		$this.FileExtension        = [string]::Empty
		$this.Parameters           = [string]::Empty
		$this.QuotaDetails         = [string]::Empty
	}
}

# Creates a single safe default SharePoint usage record for error and empty-result scenarios.
function New-DummySharePointUsageRecord {
	<#
	.SYNOPSIS
	Creates a default SharePoint usage record.
	.DESCRIPTION
	Returns a single SharePointUsageRecord object populated with safe defaults.
	.OUTPUTS
	SharePointUsageRecord[] containing one default object.
	.EXAMPLE
	New-DummySharePointUsageRecord
	.NOTES
	Author: GitHub Copilot
	Date: 2026-03-25
	Version: 1.0.0
	.VERSIONPREFERENCE
	PowerShell 7.x
	#>
	[CmdletBinding()]
	param()

	[SharePointUsageRecord]$dummyRecord = [SharePointUsageRecord]::new()

	return @($dummyRecord)
} # end function New-DummySharePointUsageRecord

# Retrieves SharePoint usage events from Purview logs for the specified period.
function Get-SharePointUsageRecordsFromPurview {
	<#
	.SYNOPSIS
	Retrieves SharePoint usage events from Purview logs.
	.DESCRIPTION
	Uses OAuth client credentials to call the Purview Management Activity API,
	reads Audit.SharePoint content blobs for the selected period, and extracts
	site visit, download, and upload events. OneDrive and personal site events
	are excluded.

	Optimizations:
	- HashSet for O(1) operation lookup
	- Pre-compiled regex for OneDrive filtering
	- Parallel content blob fetching
	- Exponential backoff retry logic
	.PARAMETER TenantIdentifier
	The Microsoft Entra tenant identifier (GUID).
	.PARAMETER ApplicationClientId
	The application (client) identifier (GUID).
	.PARAMETER ApplicationClientSecret
	The client secret associated with the application identifier.
	.PARAMETER LookbackHours
	Number of hours to query backwards from current UTC time.
	.PARAMETER MaxRetries
	Maximum retry attempts for transient failures.
	.PARAMETER ThrottleLimit
	Maximum concurrent content blob requests.
	.PARAMETER TimeoutSeconds
	HTTP request timeout in seconds.
	.OUTPUTS
	SharePointUsageRecord[]
	.EXAMPLE
	Get-SharePointUsageRecordsFromPurview -TenantIdentifier $TenantId -ApplicationClientId $ClientId -ApplicationClientSecret $ClientSecret -LookbackHours 1
	.NOTES
	Author: GitHub Copilot
	Date: 2026-03-25
	Version: 2.0.0
	.VERSIONPREFERENCE
	PowerShell 7.x
	#>
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$TenantIdentifier,

		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$ApplicationClientId,

		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$ApplicationClientSecret,

		[Parameter(Mandatory = $false)]
		[ValidateRange(1, 24)]
		[int]$LookbackHours = 1,

		[Parameter(Mandatory = $false)]
		[bool]$UseNoProxy = $true,

		[Parameter(Mandatory = $false)]
		[ValidateRange(1, 5)]
		[int]$MaxRetries = 3,

		[Parameter(Mandatory = $false)]
		[ValidateRange(1, 20)]
		[int]$ThrottleLimit = 5,

		[Parameter(Mandatory = $false)]
		[ValidateRange(10, 300)]
		[int]$TimeoutSeconds = 60
	)

	# Pre-compiled regex for OneDrive/personal site detection - compiled once, reused many times.
	[regex]$oneDriveRegex = [regex]::new('-my\.sharepoint\.com|/personal/|onedrive|' + $M365TenantName + '-my', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Compiled)

	# HashSet for O(1) target operation lookup instead of O(n) iteration.
	[System.Collections.Generic.HashSet[string]]$targetOperationsSet = [System.Collections.Generic.HashSet[string]]::new(
		[System.StringComparer]::OrdinalIgnoreCase
	)
	[void]$targetOperationsSet.Add('FileAccessed')
	[void]$targetOperationsSet.Add('FileDownloaded')
	[void]$targetOperationsSet.Add('FileUploaded')
	[void]$targetOperationsSet.Add('PageViewed')
	[void]$targetOperationsSet.Add('SetSiteProperties')
	[void]$targetOperationsSet.Add('SiteCollectionQuotaModified')
	[void]$targetOperationsSet.Add('SiteStorageLimitModeChanged')

	# Invokes Purview Management Activity API with retry logic and timeout handling.
	function Invoke-PurviewManagementApiRequestWithRetry {
		[CmdletBinding()]
		param(
			[Parameter(Mandatory = $true)]
			[ValidateNotNullOrEmpty()]
			[string]$Uri,

			[Parameter(Mandatory = $true)]
			[ValidateNotNull()]
			[hashtable]$Headers,

			[Parameter(Mandatory = $false)]
			[ValidateSet("Get", "Post")]
			[string]$Method = "Get",

			[Parameter(Mandatory = $false)]
			[AllowNull()]
			[object]$Body = $null,

			[Parameter(Mandatory = $false)]
			[string]$ContentType = [string]::Empty,

			[Parameter(Mandatory = $false)]
			[bool]$NoProxy = $true,

			[Parameter(Mandatory = $false)]
			[int]$MaxRetryAttempts = 3,

			[Parameter(Mandatory = $false)]
			[int]$TimeoutInSeconds = 60
		)

		[int]$retryAttempt        = 0
		[int]$baseDelaySeconds    = 2
		[Exception]$lastException = $null

		while ($retryAttempt -lt $MaxRetryAttempts) {
			try {
				[hashtable]$requestParameters = @{
					Method      = $Method
					Uri         = $Uri
					Headers     = $Headers
					ErrorAction = "Stop"
					TimeoutSec  = $TimeoutInSeconds
				}

				if ($null -ne $Body) {
					$requestParameters["Body"] = $Body
				}

				if (-not [string]::IsNullOrWhiteSpace($ContentType)) {
					$requestParameters["ContentType"] = $ContentType
				}

				if ($NoProxy) {
					$requestParameters["NoProxy"] = $true
				}

				return Invoke-RestMethod @requestParameters
			}
			catch {
				$lastException   = $_.Exception
				[string]$errorMessage = [string]$_.Exception.Message

				# Check if error is retryable (transient network/server errors).
				[bool]$isRetryable = $false
				if ($errorMessage -match '(timeout|timed out|503|502|500|429|too many requests|connection|network|temporarily unavailable)') {
					$isRetryable = $true
				}

				if ($null -ne $_.Exception.Response) {
					[int]$statusCode = [int]$_.Exception.Response.StatusCode
					if ($statusCode -in @(429, 500, 502, 503, 504)) {
						$isRetryable = $true
					}
				}

				$retryAttempt++

				if ($isRetryable -and $retryAttempt -lt $MaxRetryAttempts) {
					# Exponential backoff: 2^attempt * base delay + random jitter.
					[int]$delaySeconds = [Math]::Pow(2, $retryAttempt) * $baseDelaySeconds
					[int]$jitterMs     = Get-Random -Minimum 100 -Maximum 500
					[int]$totalDelayMs = ($delaySeconds * 1000) + $jitterMs

					Write-Verbose -Message "Transient error on attempt $retryAttempt. Retrying in $delaySeconds seconds. Error: $errorMessage"
					Start-Sleep -Milliseconds $totalDelayMs
				}
				else {
					# Non-retryable error or max retries exhausted - build detailed message.
					[string]$detailedMessage = $errorMessage
					[string]$uriHost         = ([Uri]$Uri).Host
					$detailedMessage         = "$detailedMessage | RequestHost: $uriHost | Attempts: $retryAttempt"

					if ($null -ne $_.ErrorDetails -and -not [string]::IsNullOrWhiteSpace($_.ErrorDetails.Message)) {
						$detailedMessage = "$detailedMessage | ErrorDetails: $($_.ErrorDetails.Message)"
					}

					if ($null -ne $_.Exception.Response) {
						try {
							$responseStream = $_.Exception.Response.GetResponseStream()
							if ($null -ne $responseStream) {
								$streamReader         = [System.IO.StreamReader]::new($responseStream)
								[string]$responseBody = $streamReader.ReadToEnd()
								$streamReader.Dispose()
								$responseStream.Dispose()

								if (-not [string]::IsNullOrWhiteSpace($responseBody)) {
									$detailedMessage = "$detailedMessage | ResponseBody: $responseBody"
								}
							}
						}
						catch {
							# Keep original exception message when response parsing fails.
						}
					}

					throw $detailedMessage
				}
			}
		}

		# Should not reach here, but safety throw.
		throw "Maximum retry attempts ($MaxRetryAttempts) exhausted for URI: $Uri. Last error: $($lastException.Message)"
	} # end function Invoke-PurviewManagementApiRequestWithRetry

	# Asserts Audit.SharePoint subscription exists before querying content.
	function Assert-SharePointAuditSubscription {
		[CmdletBinding()]
		param(
			[Parameter(Mandatory = $true)]
			[ValidateNotNullOrEmpty()]
			[string]$TenantIdentifier,

			[Parameter(Mandatory = $true)]
			[ValidateNotNull()]
			[hashtable]$Headers,

			[Parameter(Mandatory = $false)]
			[bool]$NoProxy = $true,

			[Parameter(Mandatory = $false)]
			[int]$MaxRetryAttempts = 3,

			[Parameter(Mandatory = $false)]
			[int]$TimeoutInSeconds = 60
		)

		[string]$subscriptionListUri = "https://manage.office.com/api/v1.0/$TenantIdentifier/activity/feed/subscriptions/list"
		Write-Verbose -Message "Listing subscriptions: $subscriptionListUri"
		[object[]]$subscriptions     = @(Invoke-PurviewManagementApiRequestWithRetry -Method "Get" -Uri $subscriptionListUri -Headers $Headers -NoProxy $NoProxy -MaxRetryAttempts $MaxRetryAttempts -TimeoutInSeconds $TimeoutInSeconds)

		[bool]$subscriptionExists  = $false
		[bool]$subscriptionEnabled = $false

		foreach ($subscription in $subscriptions) {
			if ([string]$subscription.contentType -eq "Audit.SharePoint") {
				$subscriptionExists = $true
				if ([string]$subscription.status -eq "enabled") {
					$subscriptionEnabled = $true
				}
				break
			}
		}

		if ((-not $subscriptionExists) -or (-not $subscriptionEnabled)) {
			[string]$subscriptionStartUri = "https://manage.office.com/api/v1.0/$TenantIdentifier/activity/feed/subscriptions/start?contentType=Audit.SharePoint"
			Write-Verbose -Message "Starting subscription: $subscriptionStartUri"
			[int]$errorCountBeforeStart = $Error.Count
			try {
				[void](Invoke-PurviewManagementApiRequestWithRetry -Method "Post" -Uri $subscriptionStartUri -Headers $Headers -NoProxy $NoProxy -MaxRetryAttempts $MaxRetryAttempts -TimeoutInSeconds $TimeoutInSeconds)
			}
			catch {
				[string]$startSubscriptionError = [string]$_.Exception.Message
				if ($startSubscriptionError -notmatch "already\s+enabled") {
					throw
				}
				# Remove all benign "already enabled" errors that accumulated
				# in $Error (Invoke-RestMethod + re-throw each add one entry).
				[int]$errorsToRemove = $Error.Count - $errorCountBeforeStart
				for ([int]$errorIndex = 0; $errorIndex -lt $errorsToRemove; $errorIndex++) {
					$Error.RemoveAt(0)
				}
			}
		}
	} # end function Assert-SharePointAuditSubscription

	# Thread-safe collection for parallel processing results (uses PSObject because custom classes aren't available in parallel runspaces).
	[System.Collections.Concurrent.ConcurrentBag[psobject]]$usageRecordsBag = [System.Collections.Concurrent.ConcurrentBag[psobject]]::new()

	# Counters for diagnostics.
	[int]$contentItemCount      = 0
	[int]$oneDriveExcludedCount = 0
	[datetime]$endTimeUtc       = [datetime]::UtcNow
	[datetime]$startTimeUtc     = $endTimeUtc.AddHours(-1 * $LookbackHours)

	try {
		[hashtable]$tokenRequestBody = @{
			client_id     = $ApplicationClientId
			client_secret = $ApplicationClientSecret
			scope         = "https://manage.office.com/.default"
			grant_type    = "client_credentials"
		}

		[hashtable]$tokenHeaders = @{
			Accept = "application/json"
		}

		[string]$tokenUri              = "https://login.microsoftonline.com/$TenantIdentifier/oauth2/v2.0/token"
		[pscustomobject]$tokenResponse = Invoke-PurviewManagementApiRequestWithRetry -Method "Post" -Uri $tokenUri -Headers $tokenHeaders -Body $tokenRequestBody -ContentType "application/x-www-form-urlencoded" -NoProxy $UseNoProxy -MaxRetryAttempts $MaxRetries -TimeoutInSeconds $TimeoutSeconds
		[string]$accessToken           = [string]$tokenResponse.access_token

		if ([string]::IsNullOrWhiteSpace($accessToken)) {
			throw "Access token was not returned by Microsoft Entra ID."
		}

		[hashtable]$headers = @{
			Authorization = "Bearer $accessToken"
			Accept        = "application/json"
		}

		Assert-SharePointAuditSubscription -TenantIdentifier $TenantIdentifier -Headers $headers -NoProxy $UseNoProxy -MaxRetryAttempts $MaxRetries -TimeoutInSeconds $TimeoutSeconds

		# Collect all content URIs first to enable parallel fetching.
		[System.Collections.Generic.List[string]]$allAuditDataUris = [System.Collections.Generic.List[string]]::new()
		[datetime]$windowStartUtc                                  = $startTimeUtc

		while ($windowStartUtc -lt $endTimeUtc) {
			[datetime]$windowEndUtc = $windowStartUtc.AddHours(24)

			if ($windowEndUtc -gt $endTimeUtc) {
				$windowEndUtc = $endTimeUtc
			}

			[string]$startTimeFormatted = $windowStartUtc.ToString("yyyy-MM-ddTHH:mm:ss")
			[string]$endTimeFormatted   = $windowEndUtc.ToString("yyyy-MM-ddTHH:mm:ss")

			[string]$contentUri = "https://manage.office.com/api/v1.0/$TenantIdentifier/activity/feed/subscriptions/content?contentType=Audit.SharePoint&startTime=$startTimeFormatted&endTime=$endTimeFormatted"
			Write-Verbose -Message "Querying content URI: $contentUri"
			[object]$contentResponse = Invoke-PurviewManagementApiRequestWithRetry -Method "Get" -Uri $contentUri -Headers $headers -NoProxy $UseNoProxy -MaxRetryAttempts $MaxRetries -TimeoutInSeconds $TimeoutSeconds

			[object[]]$contentItems = @()
			if ($null -ne $contentResponse) {
				if ($contentResponse -is [System.Collections.IEnumerable] -and -not ($contentResponse -is [string])) {
					$contentItems = @($contentResponse)
				}
				elseif ($null -ne $contentResponse.contentUri) {
					$contentItems = @($contentResponse)
				}
			}

			$contentItemCount += $contentItems.Count

			foreach ($contentItem in $contentItems) {
				if ($null -eq $contentItem -or $null -eq $contentItem.contentUri) {
					continue
				}

				[string]$uriValue = [string]$contentItem.contentUri
				if (-not [string]::IsNullOrWhiteSpace($uriValue)) {
					[void]$allAuditDataUris.Add($uriValue)
				}
			}

			$windowStartUtc = $windowEndUtc
		}

		Write-Verbose -Message "Collected $($allAuditDataUris.Count) content blob URIs to process in parallel"

		# Process content blobs in parallel for improved throughput.
		if ($allAuditDataUris.Count -gt 0) {
			$allAuditDataUris | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
				# Note: $using: variables cannot have type casts in parallel runspaces.
				$auditDataUri     = $_
				$parallelHeaders  = $using:headers
				$parallelNoProxy  = $using:UseNoProxy
				$parallelMaxRetries = $using:MaxRetries
				$parallelTimeout  = $using:TimeoutSeconds
				$resultBag        = $using:usageRecordsBag
				$operationsSet    = $using:targetOperationsSet
				$oneDrivePattern  = $using:oneDriveRegex

				try {
					# Inline retry logic for parallel execution. - works
					[object]$rawAuditResponse = $null
					[int]$attempt             = 0

					while ($attempt -lt $parallelMaxRetries) {
						try {
							[hashtable]$requestParams = @{
								Method      = "Get"
								Uri         = $auditDataUri
								Headers     = $parallelHeaders
								ErrorAction = "Stop"
								TimeoutSec  = $parallelTimeout
							}

							if ($parallelNoProxy) {
								$requestParams["NoProxy"] = $true
							}

							$rawAuditResponse = Invoke-RestMethod @requestParams
							break
						}
						catch {
							$attempt++
							if ($attempt -lt $parallelMaxRetries) {
								[int]$delayMs = [Math]::Pow(2, $attempt) * 2000 + (Get-Random -Minimum 100 -Maximum 500)
								Start-Sleep -Milliseconds $delayMs
							}
							else {
								throw
							}
						}
					}

					[object[]]$auditEvents = @()
					if ($rawAuditResponse -is [string]) {
						$auditEvents = @($rawAuditResponse | ConvertFrom-Json -Depth 20)
					}
					elseif ($null -ne $rawAuditResponse) {
						$auditEvents = @($rawAuditResponse)
					}

					foreach ($auditEvent in $auditEvents) {
						if ($null -eq $auditEvent -or $null -eq $auditEvent.PSObject) {
							continue
						}

						[string]$workloadValue  = if ($null -ne $auditEvent.PSObject.Properties['Workload']) { [string]$auditEvent.Workload } else { [string]::Empty }
						[string]$operationValue = if ($null -ne $auditEvent.PSObject.Properties['Operation']) { [string]$auditEvent.Operation } else { [string]::Empty }
						[string]$siteUrlValue   = if ($null -ne $auditEvent.PSObject.Properties['SiteUrl']) { [string]$auditEvent.SiteUrl } else { [string]::Empty }

						# Filter for SharePoint workload only.
						if ($workloadValue -ne 'SharePoint') {
							continue
						}

						# O(1) HashSet lookup instead of O(n) loop.
						if (-not $operationsSet.Contains($operationValue)) {
							continue
						}

						# Pre-compiled regex for OneDrive filtering.
						if (-not [string]::IsNullOrWhiteSpace($siteUrlValue) -and $oneDrivePattern.IsMatch($siteUrlValue)) {
							continue
						}

						# Create PSCustomObject because custom classes aren't available in parallel runspaces.
						[string]$userIdCandidate = [string]$auditEvent.UserId
						[string]$resolvedUserId  = if (-not [string]::IsNullOrWhiteSpace($userIdCandidate)) { $userIdCandidate } else { [string]$auditEvent.UserKey }
						[string]$objectIdValue   = if ($null -ne $auditEvent.PSObject.Properties['ObjectId']) { [string]$auditEvent.ObjectId } else { [string]::Empty }

						# Skip SharePoint system layout pages.
						if ($objectIdValue.StartsWith("https://${M365TenantName}.sharepoint.com/_layouts/15", [System.StringComparison]::OrdinalIgnoreCase)) {
							continue
						}

						# Skip OneDrive personal site pages.
						if ($objectIdValue.StartsWith("https://${M365TenantName}-my.sharepoint.com/personal", [System.StringComparison]::OrdinalIgnoreCase)) {
							continue
						}

						# Skip OneDrive layout pages.
						if ($objectIdValue.StartsWith("https://${M365TenantName}-my.sharepoint.com/_layouts/", [System.StringComparison]::OrdinalIgnoreCase)) {
							continue
						}

						# Skip content storage pages.
						if ($objectIdValue.StartsWith("https://${M365TenantName}.sharepoint.com/contentstorage", [System.StringComparison]::OrdinalIgnoreCase)) {
							continue
						}

						# Define admin operations that should not be filtered by SiteNameShort.
						[bool]$isAdminOperation = ($operationValue -eq 'SetSiteProperties' -or $operationValue -eq 'SiteCollectionQuotaModified' -or $operationValue -eq 'SiteStorageLimitModeChanged')

						# Derive SiteNameShort: for PageViewed use ObjectId (since SiteUrl is empty), otherwise use SiteUrl.
						# Extract only the site name (first path segment after /sites/).
						[string]$sourceUrlForSiteName = if ($operationValue -eq 'PageViewed' -and -not [string]::IsNullOrWhiteSpace($objectIdValue)) { $objectIdValue } else { $siteUrlValue }
						[string]$siteNameShortValue   = [string]::Empty

						if ($operationValue -eq 'SiteCollectionQuotaModified' -and -not [string]::IsNullOrWhiteSpace($objectIdValue)) {
							$siteNameShortValue = $objectIdValue -replace ('^https://' + $M365TenantName + '\.sharepoint\.com/sites/'), ''
						}
						elseif ($sourceUrlForSiteName -match '/sites/([^/]+)') {
							$siteNameShortValue = $matches[1]
						}
						elseif ($isAdminOperation -and -not [string]::IsNullOrWhiteSpace($sourceUrlForSiteName)) {
							# For admin operations without /sites/, extract host or use URL as identifier.
							try {
								[Uri]$parsedUri     = [Uri]$sourceUrlForSiteName
								$siteNameShortValue = $parsedUri.Host
							}
							catch {
								$siteNameShortValue = $sourceUrlForSiteName
							}
						}

						# Skip records where SiteNameShort is empty (no valid site name could be derived), except for admin operations.
						if ([string]::IsNullOrWhiteSpace($siteNameShortValue) -and -not $isAdminOperation) {
							continue
						}

						# For admin operations with still-empty SiteNameShort, use a placeholder.
						if ([string]::IsNullOrWhiteSpace($siteNameShortValue) -and $isAdminOperation) {
							$siteNameShortValue = 'TenantAdmin'
						}

						# Extract Parameters for admin operations (contains quota values and other settings).
						[string]$parametersValue = [string]::Empty
						if ($isAdminOperation -and $null -ne $auditEvent.PSObject.Properties['Parameters']) {
							try {
								$parametersValue = ($auditEvent.Parameters | ConvertTo-Json -Depth 10 -Compress)
							}
							catch {
								$parametersValue = [string]::Empty
							}
						}

						# Extract specific quota-related properties into a JSON object.
						[string]$quotaDetailsValue = [string]::Empty
						if ($isAdminOperation) {
							[hashtable]$quotaHashtable = @{}
							[string[]]$quotaPropertyNames = @(
								'StorageQuota',
								'StorageQuotaWarningLevel',
								'SiteCollectionQuota',
								'ResourceQuota',
								'ResourceQuotaWarningLevel',
								'StorageMaximumLevel',
								'StorageWarningLevel',
								'UserCodeMaximumLevel',
								'UserCodeWarningLevel',
								'NewValue',
								'OldValue',
								'Value',
								'EventData'
							)

							foreach ($quotaPropName in $quotaPropertyNames) {
								if ($null -ne $auditEvent.PSObject.Properties[$quotaPropName] -and $null -ne $auditEvent.$quotaPropName) {
									[string]$propValue = [string]$auditEvent.$quotaPropName

									# Parse XML-style EventData content into individual properties.
									if ($quotaPropName -eq 'EventData' -and $propValue -match '<[A-Z_]+>') {
										[regex]$xmlTagPattern = [regex]::new('<([A-Z_]+)>([^<]*)</\1>')
										[System.Text.RegularExpressions.MatchCollection]$xmlMatches = $xmlTagPattern.Matches($propValue)
										foreach ($xmlMatch in $xmlMatches) {
											[string]$tagName  = $xmlMatch.Groups[1].Value
											[string]$tagValue = $xmlMatch.Groups[2].Value
											# Attempt to parse as number for cleaner JSON output.
											if ($tagValue -match '^\d+$') {
												$quotaHashtable[$tagName] = [long]$tagValue
											}
											else {
												$quotaHashtable[$tagName] = $tagValue
											}
										}
									}
									else {
										# For non-EventData properties, add directly.
										if ($propValue -match '^\d+$') {
											$quotaHashtable[$quotaPropName] = [long]$propValue
										}
										else {
											$quotaHashtable[$quotaPropName] = $propValue
										}
									}
								}
							}

							if ($quotaHashtable.Count -gt 0) {
								$quotaDetailsValue = ($quotaHashtable | ConvertTo-Json -Depth 3 -Compress)
							}
						}

					[pscustomobject]$record = [pscustomobject]@{
						CreationTime   = if ($null -ne $auditEvent.CreationTime) { [datetime]$auditEvent.CreationTime } else { [datetime]::MinValue }
						UserId         = $resolvedUserId
						Operation      = $operationValue
						SiteUrl        = $siteUrlValue
						SiteNameShort  = $siteNameShortValue
						ObjectId       = $objectIdValue
						FileExtension  = if ($null -ne $auditEvent.PSObject.Properties['SourceFileExtension']) { [string]$auditEvent.SourceFileExtension } else { [string]::Empty }
						Parameters     = $parametersValue
						QuotaDetails   = $quotaDetailsValue
					}

					[void]$resultBag.Add($record)
				}
			}
			catch {
				# Log parallel processing errors but continue with other URIs.
				Write-Warning -Message "Failed to process content blob: $auditDataUri. Error: $($_.Exception.Message)"
			}
		}
	}

	# Convert PSCustomObjects from parallel processing back to strongly-typed SharePointUsageRecord objects.
	[psobject[]]$rawRecords                    = $usageRecordsBag.ToArray()
	[System.Collections.Generic.List[SharePointUsageRecord]]$typedRecordsList = [System.Collections.Generic.List[SharePointUsageRecord]]::new()

	foreach ($rawRecord in $rawRecords) {
		[SharePointUsageRecord]$typedRecord = [SharePointUsageRecord]::new()
		$typedRecord.CreationTime           = [datetime]$rawRecord.CreationTime
		$typedRecord.UserId                 = [string]$rawRecord.UserId
		$typedRecord.Operation              = [string]$rawRecord.Operation
		$typedRecord.SiteUrl                = [string]$rawRecord.SiteUrl
		$typedRecord.SiteNameShort          = [string]$rawRecord.SiteNameShort
		$typedRecord.ObjectId               = [string]$rawRecord.ObjectId
		$typedRecord.FileExtension          = [string]$rawRecord.FileExtension
		$typedRecord.Parameters             = [string]$rawRecord.Parameters
		$typedRecord.QuotaDetails           = [string]$rawRecord.QuotaDetails
		[void]$typedRecordsList.Add($typedRecord)
	}

	# Deduplicate records: keep one record per hour when all other fields are identical.
	# Uses a HashSet with composite key (hour-truncated time + all other fields) for O(1) lookup.
	[System.Collections.Generic.HashSet[string]]$seenRecordsSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
	[System.Collections.Generic.List[SharePointUsageRecord]]$deduplicatedList = [System.Collections.Generic.List[SharePointUsageRecord]]::new()

	foreach ($typedRecord in $typedRecordsList) {
		# Truncate CreationTime to the hour (remove minutes, seconds, milliseconds).
		[datetime]$truncatedTime = [datetime]::new(
			$typedRecord.CreationTime.Year,
			$typedRecord.CreationTime.Month,
			$typedRecord.CreationTime.Day,
			$typedRecord.CreationTime.Hour,
			0,
			0,
			[System.DateTimeKind]::Utc
		)

		# Build composite key from truncated hour and all other fields.
		[string]$compositeKey = "$($truncatedTime.ToString('o'))|$($typedRecord.UserId)|$($typedRecord.Operation)|$($typedRecord.SiteUrl)|$($typedRecord.SiteNameShort)|$($typedRecord.ObjectId)|$($typedRecord.FileExtension)|$($typedRecord.Parameters)|$($typedRecord.QuotaDetails)"

		# Add only if this combination hasn't been seen before.
		if ($seenRecordsSet.Add($compositeKey)) {
			[void]$deduplicatedList.Add($typedRecord)
		}
	}

	[int]$duplicatesRemoved = $typedRecordsList.Count - $deduplicatedList.Count
	Write-Verbose -Message "Deduplication removed $duplicatesRemoved records (same user/operation/site within same hour)"

	[SharePointUsageRecord[]]$finalRecords = $deduplicatedList.ToArray()

	Write-Verbose -Message "Purview content items: $contentItemCount"
	Write-Verbose -Message "Content blob URIs processed: $($allAuditDataUris.Count)"
	Write-Verbose -Message "Final usage records after deduplication: $($finalRecords.Count)"

	if ($finalRecords.Count -eq 0) {
		return (New-DummySharePointUsageRecord)
	}

	return $finalRecords
}
catch {
	Write-Error -Message "Failed to retrieve SharePoint usage records from Purview logs. $($_.Exception.Message)"
	return (New-DummySharePointUsageRecord)
}
} # end function Get-SharePointUsageRecordsFromPurview

[SharePointUsageRecord[]]$sharePointUsageRecords = Get-SharePointUsageRecordsFromPurview -TenantIdentifier $TenantId -ApplicationClientId $ClientId -ApplicationClientSecret $ClientSecret -LookbackHours $HoursBack -UseNoProxy $BypassProxy -MaxRetries $MaxRetryCount -ThrottleLimit $ParallelThrottleLimit -TimeoutSeconds $RequestTimeoutSeconds
[System.Collections.ArrayList]$SharePointUsageArrayList = [System.Collections.ArrayList]::new()

foreach ($recordItem in $sharePointUsageRecords) {
	[psobject]$outputRecord = $recordItem | Select-Object -Property CreationTime, UserId, Operation, SiteNameShort, FileExtension, QuotaDetails
	[void]$SharePointUsageArrayList.Add($outputRecord)
}

# Build global variable with all Data Collection Endpoints, which can be viewed by Log Ingestion app
try {
    $global:AzDceDetails = Get-AzDceListAll -AzAppId $LogIngestAppId `
                                            -AzAppSecret $LogIngestAppSecret `
                                            -TenantId $TenantId `
                                            -Verbose:$EnableVerbose
}
catch {
    Write-Error "Failed to retrieve Data Collection Endpoints: $_"
}

# Build global variable with all Data Collection Rules, which can be viewed by Log Ingestion app
try {
    $global:AzDcrDetails = Get-AzDcrListAll -AzAppId $LogIngestAppId `
                                            -AzAppSecret $LogIngestAppSecret `
                                            -TenantId $TenantId `
                                            -Verbose:$EnableVerbose
}
catch {
    Write-Error "Failed to retrieve Data Collection Rules: $_"
}

# Transform SharePoint usage records into strongly typed objects for log ingestion
[System.Collections.Generic.List[PSCustomObject]]$TransformedDataList = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($sharePointRecord in $SharePointUsageArrayList) {
    [PSCustomObject]$ItemObject = [PSCustomObject]@{
        CreationTime  = $sharePointRecord.CreationTime
        UserId        = $sharePointRecord.UserId
        Operation     = $sharePointRecord.Operation
        SiteNameShort = $sharePointRecord.SiteNameShort
        FileExtension = $sharePointRecord.FileExtension
        QuotaDetails  = $sharePointRecord.QuotaDetails
    }
    $TransformedDataList.Add($ItemObject)
}

[array]$DataVariable = $TransformedDataList.ToArray()

# Add CollectionTime to existing array
$DataVariable = Add-CollectionTimeToAllEntriesInArray -Data $DataVariable -Verbose:$EnableVerbose

# Add Computer information to existing array
$DataVariable = Add-ColumnDataToAllEntriesInArray -Data $DataVariable -Column1Name Computer -Column1Data $Env:ComputerName

# Validate and fix schema data structure of source data
$DataVariable = ValidateFix-AzLogAnalyticsTableSchemaColumnNames -Data $DataVariable -Verbose:$EnableVerbose

# Align data structure with schema (requirement for Data Collection Rule)
$DataVariable = Build-DataArrayToAlignWithSchema -Data $DataVariable -Verbose:$EnableVerbose

if ($DataVariable) {
    try {
        $ResultManagement = CheckCreateUpdate-TableDcr-Structure -AzLogWorkspaceResourceId $LogAnalyticsWorkspaceResourceId `
                                                                 -AzAppId $LogIngestAppId `
                                                                 -AzAppSecret $LogIngestAppSecret `
                                                                 -TenantId $TenantId `
                                                                 -DceName $AzDceName `
                                                                 -DcrName $AzDcrName `
                                                                 -DcrResourceGroup $AzDcrResourceGroup `
                                                                 -TableName $TableName `
                                                                 -Data $DataVariable `
                                                                 -LogIngestServicePricipleObjectId $AzDcrLogIngestServicePrincipalObjectId `
                                                                 -AzDcrSetLogIngestApiAppPermissionsDcrLevel $AzDcrSetLogIngestApiAppPermissionsDcrLevel `
                                                                 -AzLogDcrTableCreateFromAnyMachine $AzLogDcrTableCreateFromAnyMachine `
                                                                 -AzLogDcrTableCreateFromReferenceMachine $AzLogDcrTableCreateFromReferenceMachine
    }
    catch {
        Write-Error "Failed to check or update Table and Data Collection Rule structure: $_"
    }
} # if $DataVariable

if ($DataVariable) {
    try {
        $ResultPost = Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $AzDceName `
                                                                        -DcrName $AzDcrName `
                                                                        -Data $DataVariable `
                                                                        -TableName $TableName `
                                                                        -AzAppId $LogIngestAppId `
                                                                        -AzAppSecret $LogIngestAppSecret `
                                                                        -TenantId $TenantId `
                                                                        -BatchAmount 100 `
                                                                        -Verbose:$EnableVerbose
    }
    catch {
        Write-Error "Failed to post data to Log Analytics via Data Collection Rule: $_"
    }
}

[string]$EventLogSource = "Publish-SharePointUsageInformationToLAW"

# Ensure the event log source exists in the Application log.
# Avoid [System.Diagnostics.EventLog]::SourceExists() because it scans ALL logs
# including Security, which fails without elevation.
try {
    [string]$registryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\EventLog\Application\$EventLogSource"
    if (-not (Test-Path -Path $registryPath)) {
        [System.Diagnostics.EventLog]::CreateEventSource($EventLogSource, "Application")
    }
}
catch {
    Write-Warning "Unable to verify or create event log source '$EventLogSource': $($_.Exception.Message)"
}

if ($Error.Count -gt 0) {
    [string]$ErrorSummary = "The following non-terminating errors occurred during script execution:`r`n"
    foreach ($ErrorRecord in $Error) {
        [string]$ErrorMessage = $ErrorRecord.ToString()
        $ErrorSummary += "$ErrorMessage`r`n"
        Write-Warning $ErrorMessage
    }

    try {
        Write-EventLog -LogName "Application" `
                       -Source $EventLogSource `
                       -EntryType Error `
                       -EventId 1011 `
                       -Message $ErrorSummary
    }
    catch {
        Write-Warning "Failed to write errors to the event log: $($_.Exception.Message)"
    }

    $Error.Clear()
}


#endregion Log Analytics Ingestion
