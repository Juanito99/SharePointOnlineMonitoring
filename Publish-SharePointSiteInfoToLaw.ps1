
#requires -Version 7.0
<#
.SYNOPSIS
    Retrieves SharePoint site information including usage details, file counts, and display names
    via Microsoft Graph API.

.DESCRIPTION
    This script authenticates against Microsoft Graph using client credentials (application permissions)
    and retrieves SharePoint site usage details from the reports endpoint. It enriches the usage data
    with display names from the sites API and exports the combined information to a CSV file.

    The following data is collected per site:
    - Display name and URL-derived name
    - Site identifier and deletion status
    - Last activity date
    - File count and active file count
    - Page view count and visited page count
    - Storage used and allocated (in bytes)
    - Report refresh date

.PARAMETER OutputFile
    The file path for the exported CSV output. Defaults to "SharePointSitesInfo.csv".

.PARAMETER Days
    The number of days to look back for activity data. Valid values are 7, 30, 90, or 180.
    Defaults to 7.

.OUTPUTS
    CSV file containing an array of objects with properties: siteName, siteNameFromUrl, siteId,
    isDeleted, lastActivityDate, fileCount, activeFileCount, pageViewCount, storageUsedInBytes,
    storageAllocatedInBytes, visitedPageCount, reportRefreshDate.

.EXAMPLE
    .\Get-SharePointSitesInformation.ps1 -Days 30 -OutputFile "SharePointReport.csv"

.NOTES
    Author          : Ruben
    Date            : 2026-03-14
    Version         : 2.0
    Prerequisites   : Azure AD app registration with Reports.Read.All and Sites.Read.All
                      application permissions.

.VERSIONPREFERENCE
    PowerShell 7.x
#>
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputFile = 'SharePointSitesInfo.csv',

    [Parameter(Mandatory = $false)]
    [ValidateSet(7, 30, 90, 180)]
    [int]$Days = 7
)

Connect-AzAccount -Identity

#region Configuration

[bool]$EnableVerbose                                              = $true
[string]$VerbosePreference                                        = "SilentlyContinue"  # "Continue"

[string]$TenantId                                                 = "<your-azure-ad-tenant-id>"

# Azure Key Vault - Log Ingestion Application
[string]$AzureKeyVaultName                                        = "<your-key-vault-name>"
[string]$AzureKeyVaultSecretName                                  = "<your-log-ingestion-app-secret-name>"

[object]$AzureKeyVaultSecret                                      = Get-AzKeyVaultSecret -VaultName $AzureKeyVaultName -Name $AzureKeyVaultSecretName
[System.Security.SecureString]$AzureKeyVaultSecretValue           = $AzureKeyVaultSecret.SecretValue
[string]$LogIngestAppSecret                                       = $AzureKeyVaultSecretValue | ConvertFrom-SecureString -AsPlainText

[string]$ClientSecret                                             = Get-AzKeyVaultSecret -VaultName $AzureKeyVaultName -Name '<your-graph-app-client-secret-name>' |
                                                                        Select-Object -ExpandProperty SecretValue |
                                                                        ConvertFrom-SecureString -AsPlainText

[string]$ClientId                                                 = Get-AzKeyVaultSecret -VaultName $AzureKeyVaultName -Name '<your-graph-app-client-id-secret-name>' |
                                                                        Select-Object -ExpandProperty SecretValue |
                                                                        ConvertFrom-SecureString -AsPlainText

[string]$LogIngestAppId                                           = "<your-log-ingestion-app-client-id>"
[string]$AzDcrLogIngestServicePrincipalObjectId                    = "<your-log-ingestion-app-service-principal-object-id>"

[string]$LogAnalyticsWorkspaceResourceId                          = "/subscriptions/<your-subscription-id>/resourceGroups/<your-resource-group>/providers/Microsoft.OperationalInsights/workspaces/<your-law-name>"

[string]$AzDcrResourceGroup                                       = "<your-dcr-resource-group>"
[bool]$AzDcrSetLogIngestApiAppPermissionsDcrLevel                  = $false
[array]$AzLogDcrTableCreateFromReferenceMachine                   = @()
[bool]$AzLogDcrTableCreateFromAnyMachine                          = $true

[string]$AzDceName                                                = "<your-data-collection-endpoint-name>"
[string]$TableName                                                = "sharepointsiteinformation"
[string]$AzDcrName                                                = "dcr-prd-" + $TableName + "_CL"

#endregion Configuration

#region Functions

# Acquires an OAuth2 access token from Azure Active Directory using client credentials
function Get-GraphAccessToken {
    <#
    .SYNOPSIS
        Acquires a Microsoft Graph API access token using client credentials flow.

    .DESCRIPTION
        Authenticates against the Microsoft identity platform token endpoint using the client
        credentials grant type and returns a bearer access token for Microsoft Graph API calls.

    .PARAMETER ClientId
        The Azure Active Directory application (client) identifier.

    .PARAMETER ClientSecret
        The client secret for the Azure Active Directory application.

    .PARAMETER TenantId
        The Azure Active Directory tenant identifier.

    .OUTPUTS
        [string] The bearer access token, or [string]::Empty on failure.

    .EXAMPLE
        [string]$Token = Get-GraphAccessToken -ClientId $Id -ClientSecret $Secret -TenantId $Tenant

    .NOTES
        Author  : yourname
        Version : 2.0

    .VERSIONPREFERENCE
        PowerShell 7.x
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientId,

        [Parameter(Mandatory = $true)]
        [string]$ClientSecret,

        [Parameter(Mandatory = $true)]
        [string]$TenantId
    )

    [string]$TokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

    [hashtable]$RequestBody = @{
        grant_type    = "client_credentials"
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "https://graph.microsoft.com/.default"
    }

    try {
        [PSObject]$TokenResponse = Invoke-RestMethod -Uri $TokenUrl `
            -Method Post `
            -Body $RequestBody `
            -ContentType "application/x-www-form-urlencoded"
        return [string]$TokenResponse.access_token
    }
    catch {
        [string]$StatusCode = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { 'N/A' }
        [string]$ErrorBody  = if ($_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
        Write-Error "Failed to obtain access token.`n  TenantId : $TenantId`n  ClientId : $ClientId`n  Endpoint : $TokenUrl`n  Status   : $StatusCode`n  Detail   : $ErrorBody"
        return [string]::Empty
    }
} # end function Get-GraphAccessToken

# Retrieves the SharePoint site usage detail report as a CSV string from Microsoft Graph
function Get-SharePointSiteUsageDetail {
    <#
    .SYNOPSIS
        Retrieves the SharePoint site usage detail report from Microsoft Graph.

    .DESCRIPTION
        Calls the getSharePointSiteUsageDetail report endpoint which returns per-site usage data
        including storage, file counts, page views, and activity dates as a CSV-formatted string.

    .PARAMETER AccessToken
        A valid Microsoft Graph API bearer access token.

    .PARAMETER Days
        The reporting period in days. Valid values are 7, 30, 90, or 180.

    .OUTPUTS
        [string] CSV-formatted response from the Graph reports endpoint, or [string]::Empty on failure.

    .EXAMPLE
        [string]$CsvData = Get-SharePointSiteUsageDetail -AccessToken $Token -Days 7

    .NOTES
        Author  : yourname
        Version : 2.0

    .VERSIONPREFERENCE
        PowerShell 7.x
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$AccessToken,

        [Parameter(Mandatory = $false)]
        [ValidateSet(7, 30, 90, 180)]
        [int]$Days = 7
    )

    [hashtable]$RequestHeaders = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type"  = "application/json"
    }

    [string]$RequestUri = "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='D$Days')"

    try {
        Write-Verbose "Fetching SharePoint site usage details from: $RequestUri"
        [string]$Response = Invoke-RestMethod -Uri $RequestUri -Method Get -Headers $RequestHeaders
        return $Response
    }
    catch {
        [string]$StatusCode = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { 'N/A' }
        [string]$ErrorBody  = if ($_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
        Write-Error "Failed to retrieve SharePoint site usage details.`n  URI    : $RequestUri`n  Period : D$Days`n  Status : $StatusCode`n  Detail : $ErrorBody"
        return [string]::Empty
    }
} # end function Get-SharePointSiteUsageDetail

# Retrieves all SharePoint sites with display names via Microsoft Graph sites API with pagination
function Get-AllSharePointSites {
    <#
    .SYNOPSIS
        Retrieves all SharePoint sites with their display names and web URLs.

    .DESCRIPTION
        Calls the Microsoft Graph getAllSites endpoint with pagination to retrieve the complete list
        of SharePoint sites in the tenant. Returns site identifier, display name, and web URL for
        each site. On error or empty result, returns an array with a single dummy object containing
        all expected properties.

    .PARAMETER AccessToken
        A valid Microsoft Graph API bearer access token.

    .OUTPUTS
        [PSCustomObject[]] Array of objects with properties: id, displayName, webUrl.

    .EXAMPLE
        [PSCustomObject[]]$Sites = Get-AllSharePointSites -AccessToken $Token

    .NOTES
        Author  : yourname
        Version : 2.0

    .VERSIONPREFERENCE
        PowerShell 7.x
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$AccessToken
    )

    [hashtable]$RequestHeaders = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type"  = "application/json"
    }

    [System.Collections.Generic.List[PSObject]]$AllSites = [System.Collections.Generic.List[PSObject]]::new()
    [string]$RequestUri                                  = "https://graph.microsoft.com/v1.0/sites/getAllSites?`$select=id,displayName,webUrl&`$top=1000"

    try {
        while ($RequestUri) {
            [PSObject]$Response = Invoke-RestMethod -Uri $RequestUri -Method Get -Headers $RequestHeaders
            $RequestUri         = $Response.'@odata.nextLink'

            if ($Response.value) {
                foreach ($SiteEntry in $Response.value) {
                    $AllSites.Add($SiteEntry)
                }
            }
        }
    }
    catch {
        [string]$StatusCode = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { 'N/A' }
        [string]$ErrorBody  = if ($_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
        Write-Warning "Failed to retrieve SharePoint sites list.`n  URI         : $RequestUri`n  Sites so far: $($AllSites.Count)`n  Status      : $StatusCode`n  Detail      : $ErrorBody"
    }

    if ($AllSites.Count -eq 0) {
        [PSCustomObject]$DummySite = [PSCustomObject]@{
            id          = [string]::Empty
            displayName = [string]::Empty
            webUrl      = [string]::Empty
        }
        return @($DummySite)
    }

    return $AllSites.ToArray()
} # end function Get-AllSharePointSites

# Converts a raw CSV string as returned by the Graph reports API into an array of PSObjects
function ConvertFrom-CsvString {
    <#
    .SYNOPSIS
        Converts a raw CSV string into an array of PowerShell objects.

    .DESCRIPTION
        Strips the UTF-8 BOM that the Microsoft Graph reports API prepends to CSV responses and
        parses the content into PSObjects using the built-in ConvertFrom-Csv cmdlet.
        Returns an empty array if the input is null or whitespace.

    .PARAMETER CsvContent
        The raw CSV content string to parse.

    .OUTPUTS
        [PSObject[]] Array of parsed objects, or an empty array if the input is empty.

    .EXAMPLE
        [PSObject[]]$ParsedData = ConvertFrom-CsvString -CsvContent $RawCsv

    .NOTES
        Author  : yourname
        Version : 2.0

    .VERSIONPREFERENCE
        PowerShell 7.x
    #>
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$CsvContent
    )

    $CsvContent = $CsvContent -replace '^\xEF\xBB\xBF', ''
    $CsvContent = $CsvContent.TrimStart([char]0xFEFF)

    if ([string]::IsNullOrWhiteSpace($CsvContent)) {
        return @()
    }

    try {
        [PSObject[]]$ParsedObjects = $CsvContent | ConvertFrom-Csv
        return $ParsedObjects
    }
    catch {
        [int]$PreviewLength      = [System.Math]::Min(200, $CsvContent.Length)
        [string]$ContentPreview = $CsvContent.Substring(0, $PreviewLength) -replace '[\r\n]+', ' '
        Write-Error "Failed to parse CSV content.`n  Exception    : $($_.Exception.GetType().FullName)`n  Message      : $($_.Exception.Message)`n  ContentLength: $($CsvContent.Length)`n  Preview      : $ContentPreview"
        return @()
    }
} # end function ConvertFrom-CsvString

# Transforms and merges usage detail data with site display names into the final output format
function ConvertTo-SharePointSiteInformation {
    <#
    .SYNOPSIS
        Transforms SharePoint usage detail data and site metadata into a unified site information
        array.

    .DESCRIPTION
        Takes the parsed usage detail objects from the reports API and the full sites list from the
        sites API, builds internal lookup tables for display names and web URLs, and produces a
        combined output array with all relevant site information properties.
        On error or empty input, returns an array with a single dummy object containing all expected
        properties populated with safe defaults.

    .PARAMETER UsageDetailList
        Array of PSObjects parsed from the getSharePointSiteUsageDetail report CSV response.

    .PARAMETER SitesList
        Array of PSObjects from the getAllSites endpoint containing id, displayName, and webUrl.

    .OUTPUTS
        [PSCustomObject[]] Array of objects with properties: siteName, siteNameFromUrl, siteId,
        isDeleted, lastActivityDate, fileCount, activeFileCount, pageViewCount, storageUsedInBytes,
        storageAllocatedInBytes, visitedPageCount, reportRefreshDate.

    .EXAMPLE
        [PSCustomObject[]]$Result = ConvertTo-SharePointSiteInformation `
            -UsageDetailList $UsageData -SitesList $AllSites

    .NOTES
        Author  : yourname
        Version : 2.0

    .VERSIONPREFERENCE
        PowerShell 7.x
    #>
    param(
        [Parameter(Mandatory = $true)]
        [PSObject[]]$UsageDetailList,

        [Parameter(Mandatory = $true)]
        [PSObject[]]$SitesList
    )

    # Dummy object template for error or empty result scenarios
    [PSCustomObject]$DummySiteInformation = [PSCustomObject]@{
        siteName                = [string]::Empty
        siteNameFromUrl         = [string]::Empty
        siteId                  = [string]::Empty
        isDeleted               = $false
        lastActivityDate        = [datetime]::MinValue
        fileCount               = [int]0
        activeFileCount         = [int]0
        pageViewCount           = [int]0
        storageUsedInBytes      = [int64]0
        storageAllocatedInBytes = [int64]0
        visitedPageCount        = [int]0
        reportRefreshDate       = [string]::Empty
    }

    if ($UsageDetailList.Count -eq 0) {
        Write-Warning "No usage detail data provided for transformation."
        return @($DummySiteInformation)
    }

    # Build lookup hashtables keyed by site GUID for fast name and URL resolution
    [hashtable]$SiteNameLookup   = @{}
    [hashtable]$SiteWebUrlLookup = @{}

    foreach ($SiteEntry in $SitesList) {
        [string]$GraphSiteId = [string]$SiteEntry.id
        if ([string]::IsNullOrEmpty($GraphSiteId)) {
            continue
        }

        # Graph site id format: "contoso.sharepoint.com,guid1,guid2" — first GUID is the site collection id
        if ($GraphSiteId -match ',([0-9a-fA-F-]{36}),') {
            [string]$ShortId                = $Matches[1].ToLowerInvariant()
            $SiteNameLookup[$ShortId]   = [string]$SiteEntry.displayName
            $SiteWebUrlLookup[$ShortId] = [string]$SiteEntry.webUrl
        }
        # Also key by full id in case the report uses that format
        $SiteNameLookup[$GraphSiteId.ToLowerInvariant()]   = [string]$SiteEntry.displayName
        $SiteWebUrlLookup[$GraphSiteId.ToLowerInvariant()] = [string]$SiteEntry.webUrl
    }

    # Transform each usage detail row into the output format
    [System.Collections.Generic.List[PSCustomObject]]$TransformedSites = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($UsageDetailEntry in $UsageDetailList) {
        try {
            # Extract site URL using a case-insensitive property name match
            [string]$SiteUrl          = [string]::Empty
            [PSObject]$SiteUrlProperty = $UsageDetailEntry.PSObject.Properties |
                Where-Object {
                    ($_.Name -replace '[^A-Za-z0-9]', '').ToLowerInvariant() -eq 'siteurl'
                } |
                Select-Object -First 1

            if ($SiteUrlProperty -and $SiteUrlProperty.Value) {
                $SiteUrl = [string]$SiteUrlProperty.Value
            }

            # Parse last activity date safely
            [datetime]$LastActivityDate = [datetime]::MinValue
            if ($UsageDetailEntry.'Last Activity Date' -and
                $UsageDetailEntry.'Last Activity Date' -ne 'False' -and
                $UsageDetailEntry.'Last Activity Date' -ne '') {
                try {
                    $LastActivityDate = [datetime]::Parse($UsageDetailEntry.'Last Activity Date')
                }
                catch {
                    Write-Verbose "Could not parse date '$($UsageDetailEntry.'Last Activity Date')' for site $SiteUrl"
                }
            }

            # Resolve site identifier from the usage detail row
            [string]$SiteIdentifier = if ($UsageDetailEntry.'Site Id' -and
                $UsageDetailEntry.'Site Id' -ne '') {
                $UsageDetailEntry.'Site Id'
            }
            else {
                [string]::Empty
            }

            # Look up display name from the sites list using the site GUID
            [string]$SiteName = [string]::Empty
            if ($SiteIdentifier -ne '' -and
                $SiteNameLookup.ContainsKey($SiteIdentifier.ToLowerInvariant())) {
                $SiteName = $SiteNameLookup[$SiteIdentifier.ToLowerInvariant()]
            }

            # Derive a fallback name from the web URL returned by the sites API
            [string]$SiteNameFromUrl = [string]::Empty
            [string]$WebUrl          = [string]::Empty
            if ($SiteIdentifier -ne '' -and
                $SiteWebUrlLookup.ContainsKey($SiteIdentifier.ToLowerInvariant())) {
                $WebUrl = $SiteWebUrlLookup[$SiteIdentifier.ToLowerInvariant()]
            }
            if ($WebUrl -ne '') {
                [string]$UrlPath = ([System.Uri]$WebUrl).AbsolutePath.TrimEnd('/')
                if ($UrlPath -ne '' -and $UrlPath -ne '/') {
                    $SiteNameFromUrl = [System.Uri]::UnescapeDataString($UrlPath.Split('/')[-1])
                }
            }

            # File and page counts are included per-site in the usage detail response
            [int]$FileCount       = if ($UsageDetailEntry.'File Count' -and $UsageDetailEntry.'File Count' -ne '') { [int]$UsageDetailEntry.'File Count' } else { 0 }
            [int]$ActiveFileCount = if ($UsageDetailEntry.'Active File Count' -and $UsageDetailEntry.'Active File Count' -ne '') { [int]$UsageDetailEntry.'Active File Count' } else { 0 }
            [int]$PageViewCount   = if ($UsageDetailEntry.'Page View Count' -and $UsageDetailEntry.'Page View Count' -ne '') { [int]$UsageDetailEntry.'Page View Count' } else { 0 }

            [PSCustomObject]$SiteInformation = [PSCustomObject]@{
                siteName                = $SiteName
                siteNameFromUrl         = $SiteNameFromUrl
                siteId                  = $SiteIdentifier
                isDeleted               = ($UsageDetailEntry.'Is Deleted' -eq 'True')
                lastActivityDate        = $LastActivityDate
                fileCount               = $FileCount
                activeFileCount         = $ActiveFileCount
                pageViewCount           = $PageViewCount
                storageUsedInBytes      = if ($UsageDetailEntry.'Storage Used (Byte)' -and $UsageDetailEntry.'Storage Used (Byte)' -ne '') { [int64]$UsageDetailEntry.'Storage Used (Byte)' } else { [int64]0 }
                storageAllocatedInBytes = if ($UsageDetailEntry.'Storage Allocated (Byte)' -and $UsageDetailEntry.'Storage Allocated (Byte)' -ne '') { [int64]$UsageDetailEntry.'Storage Allocated (Byte)' } else { [int64]0 }
                visitedPageCount        = if ($UsageDetailEntry.'Visited Page Count' -and $UsageDetailEntry.'Visited Page Count' -ne '') { [int]$UsageDetailEntry.'Visited Page Count' } else { [int]0 }
                reportRefreshDate       = if ($UsageDetailEntry.'Report Refresh Date') { [string]$UsageDetailEntry.'Report Refresh Date' } else { [string]::Empty }
            }

            $TransformedSites.Add($SiteInformation)
        }
        catch {
            Write-Warning "Error processing usage detail row: $_"
            continue
        }
    }

    if ($TransformedSites.Count -eq 0) {
        Write-Warning "No site information was produced during transformation."
        return @($DummySiteInformation)
    }

    return $TransformedSites.ToArray()
} # end function ConvertTo-SharePointSiteInformation

#endregion Functions

#region Main Script Execution

try {
    Write-Host "Starting SharePoint Sites Information Retrieval via Reports Endpoint..." -ForegroundColor Green

    # Validate parameters
    if (-not $ClientId -or -not $ClientSecret -or -not $TenantId) {
        Write-Error "ClientId, ClientSecret, and TenantId are required."
        exit 1
    }

    # Acquire access token
    Write-Host "Acquiring access token..." -ForegroundColor Cyan
    [string]$AccessToken = Get-GraphAccessToken -ClientId $ClientId -ClientSecret $ClientSecret -TenantId $TenantId

    if ($AccessToken -eq [string]::Empty) {
        Write-Error "Failed to acquire access token. Aborting script execution."
        exit 1
    }

    # Retrieve SharePoint site usage detail report
    Write-Host "Retrieving SharePoint site usage details (last $Days days)..." -ForegroundColor Cyan
    [string]$UsageDetailCsvContent = Get-SharePointSiteUsageDetail -AccessToken $AccessToken -Days $Days

    if ($UsageDetailCsvContent -eq [string]::Empty) {
        Write-Error "No usage detail data returned from the Graph API. Aborting script execution."
        exit 1
    }

    # Parse usage detail CSV into objects
    Write-Host "Processing usage data..." -ForegroundColor Cyan
    [PSObject[]]$UsageDetailList = ConvertFrom-CsvString -CsvContent $UsageDetailCsvContent

    if ($UsageDetailList.Count -eq 0) {
        Write-Warning "No SharePoint site usage data found after parsing."
        exit 0
    }

    Write-Host "Parsed $($UsageDetailList.Count) site usage record(s)." -ForegroundColor Cyan

    # Show actual column names returned by the API for diagnostics
    [string[]]$ColumnNames = $UsageDetailList[0].PSObject.Properties.Name
    Write-Host "API columns: $($ColumnNames -join ' | ')" -ForegroundColor Yellow

    # Retrieve all SharePoint sites for display name enrichment
    Write-Host "Fetching SharePoint site display names..." -ForegroundColor Cyan
    [PSCustomObject[]]$AllSites = Get-AllSharePointSites -AccessToken $AccessToken
    Write-Host "Retrieved $($AllSites.Count) site(s) with display names." -ForegroundColor Cyan

    # Transform and merge usage data with site metadata
    Write-Host "Transforming and merging site information..." -ForegroundColor Cyan
    [PSCustomObject[]]$SharePointSiteResults = ConvertTo-SharePointSiteInformation `
        -UsageDetailList $UsageDetailList `
        -SitesList $AllSites

    # Export results to CSV
    if ($SharePointSiteResults.Count -gt 0) {
        $SharePointSiteResults |
            Export-Csv -Path $OutputFile -NoTypeInformation -Force
        Write-Host "Results exported to: $OutputFile" -ForegroundColor Green
        Write-Host "Total sites processed: $($SharePointSiteResults.Count)" -ForegroundColor Green
    }
    else {
        Write-Warning "No site information was collected."
    }
}
catch {
    [string]$StatusCode = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { 'N/A' }
    [string]$ErrorBody  = if ($_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
    Write-Error "An unexpected error occurred during SharePoint site information retrieval.`n  Exception : $($_.Exception.GetType().FullName)`n  Message   : $ErrorBody`n  Status    : $StatusCode`n  Line      : $($_.InvocationInfo.ScriptLineNumber)`n  Command   : $($_.InvocationInfo.Line.Trim())"
    exit 1
}

#endregion Main Script Execution

#region Log Analytics Ingestion

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

# Transform SharePoint site results into strongly typed objects for log ingestion
[System.Collections.Generic.List[PSCustomObject]]$TransformedDataList = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($SiteResultItem in $SharePointSiteResults) {
    [PSCustomObject]$ItemObject = [PSCustomObject]@{
        siteName                = [string]$SiteResultItem.siteName
        siteNameFromUrl         = [string]$SiteResultItem.siteNameFromUrl
        siteId                  = [string]$SiteResultItem.siteId
        isDeleted               = [bool]$SiteResultItem.isDeleted
        lastActivityDate        = [datetime]$SiteResultItem.lastActivityDate
        fileCount               = [int]$SiteResultItem.fileCount
        activeFileCount         = [int]$SiteResultItem.activeFileCount
        pageViewCount           = [int]$SiteResultItem.pageViewCount
        storageUsedInBytes      = [long]$SiteResultItem.storageUsedInBytes
        storageAllocatedInBytes = [long]$SiteResultItem.storageAllocatedInBytes
        visitedPageCount        = [int]$SiteResultItem.visitedPageCount
        reportRefreshDate       = [string]$SiteResultItem.reportRefreshDate
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
        # The DCE enforces a 1 MB per-request hard limit.
        # Split the data into size-measured batches so each POST stays safely under that limit.
        [int]$MaxBatchBytes                                     = 900 * 1024   # 900 KB — leaves headroom under the 1 MB DCE limit
        [System.Collections.Generic.List[object]]$CurrentBatch = [System.Collections.Generic.List[object]]::new()
        [int]$CurrentBatchBytes                                 = 0
        [int]$BatchIndex                                        = 0

        foreach ($Record in $DataVariable) {
            [int]$RecordBytes = [System.Text.Encoding]::UTF8.GetByteCount(($Record | ConvertTo-Json -Compress -Depth 5))

            # Flush the current batch before it would exceed the size limit
            if ($CurrentBatch.Count -gt 0 -and ($CurrentBatchBytes + $RecordBytes) -gt $MaxBatchBytes) {
                $BatchIndex++
                Write-Host "Posting batch $BatchIndex ($($CurrentBatch.Count) records, $([math]::Round($CurrentBatchBytes / 1KB, 1)) KB)..." -ForegroundColor Cyan
                Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $AzDceName `
                                                                   -DcrName $AzDcrName `
                                                                   -Data $CurrentBatch.ToArray() `
                                                                   -TableName $TableName `
                                                                   -AzAppId $LogIngestAppId `
                                                                   -AzAppSecret $LogIngestAppSecret `
                                                                   -TenantId $TenantId `
                                                                   -BatchAmount $CurrentBatch.Count `
                                                                   -Verbose:$EnableVerbose | Out-Null
                $CurrentBatch.Clear()
                $CurrentBatchBytes = 0
            }

            $CurrentBatch.Add($Record)
            $CurrentBatchBytes += $RecordBytes
        }

        # Post any remaining records
        if ($CurrentBatch.Count -gt 0) {
            $BatchIndex++
            Write-Host "Posting batch $BatchIndex ($($CurrentBatch.Count) records, $([math]::Round($CurrentBatchBytes / 1KB, 1)) KB)..." -ForegroundColor Cyan
            Post-AzLogAnalyticsLogIngestCustomLogDcrDce-Output -DceName $AzDceName `
                                                               -DcrName $AzDcrName `
                                                               -Data $CurrentBatch.ToArray() `
                                                               -TableName $TableName `
                                                               -AzAppId $LogIngestAppId `
                                                               -AzAppSecret $LogIngestAppSecret `
                                                               -TenantId $TenantId `
                                                               -BatchAmount $CurrentBatch.Count `
                                                               -Verbose:$EnableVerbose | Out-Null
        }

        Write-Host "Log Analytics ingestion complete. Total batches posted: $BatchIndex" -ForegroundColor Green
    }
    catch {
        [string]$StatusCode = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { 'N/A' }
        [string]$ErrorBody  = if ($_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
        Write-Error "Failed to post data to Log Analytics via Data Collection Rule.`n  Batch    : $BatchIndex`n  Status   : $StatusCode`n  Detail   : $ErrorBody"
    }
}

if ($Error.Count -gt 0) {
    Write-Warning "The following errors occurred during script execution:"
    foreach ($ErrorRecord in $Error) {
        Write-Warning $ErrorRecord.ToString()
    }
    $Error.Clear()
}

#endregion Log Analytics Ingestion
