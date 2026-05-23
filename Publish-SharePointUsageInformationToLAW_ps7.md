## How User Activity Gets Into Log Analytics

While `Publish-SharePointSiteInfoToLaw.ps1` handles site-level metadata and storage, this script — `Publish-SharePointUsageInformationToLAW_ps7.ps1` — captures what users are actually *doing*: accessing files, downloading, uploading, and viewing pages.

It runs **hourly** and pulls from a completely different source: the **Microsoft Purview Management Activity API** (Office 365 audit logs) rather than the Graph reports endpoint.

### What It Collects

Each run collects the last hour of `Audit.SharePoint` events and filters for seven operations:

| Operation | What it represents |
|---|---|
| `FileAccessed` | A user opened or previewed a file |
| `FileDownloaded` | A file was downloaded |
| `FileUploaded` | A file was uploaded |
| `PageViewed` | A SharePoint page was viewed |
| `SetSiteProperties` | Admin changed site settings |
| `SiteCollectionQuotaModified` | Storage quota was changed |
| `SiteStorageLimitModeChanged` | Storage limit mode was changed |

OneDrive and personal site events are excluded at collection time using a pre-compiled regex — so only team site and project site activity lands in the table.

### How It Works

**1. Subscription check**
Before querying, the script verifies the `Audit.SharePoint` subscription is active. If not, it starts it automatically.

**2. Content blob discovery**
The API returns pointers to content blobs rather than events directly. The script collects all blob URIs for the time window first, then fetches them in parallel.

**3. Parallel blob fetching**
Content blobs are processed using `ForEach-Object -Parallel` with a configurable throttle limit. Each parallel runspace applies the operation filter (via a `HashSet` for O(1) lookup) and the OneDrive filter before adding records to a thread-safe `ConcurrentBag`.

**4. Deduplication**
After collection, records are deduplicated using a composite key that truncates `CreationTime` to the hour. This avoids duplicate entries when the script overlaps with a previous run.

**5. Strong typing + ingestion**
`PSCustomObject` results from the parallel runspaces are converted back to strongly-typed `SharePointUsageRecord` objects, then pushed to Log Analytics via the same DCR/DCE pipeline used by the site information script.

### Key Design Choices

| Decision | Reason |
|---|---|
| Purview audit API instead of Graph reports | Provides raw per-event data with user identity, file path, and operation — the reports API only gives aggregated counts |
| Hourly schedule | Keeps latency low; audit content blobs are typically available within minutes |
| Parallel blob fetching | Audit logs for an active tenant can produce hundreds of blobs per hour; sequential fetching would be too slow |
| HashSet for operation filtering | Avoids O(n) linear scan on every event — important at scale |
| Exponential backoff with jitter | The Purview API is rate-limited; random jitter prevents retry storms when multiple runbooks execute simultaneously |
| Registry-based secret retrieval | Script runs on an on-premises Hybrid Worker where Azure Key Vault connectivity may not be available; registry provides a local secure store |
| Hour-level deduplication | Purview content blobs can overlap between API poll windows; deduplication prevents inflated event counts in Log Analytics |


## Blog description:

This PowerShell script serves as the user activity data collection layer for the SharePoint Online dashboard. It runs hourly as an Azure Automation runbook with a Hybrid Worker and captures raw per-event audit data — file accesses, downloads, uploads, page views, and site configuration changes — across the entire SharePoint Online tenant.

The script authenticates against the Microsoft Purview Management Activity API and verifies that the `Audit.SharePoint` subscription is active, starting it automatically if not. Rather than querying events directly, the API exposes content blob URIs for the requested time window; the script collects all blob pointers first, then fetches them in parallel using `ForEach-Object -Parallel` with a configurable throttle limit. Each parallel runspace applies an operation filter using a `HashSet` for O(1) lookup and excludes OneDrive and personal site events via a pre-compiled regex, ensuring only team and project site activity is retained. After collection, records are deduplicated using a composite key truncated to the hour to prevent overlap between consecutive runs.

All collected data is normalised into a flat schema aligned with a Data Collection Rule (DCR) and ingested into Log Analytics via a Data Collection Endpoint (DCE). The table schema for `sharepointusageinfo_CL` is automatically created or updated if the data structure changes, making per-event activity data directly available for KQL queries and dashboard visualisations without manual table management.
