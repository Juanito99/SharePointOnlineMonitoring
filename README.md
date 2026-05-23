## How the Data Gets Into Log Analytics

Before the KQL queries can do anything, data needs to land in your Log Analytics workspace. That's what `Publish-SharePointSiteInfoToLaw.ps1` handles.

The script runs on a schedule (via Azure Automation with a Hybrid Worker) and has two phases: collect and ingest.

### Phase 1 — Collect from Microsoft Graph

It authenticates using **Managed Identity** — no credentials in the script. Secrets needed for the log ingestion app are pulled from **Azure Key Vault** at runtime.

Two Graph API calls are made:

- **`getSharePointSiteUsageDetail`** — the reports endpoint returns a per-site CSV with storage, file counts, page views, and last activity date. You can configure the lookback period: 7, 30, 90, or 180 days.
- **`getAllSites`** — paginates through all sites to retrieve display names and web URLs. This is necessary because the reports endpoint only returns site GUIDs, not readable names.

The script merges both datasets using the site GUID as the join key, producing a clean object per site with human-readable names alongside all the usage metrics.

> One thing worth noting: Graph's CSV responses include a UTF-8 BOM. The script strips that before parsing — a small detail that trips up a lot of implementations.

### Phase 2 — Ingest into Log Analytics

Once the data is shaped, it's pushed to Log Analytics using the **Data Collection Rule (DCR) / Data Collection Endpoint (DCE)** ingestion pipeline via the `AzLogDcrIngestPS` module.

What that means in practice:
- The table schema (`sharepointsiteinformation_CL`) is **automatically created or updated** if the data structure changes — no manual table management.
- Data is posted in batches of 100 records via the Log Ingestion API.
- The script adds a `CollectionTime` and `Computer` column to every record before posting.

### Key Design Choices

| Decision | Reason |
|---|---|
| Managed Identity auth | No stored credentials, works natively in Azure Automation |
| Key Vault for app secrets | Keeps the log ingestion app credentials out of the script and source control |
| Graph reports endpoint | Covers all sites including those with no recent activity — more complete than usage APIs |
| Display name enrichment | The reports API only returns GUIDs; readable names require a separate sites API call |
| Auto schema management | Avoids breaking the pipeline when new fields are added to the output |



## Blog description:

This PowerShell script serves as the data collection layer for the SharePoint Online site metadata portion of the dashboard. It runs as an Azure Automation runbook with a Hybrid Worker and gathers per-site storage, file count, page view, and activity data across the entire SharePoint Online tenant.

The script authenticates using a managed identity and retrieves required secrets from Azure Key Vault. Site usage metrics are collected via the Microsoft Graph `getSharePointSiteUsageDetail` reports endpoint, which returns a per-site CSV covering storage consumed, active file counts, page views, and last activity date. A second Graph call to `getAllSites` paginates through the tenant to retrieve human-readable display names and web URLs, which are then merged with the usage data using the site GUID as the join key.

All collected data is normalised into a flat schema aligned with a Data Collection Rule (DCR) and ingested into Log Analytics via a Data Collection Endpoint (DCE). The table schema for `sharepointsiteinformation_CL` is automatically created or updated if the data structure changes, making the data directly available for KQL queries and dashboard visualisations without manual table management.
