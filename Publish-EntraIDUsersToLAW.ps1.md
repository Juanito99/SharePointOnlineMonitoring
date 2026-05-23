## How the Data Gets Into Log Analytics

Before the KQL queries can do anything, data needs to land in your Log Analytics workspace. That's what `Publish-EntraIDUsersToLAW.ps1` handles.

The script runs on a schedule (via Azure Automation with a Hybrid Worker) and has two phases: collect and ingest.

### Phase 1 — Collect from Microsoft Graph

It authenticates using **Managed Identity** — no credentials in the script. The secret required for the log ingestion application is pulled from **Azure Key Vault** at runtime.

One Graph API call is made:

- **`Get-MgUser`** — retrieves all enabled member accounts synchronised from on-premises Active Directory. The filter `accountEnabled eq true and userType eq 'Member'` is applied server-side; the script then further discards any user whose `OnPremisesDistinguishedName` is empty, ensuring only hybrid-synced identities are included.

The following attributes are collected per user: `id`, `displayName`, `userPrincipalName`, `userType`, `givenName`, `surname`, `mail`, `department`, `city`, `postalCode`, `state`, `country`, `companyName`, `onPremisesSamAccountName`, `onPremisesDistinguishedName`, `onPremisesExtensionAttributes` (attributes 1, 2, 5, 10, 11, 12), and `signInActivity` (last interactive sign-in date).

> The `signInActivity` property requires the `AuditLog.Read.All` Graph permission. If that permission is absent, the script automatically retries the query without it and continues — it degrades gracefully rather than failing.

All collected users are mapped into a strongly typed `EntraUserRecord` class. Comma characters are stripped from every string field to prevent CSV-level parsing issues downstream.

### Phase 2 — Ingest into Log Analytics

Once the user data is shaped, it is pushed to Log Analytics using the **Data Collection Rule (DCR) / Data Collection Endpoint (DCE)** ingestion pipeline via the `AzLogDcrIngestPS` module.

What that means in practice:
- The table schema (`entraidusers_CL`) is **automatically created or updated** if the data structure changes — no manual table management.
- A `CollectionTime` and `Computer` column are added to every record before posting.
- Records are split into payload batches capped at 500 KB each to respect the Log Ingestion API limits. Any individual batch that still exceeds 1 MB is skipped with a warning rather than crashing the run.
- Each batch is posted with **exponential back-off retry logic** (up to 4 attempts) covering transient HTTP errors (408, 429, 500, 502, 503, 504).

### Key Design Choices

| Decision | Reason |
|---|---|
| Managed Identity auth | No stored credentials, works natively in Azure Automation |
| Key Vault for app secret | Keeps the log ingestion app credentials out of the script and source control |
| Server-side Graph filter | Reduces data transfer by filtering enabled members directly in the API call |
| On-premises DN check | Limits the dataset to hybrid-synced users aligned with the existing AAD export runbook |
| Graceful signInActivity fallback | Avoids a hard failure when the Automation account lacks the AuditLog.Read.All permission |
| Strongly typed `EntraUserRecord` class | Enforces a stable, predictable schema regardless of what Graph returns for optional fields |
| Byte-size batch splitting | Prevents oversized payloads from being rejected by the Log Ingestion API |
| Exponential back-off retries | Handles transient API throttling and intermittent network errors without manual intervention |
| Auto schema management | Avoids breaking the ingestion pipeline when new fields are added to the output |



## Blog Description

This PowerShell script serves as the data collection layer for the Entra ID user metadata portion of the dashboard. It runs as an Azure Automation runbook with a Hybrid Worker and gathers user identity and directory attributes for all on-premises-synchronised member accounts in the Entra ID tenant.

The script authenticates using a managed identity and retrieves the required log ingestion secret from Azure Key Vault. User data is collected via the Microsoft Graph `Get-MgUser` cmdlet with a server-side filter for enabled member accounts. An additional local filter discards any user without an `OnPremisesDistinguishedName`, keeping the dataset consistent with the existing on-premises AD export pipeline. Where the `AuditLog.Read.All` permission is available, the last interactive sign-in date is included; if it is not, the script retries the query without that field and continues without interruption.

All collected users are normalised into a flat schema using a strongly typed `EntraUserRecord` class and ingested into Log Analytics via a Data Collection Endpoint (DCE) and Data Collection Rule (DCR). Records are split into payload batches capped at 500 KB, and each batch is posted with exponential back-off retry logic to handle transient API errors. The table schema for `entraidusers_CL` is automatically created or updated if the data structure changes, making the data directly available for KQL queries and dashboard visualisations without manual table management.
