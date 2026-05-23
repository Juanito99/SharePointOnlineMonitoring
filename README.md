## Teaser

Microsoft 365 tells you SharePoint is being used. It doesn't tell you *by whom, for what, or whether half your sites are just wasting storage*.

This post walks through a complete SharePoint Online monitoring solution built on Log Analytics: three PowerShell scripts that collect data from Microsoft Graph, the Purview audit API, and Entra ID, and 20 KQL queries that turn that data into actionable insights — storage trends, dormancy analysis, external access, and usage broken down by city, department, and business function.

No custom connectors. No third-party tools. Just data you already have, made visible.

---

## Introduction

This dashboard provides a structured analysis of SharePoint Online using KQL-based metrics derived from site metadata, user activity events, and Entra ID user attributes collected into Log Analytics via three PowerShell scripts.
It focuses on identifying storage growth patterns, site engagement levels, and governance gaps, while also highlighting external access exposure and organisational usage distribution.

The queries analyse daily and 14-day storage trends, growth velocity per site, and current versus historical consumption, enabling comparison between the largest and fastest-growing sites. Activity metrics from the Purview audit stream are used to identify highly engaged sites and to distinguish active collaboration from dormant or abandoned content.

In addition, the dashboard summarises the governance posture across the SharePoint estate, including site ownership accountability, active versus dormant classification, and inactivity thresholds at 90 and 365 days. Usage breakdowns by city, department, and business function — enriched via Entra ID attributes — connect SharePoint behaviour to organisational structure, and external access analysis covers both individual users and domains.

---

## How the Data Gets There

Before the queries can run, data needs to land in Log Analytics. Three PowerShell scripts handle that — and it's worth understanding what each one feeds, because the KQL queries target their output tables directly.

**`Publish-SharePointSiteInfoToLaw.ps1`** runs on a schedule and calls the Microsoft Graph reports API to collect per-site metadata: storage used and allocated, file counts, last activity date, and page views. It enriches the data with human-readable site names (the reports API only returns GUIDs) and writes everything to the `sharepointsiteinformation_CL` table. This is the source for all storage, dormancy, and site health queries.

**`Publish-SharePointUsageInformationToLAW_ps7.ps1`** runs hourly and pulls raw user activity events from the Microsoft Purview Management Activity API — file accesses, downloads, uploads, and page views. It filters out OneDrive and personal site traffic, deduplicates overlapping poll windows, and writes event-level records to the `sharepointusageinfo_CL` table. This is the source for all activity, external access, and usage pattern queries.

**`Publish-EntraIDUsersToLAW.ps1`** runs on a schedule and retrieves all on-premises-synchronised member accounts from Entra ID via the Microsoft Graph `Get-MgUser` API. It collects identity and directory attributes — display name, UPN, department, city, company, on-premises extension attributes, and last interactive sign-in date — and writes them to the `entraidusers_CL` table. This is the enrichment source that connects SharePoint activity back to organisational structure in the usage breakdown queries.



---

## What's in This Dashboard?

With those two tables in place, 20 KQL queries cover the following areas:

**Storage & Growth**
Track total storage today versus 30 days ago, spot the five fastest-growing sites by daily trend, and identify which sites are consuming the most quota — so you know where cleanup efforts will have the biggest impact.

**Site Activity & Dormancy**
See how many of your sites are genuinely active versus dormant. Queries surface sites with no activity in 90 or 365 days, ranked by storage size — giving you a prioritised cleanup list rather than a flat dump.

**Ownership Gaps**
Find sites that hold files but have no traceable user activity. These are your governance blind spots, and they're surfaced here before they become a compliance problem.

**Usage Patterns**
Understand which sites get the most unique users, file downloads, and active files. Break usage down by city, business function, and department — using Entra ID attributes to connect SharePoint activity back to your org structure.

**External Access**
See the top external downloaders, most-accessed sites by external users, and the top external domains — all filtered to exclude your own tenant domains.

All queries are ready to drop into Log Analytics or SquaredUp. Let's walk through them.
