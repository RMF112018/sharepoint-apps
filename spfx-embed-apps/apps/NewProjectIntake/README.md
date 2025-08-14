### New Project Intake (SPFx App)

Modern single-page intake form with canonical write-back and project-site sync. Built on SPFx, React, Fluent UI, and PnPjs in a pnpm monorepo.

### Key Capabilities
- **Data-driven form** from a generated Field Catalog
- **Canonical data model** in Preconstruction; **local mirror** on each project site
- **Two-way sync** with client outbox + backend worker (least privilege)
- **Change tracking**: list versioning + central Audit entries
- **Observability**: non-PII telemetry, sync latency, recent failures

### Data Model (High-level)

```mermaid
graph LR
  subgraph Preconstruction (Central)
    REG["Project Registry"]
    CAN["Project Facts (Canonical)"]
    AUD["Project Audit"]
    Q["Sync Queue"]
  end
  subgraph Project Site (Per-project)
    LOC["Project Facts (Local)"]
  end
  UI["NewProjectIntake Web Part"]
  BW["Sync Engine (Backend: Function/Flow)"]

  UI -- Create/Update --> LOC
  UI -- Outbox jobs (push/pull) --> Q
  BW -- Upsert canonical --> CAN
  BW -- Seed/Write local --> LOC
  BW -- Write audit --> AUD
  REG -. maintain -. CAN
  BW -. idempotent, conflict policy .-
```

### Field Mapping Registry
- Location: `src/config/fieldMappingRegistry.json`
- Purpose: one source of truth routing every form field to its destination column(s)

Schema (excerpt):

```json
{
  "schemaVersion": "1",
  "canonicalSiteRelativeUrl": "/sites/Preconstruction",
  "lists": {
    "canonicalFactsListTitle": "Project Facts (Canonical)",
    "projectSiteFactsListTitle": "Project Facts",
    "syncQueueListTitle": "Project Facts Sync Queue",
    "auditListTitle": "Project Audit"
  },
  "mappings": [
    {
      "field_key": "proposalDueDate",
      "column_internal_name": "ProposalDueDate",
      "type": "DateTime",
      "is_canonical": true,
      "is_editable_on_project_site": true,
      "validation_rules": { "required": true }
    }
  ]
}
```

### Sync Engine (Client + Backend)
- Client outbox (localStorage): jobs with `{ SyncId, ProjectId, ChangedFields, SourceSiteUrl, Action: 'push'|'pull' }`
- Backend worker (Function/Flow):
  - Validates job and role
  - Applies updates to `Project Facts (Canonical)` or seeds `Project Facts (Local)`
  - Idempotent via `SyncId` and ETag
  - Conflict policy: **Canonical wins** unless central value is older → create a Resolve task
  - Writes `Project Audit` row (non-PII snapshot) on success
  - Updates `Project Registry.LastSyncStatus/Error`
- Retries: client automatically retries on page loads; backend should use exponential backoff when calling SharePoint/Graph

### Permissions and Consent
- Client (delegated): `People.Read`, `User.ReadBasic.All`
- Backend (app-only): `Sites.Selected` (Microsoft Graph) for granular access to Preconstruction

Consent steps (summary):
1) Register Azure AD application for backend; add Graph Application permission `Sites.Selected`; grant admin consent
2) Assign site access via Graph:
   - Get site id: `GET https://graph.microsoft.com/v1.0/sites/root:/{siteRelativePath}`
   - Grant role (write) to the app:
```http
POST https://graph.microsoft.com/v1.0/sites/{siteId}/permissions
Content-Type: application/json
{
  "roles": ["write"],
  "grantedToIdentities": [{ "application": { "id": "<appId>", "displayName": "NewProjectIntake Sync" } }]
}
```
3) Rollback: remove the permission object from the site permissions collection; or remove `Sites.Selected` app permission

### Operational SOP
- **Resync a site**
  - From the app: use “Resync” (push pending) or “Pull Canonical” (admin-only backend guard)
  - From backend: seed `Project Facts (Local)` from canonical via `ProjectId`
- **Migrate schema**
  - On app load, `ensureFactsSchemaUpToDate` adds missing fields and marks deprecated read-only
  - Record `SchemaVersion` on items; re-seed if needed
- **Pause sync**
  - Client: set `src/config/sync.json` endpoint to empty to pause dispatch
  - Backend: disable worker/Flow; optionally place a banner via config
- **Rollback**
  - Revoke site permission (Sites.Selected), disable backend; the client will continue local saves with queued changes

### Observability
- Telemetry (non-PII): `localSaved`, `syncQueued`, `syncStart`, `syncSuccess`, `syncError`, `conflictDetected`, `resolved` with `CorrelationId` and `ProjectId`
- Client latency: rolling p50/p95, recent failures
- Audit: `Project Audit` list rows per central change

### Developer Notes
- Code lives under `src/webparts/demoEmbed/*` (renderer/components) and `src/data/*` (mappers/sync)
- Shared SP client in `libs/sp-client`

# demo-embed

## Summary

Short summary on functionality and used technologies.

[picture of the solution in action, if possible]

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.21.1-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

> Any special pre-requisites?

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| folder name | Author details (name, company, twitter alias with link) |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.1     | March 10, 2021   | Update comment  |
| 1.0     | January 29, 2021 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

> Include any additional steps as needed.

## Features

Description of the extension that expands upon high-level summary above.

This extension illustrates the following concepts:

- topic 1
- topic 2
- topic 3

> Notice that better pictures and documentation will increase the sample usage and the value you are providing for others. Thanks for your submissions advance.

> Share your web part with others through Microsoft 365 Patterns and Practices program to get visibility and exposure. More details on the community, open-source projects and other activities from http://aka.ms/m365pnp.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
