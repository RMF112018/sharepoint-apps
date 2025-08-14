### Admin Runbook: New Project Intake

This runbook covers data model, mapping registry, sync engine, permissions/consent, and operational procedures.

### Data Model Overview
- Central (Preconstruction): `Project Registry`, `Project Facts (Canonical)`, `Project Audit`, `Sync Queue`
- Project Site: `Project Facts (Local)`
- Primary key: `ProjectId` (GUID), generated on intake and used across all lists/sites

### Field Mapping Registry
- Path: `src/config/fieldMappingRegistry.json`
- Contains: `schemaVersion`, `canonicalSiteRelativeUrl`, `lists` titles, and `mappings[]`
- Every displayed field must have a mapping; app fails fast if missing

### Sync Engine
- Client
  - Writes locally to intake list; deltas are computed and enqueued to an outbox
  - Outbox jobs: `{ SyncId, ProjectId, ChangedFields, SourceSiteUrl, Action: 'push'|'pull' }`
  - Automatic retry on subsequent page loads
- Backend (Azure Function or Power Automate)
  - Validates identity and role; enforces origin site guardrails
  - Applies changes to Canonical or seeds Local based on `Action`
  - Idempotent: `SyncId` and SharePoint ETags
  - Conflict policy: canonical wins unless central value older → create Resolve task
  - Success: write `Project Audit` entry; set `Project Registry.LastSyncStatus='InSync'`
  - Failure: set `LastSyncStatus='Error'` and `LastSyncError` with reason; return error to client

### Permissions & Consent (Least Privilege)
- Client delegated scopes: `People.Read`, `User.ReadBasic.All` (Graph)
- Backend app-only: `Sites.Selected` (Graph) for Preconstruction site access only
- Steps:
  1. Register AAD app for backend; add Graph Application permission `Sites.Selected`; grant admin consent
  2. Grant site access on Preconstruction via Graph `sites/{siteId}/permissions` with role `write`
  3. Store backend endpoint and API key in `src/config/sync.json` (do not commit secrets)
- Rollback:
  - Remove the site permission assignment; or revoke `Sites.Selected` app permission
  - Disable the backend function/Flow

### Operational SOP
- Resync a site
  - From the app: use “Retry” (push queued) or “Pull Canonical” (backend enforces admin role)
  - From backend: call seeding path to upsert `Project Facts (Local)` by `ProjectId`
- Migrate schema
  - On load, the app calls `ensureFactsSchemaUpToDate` to add missing fields; `SchemaVersion` is stored per item
  - Use the admin button to trigger reseeding (if enabled)
- Pause sync
  - Client: clear the `endpointUrl` in `src/config/sync.json`
  - Backend: pause the function/Flow and optionally show a banner via app config
- Troubleshooting
  - Check `Project Registry.LastSyncStatus/Error`
  - Review `Project Audit` for change history
  - Use Activity tab in the app to view recent client failures and sync latency


