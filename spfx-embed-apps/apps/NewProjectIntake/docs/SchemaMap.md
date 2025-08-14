### Data Model Diagrams

Central vs Project Site

```mermaid
flowchart LR
  subgraph Central
    REG[Project Registry]
    CAN[Project Facts (Canonical)]
    AUD[Project Audit]
    Q[Sync Queue]
  end
  subgraph Project Site
    LOC[Project Facts (Local)]
  end
  UI[NewProjectIntake]
  BW[Sync Engine Backend]

  UI --> LOC
  UI --> Q
  BW --> CAN
  BW --> LOC
  BW --> AUD
```

### Field Mapping Registry
- Location: `src/config/fieldMappingRegistry.json`
- Defines: lists, schemaVersion, and `mappings[]` with `field_key`, `column_internal_name`, destinations, and validation

### Sync Engine Behavior
- Client queues `'push'|'pull'` jobs; backend applies with idempotency (`SyncId`) and ETags
- Conflict policy: canonical wins unless central value older (then create Resolve task)
- Retries: client retries on load; backend should backoff and log to Registry

# NewProjectIntake List Schema Map (version 1)

This schema is generated from `src/config/fieldCatalog.json`.

- Title display name: Job Name
- Indexed: Job Number, Proposal Due Date (if present)

| Internal Name | Display Name | Type | Required | Choices |
|---|---|---|---:|---|
| Title | Job Name | Text |  |  |

> Note: Additional fields are listed in JSON at `src/config/fieldCatalog.json`; internal names derive from `field_key` normalized for SharePoint`.

## Field Mapping Registry (SchemaVersion 1)

Config: `src/config/fieldMappingRegistry.json`

Columns:
- field_key, canonical_list, site_list, column_internal_name, type, is_canonical, is_editable_on_project_site, validation_rules

All displayed fields must resolve to a destination. The app reads this registry to route writes (canonical vs. project site) and to build forms dynamically.
