### Required SharePoint Lists and Schemas

This app requires the following lists. Types shown are SharePoint field types. Internal names are SharePoint-safe.

### 1) NewProjectIntake (Intake Submissions)
- Versioning: Enabled
- Columns:
  - Title (Text) — display name: Job Name
  - ProjectId (Text)
  - SchemaVersion (Text)
  - Dynamic fields from `src/config/fieldCatalog.json` (generated)
    - Text → Text
    - Multiline → Note
    - Number → Number
    - DateTime → DateTime
    - Boolean → Boolean
    - Choice → Choice (choices from catalog)
    - Person (single) → User (PeopleOnly)
    - Person (multi) → UserMulti (PeopleOnly)

### 2) Project Registry (Central)
- Versioning: Enabled
- Columns:
  - Title (Text) — optional
  - ProjectId (Text)
  - ProjectNumber (Text)
  - ProjectName (Text)
  - TargetSiteUrl (Text)
  - Status (Text)
  - IntakeItemId (Number)
  - CreatedBy (Text)
  - CreatedUtc (DateTime)
  - CorrelationId (Text)
  - SchemaVersion (Text)
  - LastSyncStatus (Text) — values: InSync | Pending | Error
  - LastSyncError (Note)

### 3) Project Facts (Canonical) — Central
- Versioning: Enabled
- Columns:
  - Title (Text) — optional
  - ProjectId (Text)
  - SchemaVersion (Text)
  - All mapped fields from `Field Mapping Registry` (see below)
    - Text → Text
    - Multiline → Note
    - Number → Number
    - DateTime → DateTime
    - Boolean → Boolean
    - Choice → Choice
    - Person (single) → User
    - Person (multi) → UserMulti
- Indexing (recommended):
  - JobNumber (if present)
  - ProposalDueDate (if present)

### 4) Project Facts (Local) — Per Project Site
- Versioning: Enabled
- Columns mirror Project Facts (Canonical):
  - Title (Text) — optional
  - ProjectId (Text)
  - SchemaVersion (Text)
  - All mapped fields from `Field Mapping Registry`

### 5) Project Audit — Central
- Versioning: Enabled
- Columns:
  - Title (Text) — set to ProjectId
  - ProjectId (Text)
  - FieldKeys (Note)
  - OldJson (Note)
  - NewJson (Note)
  - SourceUser (Text)
  - SyncId (Text)
  - Timestamp (DateTime)

### 6) Project Facts Sync Queue — Central
- Purpose: client outbox jobs picked by backend
- Versioning: Enabled (optional)
- Columns:
  - Title (Text) — set to ProjectId
  - ProjectId (Text)
  - ChangedJson (Note)
  - CorrelationId (Text)

### Field Mapping Registry (reference)
- Path: `src/config/fieldMappingRegistry.json`
- Lists block must include titles used above:
  - `canonicalFactsListTitle`: Project Facts (Canonical)
  - `projectSiteFactsListTitle`: Project Facts
  - `syncQueueListTitle`: Project Facts Sync Queue
  - `auditListTitle`: Project Audit
- `mappings[]` defines every field surfaced by the app and its SharePoint column internal name and type.

### Notes
- All lists should include `SchemaVersion` for compatibility checks.
- `ProjectId` is the cross-system key and must exist on Intake, Registry, both Facts lists, Audit, and Sync Queue.
- Enable list versioning on both Facts lists and the Audit list for change tracking.


