## 1.0.0 (2025-08-14)

### Purpose
- Initial customer-ready release of the HBI Form Demo SPFx web part. Demonstrates a simple form pattern that captures a Title and Notes and submits to a SharePoint list.

### Notable UX
- Clean, accessible form with required field validation and ARIA live-region status messaging.
- Theme-aware styling and Teams/SharePoint environment messaging.

### Data interactions
- Uses delegated context (no additional API permissions) via `@hbi/sp-client` (PnPjs) to:
  - Retrieve current user on mount.
  - Ensure/validate existence of target list `DemoTestList`.
  - Create list items with fields: `Title`, `Notes`.

### Performance budget
- Target bundle size: ≤ 250KB gzip per app.
- Target TTI: ≤ 2s on modern pages.
- Notes: No heavy charts/editors; single web part bundle. Batching used in library where applicable.

### Deploy notes
- Version alignment:
  - npm package: `1.0.0` (web part manifest uses `"version": "*"`).
  - solution package: `1.0.0.0`.
- Build and package:
  1) `pnpm i`
  2) `pnpm -r build`
  3) In `apps/HBIFormDemo`: `pnpm run package`
- Upload `sharepoint/solution/demo-embed.sppkg` to the tenant App Catalog.
- Approve deployment (tenant-wide deployment supported; not required).
- Add the "HBI Form Demo" web part to a page. Ensure list `DemoTestList` exists or provision per site admin if required.




