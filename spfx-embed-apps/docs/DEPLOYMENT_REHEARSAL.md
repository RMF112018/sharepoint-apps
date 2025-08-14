# Deployment Rehearsal Checklist (test tenant/site)

- Package: `apps/HBIFormDemo/sharepoint/solution/hbi-form-demo-1.0.0.sppkg`
- Solution name/ID: `hbi-form-demo-client-side-solution` / `804c482d-6f49-4e5f-bec6-eeeef567ca5a`
- Web part picker name: HBI Form Demo
- API permissions: None requested (no `webApiPermissionRequests`)

## Steps

1) Upload to Tenant App Catalog (test tenant)
- Go to `https://<test-tenant>.sharepoint.com/sites/AppCatalog`
- Open “Apps for SharePoint” and upload `hbi-form-demo-1.0.0.sppkg`
- Approve the upload

2) API access approval
- Confirm no permission prompts appear (none are requested)

3) Deployment scope
- Keep per-site install for rehearsal (tenant-wide supported but not required)

4) Install on a test site
- Test site → Site contents → New → App (or From your organization)
- Add “HBI Form Demo”

5) Add to a modern page and publish
- Edit a modern page → “+” → search “HBI Form Demo” → add → Publish

## Validate

- Load/perf: Web part renders quickly (target ≤2s), no console errors
- Theming: Correct in light/dark; environment message shows
- Form basics: Required validation blocks empty submit; description text shows
- User info: Current user label renders if available
- List ops:
  - Ensure site has a `DemoTestList` (create if needed per site admin)
  - Submit with Title/Notes → success message → item appears in `DemoTestList`
- Permissions:
  - With Contribute: submit succeeds
  - Without Contribute: submit fails gracefully with user-safe error

## Verification notes

- Note any branding/theming issues, unexpected prompts, or slow loads
- Capture quick HAR/Lighthouse sample if load seems slow
- Record the URL of the page used and the account(s) tested

## Smoke test results

| Step | Expected | Actual | Pass/Fail | Notes |
|-----|----------|--------|-----------|-------|
| Upload .sppkg | Upload + approval succeeds |  |  |  |
| API access | No prompts |  |  |  |
| Install app to site | App installs on test site |  |  |  |
| Add to page | “HBI Form Demo” appears and adds |  |  |  |
| Load/perf | Renders ≤2s, no console errors |  |  |  |
| Form required validation | Empty submit blocked |  |  |  |
| Submit item (Contribute user) | Success; item in `DemoTestList` |  |  |  |
| Submit item (Read user) | Graceful error; no item created |  |  |  |
