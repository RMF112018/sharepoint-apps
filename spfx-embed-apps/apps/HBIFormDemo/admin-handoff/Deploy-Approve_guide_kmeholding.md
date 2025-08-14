# Deploy & Approve Guide — kmeholding.sharepoint.com

Version: 1.0.0
App: HBI Form Demo (SPFx)
Package: hbi-form-demo-1.0.0.sppkg

## Overview
This guide covers Tenant App Catalog upload, API access approval (if any), tenant-wide toggle guidance, per-site install steps, and rollback.

## Prerequisites
- Tenant admin or App Catalog admin on `kmeholding.sharepoint.com`
- Access to the Tenant App Catalog site
- Package file: `hbi-form-demo-1.0.0.sppkg`

## 1) Upload to Tenant App Catalog (recommended)
Role: SharePoint admin or App Catalog admin

1. Open Microsoft 365 admin center → Admin centers → SharePoint (opens SharePoint admin center)
2. In SharePoint admin center: More features → Apps → Open
3. Click “App Catalog” → open the App Catalog site
4. In the App Catalog site, open the “Apps for SharePoint” library
5. Upload `hbi-form-demo-1.0.0.sppkg`
6. When prompted, decide on tenant-wide availability (see section 3 below), then confirm/approve the upload

[Screenshot placeholder: App Catalog upload]

## 2) API access approval
- This solution requests no additional API permissions. No action is required.
- If you see any permission prompt unexpectedly:
  1. Go to SharePoint admin center → API access
  2. Review “Pending requests”
  3. Approve only least-privilege scopes required for business need; otherwise Reject and contact the app owner

[Screenshot placeholder: API access panel (none)]

## 3) Deployment scope (tenant-wide vs per-site)
During the upload prompt, you may see “Make this solution available to all sites in the organization”.

- Per-site installation (recommended for pilots)
  - Pros: small blast radius, validate list/permissions per site
  - Cons: requires installing the app on each site that needs it

- Tenant-wide availability (supported, optional)
  - Pros: editors can add the web part on any site without per-site app install
  - Cons: broader exposure; coordinate communications and support before enabling
  - This solution supports tenant-wide via `skipFeatureDeployment: true`, but does not require it

[Screenshot placeholder: deployment dialog]

## 4) Install on a site (per-site pilot)
1. Go to the target site
2. Open Site Contents > “New” > “App” (or “From your organization”)
3. Find “HBI Form Demo” and choose “Add”

[Screenshot placeholder: Add app to site]

## 5) Rollback
- Per-site removal:
  - Site Contents > find the app > “...” > Remove
- Tenant rollback:
  - In App Catalog, remove or replace the `.sppkg` with a prior version
  - If tenant-wide was enabled, disable it and/or retract the app

[Screenshot placeholder: Remove app]

## 6) Confirm successful deployment and visibility
- In App Catalog → “Apps for SharePoint”, verify the package shows with Deployed = Yes and the expected version (1.0.0.0)
- On a test site, browse to Site contents → Add an app → confirm “HBI Form Demo” is visible in “From your organization”

## Reference
- App display name (picker): “HBI Form Demo”
- Solution name: `hbi-form-demo-client-side-solution`
- Solution ID: `804c482d-6f49-4e5f-bec6-eeeef567ca5a`
