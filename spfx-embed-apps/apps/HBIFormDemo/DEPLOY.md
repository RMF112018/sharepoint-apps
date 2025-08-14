# Production build & packaging

## Build commands
- Install deps: `pnpm i`
- Build all workspaces: `pnpm -r build`
- Package this app (minified, hashed, production URLs): `pnpm run package`
  - Equivalent to: `gulp bundle --ship && gulp package-solution --ship`

## What --ship does
- Enables production webpack mode (minification, tree-shaking)
- Emits hashed asset filenames for long-term caching
- Writes production content URLs into manifests

## Artifacts and locations
- `apps/HBIFormDemo/temp/deploy/` – hashed production bundles and assets for this app
  - Examples: `demo-embed-web-part_[hash].js`, `demo-embed-web-part_[hash].manifest.json`, CSS/PNG with content hashes
- `apps/HBIFormDemo/sharepoint/solution/demo-embed.sppkg` – solution package ready for upload to the App Catalog

## Content URLs in production
- With `includeClientSideAssets: true` (in `apps/HBIFormDemo/config/package-solution.json`), assets are embedded and served from the Tenant App Catalog ClientSideAssets library.
  - After upload, SharePoint rewrites URLs to a tenant CDN endpoint similar to:
    - `https://publiccdn.sharepointonline.com/<tenant-host>/ClientSideAssets/<guid>/demo-embed-web-part_[hash].js`
- No changes to `apps/HBIFormDemo/config/write-manifests.json` are required for this mode.

## Hosting on your own CDN (optional)
- If you plan to host assets on Azure Storage or another CDN:
  1) Set `includeClientSideAssets` to `false` in `apps/HBIFormDemo/config/package-solution.json`
  2) Set `cdnBasePath` in `apps/HBIFormDemo/config/write-manifests.json` to your CDN container URL (e.g., `https://<storage-account>.z22.web.core.windows.net/spfx/demo-embed`)
  3) Run `pnpm run package` and upload the contents of `apps/HBIFormDemo/temp/deploy/` to that CDN path
  4) Upload the new `apps/HBIFormDemo/sharepoint/solution/demo-embed.sppkg` to the App Catalog

## Upload & deploy
1) Upload `apps/HBIFormDemo/sharepoint/solution/demo-embed.sppkg` to the tenant App Catalog
2) Approve deployment (tenant-wide deployment supported; not required)
3) Add the web part "HBI Form Demo" to a page

### Tenant App Catalog upload (role-based steps for KME Holding)
- Role: SharePoint admin or App Catalog admin
1. Open Microsoft 365 admin center → Admin centers → SharePoint (opens SharePoint admin center)
2. In SharePoint admin center: More features → Apps → Open
3. Click “App Catalog” → open the App Catalog site
4. In the App Catalog site, open the “Apps for SharePoint” library
5. Upload `hbi-form-demo-1.0.0.sppkg`
6. When prompted, decide on tenant-wide availability:
   - Per-site installation (recommended for pilots):
     - Pros: limits blast radius; validate list/permissions per site
     - Cons: install on each site that needs the app
   - Tenant-wide availability (optional, supported via `skipFeatureDeployment: true`):
     - Pros: editors can add the web part on any site without per-site install
     - Cons: broader exposure; coordinate comms/support before enabling
7. Confirm/approve the upload

### API access (if prompted)
- This solution requests no additional API permissions. No action is required.
- If a prompt appears unexpectedly:
  1. Go to SharePoint admin center → API access
  2. Review Pending requests
  3. Approve only least-privilege scopes required for the business need, or Reject and contact the app owner

### Confirm deployment and visibility
- In App Catalog → Apps for SharePoint: verify the package shows Deployed = Yes and version 1.0.0.0
- On a test site: Site contents → Add an app → confirm “HBI Form Demo” appears under “From your organization”

## Verification checklist
- Web part loads with minified resources (no source maps in prod)
- Network requests serve hashed JS/CSS from App Catalog CDN
- No unexpected API permission prompts
- Form submits successfully to target list

## Tenant targeting: kmeholding.sharepoint.com
- Recommended distribution: Tenant App Catalog on `kmeholding.sharepoint.com`.
  - Pros: automatic ClientSideAssets hosting/CDN rewrite; single place to approve updates.

- Pilot fallback: Site Collection App Catalog (per-site)
  1) Ensure the target site has a Site Collection App Catalog enabled (tenant admin: Apps > Manage apps > enable site app catalog for the site)
  2) Upload `demo-embed.sppkg` to the site’s App Catalog
  3) Install the app on that site from “From your organization”

- Tenant-wide vs per-site
  - Default: install per-site for pilots to limit blast radius and validate permissions/lists
  - Once validated, you may enable tenant-wide availability (supported by this solution via `skipFeatureDeployment: true`), allowing editors to add the web part on any site. This solution does not require tenant-wide installation.

- Web part picker friendly name
  - Search for: "HBI Form Demo"
