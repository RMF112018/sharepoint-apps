# Scaffold a new SPFx app (monorepo pattern)

This repository hosts multiple SPFx apps under `apps/*` and shared libraries under `libs/*`.
Follow this pattern to create a new app that consumes shared libs (e.g., `@hbi/sp-client`).

## Prerequisites
- Node 22.14.x (Volta/NVM). Repo pins version via `.nvmrc`, `.node-version`, and `package.json` Volta.
- pnpm v10.x
- Yeoman + `@microsoft/generator-sharepoint` (installed in the workspace devDeps)

## One-command scaffold

```bash
cd spfx-embed-apps
pnpm node tools/scaffold-spfx-app.mjs --name MyNewApp --webpart MyNewApp --description "My new web part"
```

This will:
- Scaffold an SPFx React web part app into `apps/MyNewApp`
- Add `@hbi/sp-client` and `@pnp/sp` as dependencies
- Add PnPjs side-effect imports to `src/index.ts`
- Wire `this.context` from WebPart to the component props
- Set `config/serve.json` to the local workbench

## Manual steps (if needed)
1. Ensure your component props include:
```ts
import type { BaseComponentContext } from '@microsoft/sp-component-base';
export interface IMyProps { context: BaseComponentContext; }
```
2. In the WebPart `render()`, pass `context: this.context` to the component.
3. In your component, create SPFI with:
```ts
import { createSharePointClient } from '@hbi/sp-client';
const sp = createSharePointClient({ spfxContext: this.props.context });
```

## Build & Serve
```bash
pnpm install
pnpm -r build
cd apps/MyNewApp
pnpm exec gulp trust-dev-cert
pnpm exec gulp serve
```

- Local workbench: `https://localhost:4321/temp/workbench.html`
- Hosted workbench: `https://<tenant>/_layouts/15/workbench.aspx?debug=true&noredir=true&debugManifestsFile=https://localhost:4321/temp/manifests.js`

## Notes
- Each app maintains its own manifest and `config/package-solution.json`.
- Apps import shared logic only via `libs/*` packages (no cross-app imports).
- Keep web part files thin; business logic lives in shared libs.
- Respect performance, a11y, and testing gates defined in `/.cursor/rules/spfx.mdc`.
