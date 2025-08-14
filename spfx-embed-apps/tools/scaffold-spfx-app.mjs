#!/usr/bin/env node
// Scaffold a new SPFx React web part app into apps/<AppName>
// - Generates via Yeoman with a pre-filled .yo-rc.json
// - Ensures PnPjs side-effect imports exist for SPFI augmentation
// - Adds dependency on @hbi/sp-client and @pnp/sp
// - Wires web part to pass this.context down via props
// - Points serve.json to local workbench

import { spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join, dirname } from 'node:path';

function parseArgs(argv) {
  const args = {};
  for (let i = 2; i < argv.length; i++) {
    const a = argv[i];
    if (a.startsWith('--')) {
      const [k, v] = a.split('=');
      args[k.slice(2)] = v ?? 'true';
    }
  }
  return args;
}

function run(cmd, cwd, env = process.env) {
  const [bin, ...rest] = cmd.split(' ');
  const res = spawnSync(bin, rest, { stdio: 'inherit', cwd, env });
  if (res.status !== 0) {
    throw new Error(`Command failed: ${cmd}`);
  }
}

function ensureJson(file, obj) {
  writeFileSync(file, JSON.stringify(obj, null, 2) + '\n', 'utf8');
}

function upsertPackageDependency(appPkgPath, depName, depVersion) {
  const pkg = JSON.parse(readFileSync(appPkgPath, 'utf8'));
  pkg.dependencies ||= {};
  if (!pkg.dependencies[depName]) {
    pkg.dependencies[depName] = depVersion;
  }
  writeFileSync(appPkgPath, JSON.stringify(pkg, null, 2) + '\n', 'utf8');
}

function patchServeJson(serveJsonPath) {
  try {
    const json = JSON.parse(readFileSync(serveJsonPath, 'utf8'));
    json.initialPage = 'https://localhost:4321/temp/workbench.html';
    writeFileSync(serveJsonPath, JSON.stringify(json, null, 2) + '\n', 'utf8');
  } catch {}
}

function addPnPjsSideEffects(appSrcIndexPath) {
  let content = '';
  try {
    content = readFileSync(appSrcIndexPath, 'utf8');
  } catch {
    // fallthrough
  }
  const lines = [
    "import '@pnp/sp/webs';",
    "import '@pnp/sp/lists';",
    "import '@pnp/sp/items';",
    "import '@pnp/sp/batching';",
    "import '@pnp/sp/site-users/web';",
  ];
  const header = content.includes("@pnp/sp/webs") ? content : `// PnPjs side-effect imports for SPFI augmentation\n${lines.join('\n')}\n` + (content || '');
  writeFileSync(appSrcIndexPath, header, 'utf8');
}

function wireContextProp(appSrcDir) {
  // Find WebPart and component files
  // WebPart: src/webparts/<name>/<Name>WebPart.ts
  // Props: src/webparts/<name>/components/I<Name>Props.ts
  // Component: src/webparts/<name>/components/<Name>.tsx
  const glob = (pattern) => {
    // very light glob: only one wildcard segment
    const [base, rest] = pattern.split('*');
    const dir = dirname(base);
    const fs = require('node:fs');
    const path = require('node:path');
    const entries = fs.readdirSync(dir, { withFileTypes: true });
    const matches = [];
    for (const e of entries) {
      if (!e.isDirectory()) continue;
      const p = path.join(dir, e.name, rest);
      if (fs.existsSync(p)) matches.push(p);
    }
    return matches;
  };
  const webParts = glob(join(appSrcDir, 'webparts/*/*WebPart.ts'));
  for (const wp of webParts) {
    const compDir = dirname(wp) + '/components';
    const fs = require('node:fs');
    if (!fs.existsSync(compDir)) continue;
    const files = fs.readdirSync(compDir).filter((f) => f.endsWith('.tsx') || f.endsWith('.ts'));
    const propsFile = files.find((f) => f.toLowerCase().startsWith('i') && f.toLowerCase().includes('props'));
    const componentFile = files.find((f) => f.endsWith('.tsx') && !f.toLowerCase().startsWith('i'));

    // Patch WebPart to pass context
    let wpSrc = readFileSync(wp, 'utf8');
    if (!wpSrc.includes('context: this.context')) {
      wpSrc = wpSrc.replace(/(userDisplayName:\s*this\.context\.pageContext\.user\.displayName\s*\n\s*\})/, (m) => m.replace(/\}/, ',\n        context: this.context\n      }'));
      writeFileSync(wp, wpSrc, 'utf8');
    }

    // Patch props interface
    if (propsFile) {
      const pPath = join(compDir, propsFile);
      let pSrc = readFileSync(pPath, 'utf8');
      if (!pSrc.includes('BaseComponentContext')) {
        pSrc = `import type { BaseComponentContext } from '@microsoft/sp-component-base';\n` + pSrc;
      }
      if (!/context:\s*BaseComponentContext/.test(pSrc)) {
        pSrc = pSrc.replace(/}\s*$/m, '  context: BaseComponentContext;\n}\n');
      }
      writeFileSync(pPath, pSrc, 'utf8');
    }

    // Ensure component compiles with context prop (no-op if not used)
    if (componentFile) {
      const cPath = join(compDir, componentFile);
      let cSrc = readFileSync(cPath, 'utf8');
      if (!cSrc.includes('context:')) {
        // nothing to enforce here; component may not use context directly
      }
      writeFileSync(cPath, cSrc, 'utf8');
    }
  }
}

async function main() {
  const args = parseArgs(process.argv);
  const appName = args.name || args.app || 'MySpfxApp';
  const webpartName = args.webpart || appName;
  const description = args.description || `${appName} web part`;

  const repoRoot = process.cwd();
  const appDir = join(repoRoot, 'apps', appName);
  if (!existsSync(appDir)) mkdirSync(appDir, { recursive: true });

  const yoRc = {
    '@microsoft/generator-sharepoint': {
      isCreatingSolution: true,
      environment: 'spo',
      version: '1.21.1',
      libraryName: appName,
      libraryId: '00000000-0000-0000-0000-000000000000',
      packageManager: 'pnpm',
      isDomainIsolated: false,
      componentType: 'webpart',
      framework: 'react',
      webparts: [
        { name: webpartName, description, framework: 'react' }
      ],
      solutionName: appName,
      skipFeatureDeployment: true
    }
  };

  ensureJson(join(appDir, '.yo-rc.json'), yoRc);

  // Run generator (pnpm exec yo ...)
  run('pnpm exec yo @microsoft/sharepoint --skip-install --package-manager pnpm', appDir);

  // Ensure deps and side-effects
  upsertPackageDependency(join(appDir, 'package.json'), '@hbi/sp-client', 'workspace:*');
  upsertPackageDependency(join(appDir, 'package.json'), '@pnp/sp', '^3.26.0');
  // add useful scripts
  try {
    const appPkgPath = join(appDir, 'package.json');
    const pkg = JSON.parse(readFileSync(appPkgPath, 'utf8'));
    pkg.scripts ||= {};
    pkg.scripts.serve = pkg.scripts.serve || 'gulp serve';
    pkg.scripts['trust-cert'] = pkg.scripts['trust-cert'] || 'gulp trust-dev-cert';
    pkg.scripts.package = pkg.scripts.package || 'gulp bundle --ship && gulp package-solution --ship';
    writeFileSync(appPkgPath, JSON.stringify(pkg, null, 2) + '\n', 'utf8');
  } catch {}
  addPnPjsSideEffects(join(appDir, 'src', 'index.ts'));
  patchServeJson(join(appDir, 'config', 'serve.json'));
  wireContextProp(join(appDir, 'src'));

  // Print next steps
  // eslint-disable-next-line no-console
  console.log(`\nScaffolded SPFx app at apps/${appName}.\n\nNext:\n  pnpm install\n  pnpm -r build\n  cd apps/${appName} && pnpm exec gulp trust-dev-cert && pnpm exec gulp serve\n`);
}

main().catch((err) => {
  // eslint-disable-next-line no-console
  console.error(err);
  process.exit(1);
});


