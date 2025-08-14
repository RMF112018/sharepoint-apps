import type { BaseComponentContext } from '@microsoft/sp-component-base';
import { createSharePointClient } from '@hbi/sp-client';
import registry from '../config/fieldMappingRegistry.json';

export interface MigrationResult {
  migrated: string[];
  deprecated: string[];
}

function normalizeName(name: string): string { return name.replace(/[^A-Za-z0-9_]/g, ''); }

export async function ensureFactsSchemaUpToDate(context: BaseComponentContext, listTitle: string): Promise<MigrationResult> {
  const sp = createSharePointClient({ spfxContext: context });
  const list = sp.web.lists.getByTitle(listTitle);
  const fields = await list.fields.select('InternalName', 'Title', 'TypeAsString')();
  const have = new Set<string>(fields.map((f: any) => f.InternalName));
  const migrated: string[] = [];
  const deprecated: string[] = [];
  const reg: any = registry;
  const schemaVersion = reg.schemaVersion || '1';

  // Ensure SchemaVersion field exists
  if (!have.has('SchemaVersion')) {
    await list.fields.createFieldAsXml(`<Field Type="Text" DisplayName="Schema Version" StaticName="SchemaVersion" Name="SchemaVersion" />`);
    migrated.push('SchemaVersion');
  }

  // Add missing fields from mappings (non-destructive add)
  for (const m of reg.mappings || []) {
    const iname = normalizeName(m.column_internal_name || m.field_key);
    if (!have.has(iname)) {
      const type = m.type || 'Text';
      const typeXml = type === 'Multiline' ? 'Note' : (type === 'Person' ? 'User' : type);
      await list.fields.createFieldAsXml(`<Field Type="${typeXml}" DisplayName="${m.field_key}" StaticName="${iname}" Name="${iname}" />`);
      migrated.push(iname);
    }
  }

  // Mark deprecated fields read-only if present
  for (const f of fields) {
    const isMapped = (reg.mappings || []).some((m: any) => normalizeName(m.column_internal_name || m.field_key) === f.InternalName);
    if (!isMapped && f.InternalName !== 'SchemaVersion') {
      try {
        await list.fields.getByInternalNameOrTitle(f.InternalName).update({ ReadOnlyField: true });
        deprecated.push(f.InternalName);
      } catch {}
    }
  }

  // Record migration note by setting SchemaVersion on list item add/update paths
  try {
    await list.items.filter("SchemaVersion ne '" + schemaVersion + "'").top(1)();
    // No bulk update here; we set SchemaVersion when saving items
  } catch {}

  return { migrated, deprecated };
}

export async function reseedProjectSiteFromCanonical(_context: BaseComponentContext, _projectId: string): Promise<void> {
  // Intentionally left as a placeholder; call seedProjectSiteFacts(...) with canonical fields
}


