import type { BaseComponentContext } from '@microsoft/sp-component-base';
import { createSharePointClient } from '@hbi/sp-client';
import fieldCatalog from '../config/fieldCatalog.json';
import registry from '../config/fieldMappingRegistry.json';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields/list';
import '@pnp/sp/lists/web';

type Catalog = Array<{
  field_key: string;
  label: string;
  group: string;
  suggested_type: 'Text' | 'Number' | 'DateTime' | 'Choice' | 'Boolean' | 'Multiline' | 'Person';
  required?: boolean;
  choice_options?: string[];
  multi_select?: boolean;
}>;

export interface FieldSchema {
  internalName: string;
  displayName: string;
  type: 'Text' | 'Number' | 'DateTime' | 'Choice' | 'Boolean' | 'Note' | 'User' | 'UserMulti';
  required: boolean;
  choices?: string[];
}

export interface ListSchema {
  listTitle: string;
  titleDisplayName: string; // Title field display name
  fields: FieldSchema[];
  indexed: string[]; // internal names to index
}

function toInternalName(key: string): string {
  return key.replace(/[^A-Za-z0-9_]/g, '');
}

function buildSchemaFromCatalog(catalog: Catalog, listTitle = 'NewProjectIntake'): ListSchema {
  const fields: FieldSchema[] = [];
  let titleDisplay = 'Job Name';
  let jobNumberInternal = '';
  let proposalDueInternal = '';
  // Ensure SchemaVersion column exists on each Facts list
  fields.push({
    internalName: 'SchemaVersion',
    displayName: 'Schema Version',
    type: 'Text',
    required: false
  });

  for (const f of catalog) {
    const internalName = toInternalName(f.field_key);
    let type: FieldSchema['type'];
    switch (f.suggested_type) {
      case 'Text': type = 'Text'; break;
      case 'Multiline': type = 'Note'; break;
      case 'Number': type = 'Number'; break;
      case 'DateTime': type = 'DateTime'; break;
      case 'Boolean': type = 'Boolean'; break;
      case 'Choice': type = 'Choice'; break;
      case 'Person': type = f.multi_select ? 'UserMulti' : 'User'; break;
      default: type = 'Text';
    }

    // Track special fields
    if (/job\s*name/i.test(f.label)) {
      titleDisplay = 'Job Name';
    }
    if (/job\s*number/i.test(f.label)) {
      jobNumberInternal = internalName;
    }
    if (/proposal\s*due/i.test(f.label)) {
      proposalDueInternal = internalName;
    }

    // Do not add a separate field for Title mapping here; keep catalog-driven fields only
    fields.push({
      internalName,
      displayName: f.label,
      type,
      required: !!f.required,
      choices: f.suggested_type === 'Choice' ? (f.choice_options || []) : undefined
    });
  }

  const indexed: string[] = [];
  if (jobNumberInternal) indexed.push(jobNumberInternal);
  if (proposalDueInternal) indexed.push(proposalDueInternal);

  return {
    listTitle,
    titleDisplayName: titleDisplay,
    fields,
    indexed
  };
}

function fieldXml(f: FieldSchema): string {
  const base = `DisplayName="${escapeXml(f.displayName)}" StaticName="${f.internalName}" Name="${f.internalName}" ${f.required ? 'Required="TRUE"' : ''}`;
  switch (f.type) {
    case 'Text':
      return `<Field Type="Text" ${base} />`;
    case 'Note':
      return `<Field Type="Note" ${base} NumLines="6" RichText="FALSE" />`;
    case 'Number':
      return `<Field Type="Number" ${base} />`;
    case 'DateTime':
      return `<Field Type="DateTime" ${base} />`;
    case 'Boolean':
      return `<Field Type="Boolean" ${base} />`;
    case 'Choice':
      return `<Field Type="Choice" ${base}">
        <CHOICES>${(f.choices || []).map(c => `<CHOICE>${escapeXml(c)}</CHOICE>`).join('')}</CHOICES>
      </Field>`;
    case 'User':
      return `<Field Type="User" ${base} UserSelectionMode="PeopleOnly" />`;
    case 'UserMulti':
      return `<Field Type="UserMulti" ${base} UserSelectionMode="PeopleOnly" Mult="TRUE" />`;
    default:
      return `<Field Type="Text" ${base} />`;
  }
}

function escapeXml(v: string): string {
  return v.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

export async function provisionNewProjectIntakeList(context: BaseComponentContext, listTitle = 'NewProjectIntake'): Promise<ListSchema> {
  const sp = createSharePointClient({ spfxContext: context });
  const schema = buildSchemaFromCatalog(fieldCatalog as Catalog, listTitle);
  // Ensure list
  let list: any;
  try {
    list = sp.web.lists.getByTitle(schema.listTitle);
    await list();
  } catch {
    await sp.web.lists.add(schema.listTitle, 'New Project Intake', 100, true);
    list = sp.web.lists.getByTitle(schema.listTitle);
  }
  // Governance: only allow provisioning on allowed central site
  const preconRel = (registry as any).canonicalSiteRelativeUrl || '/sites/Preconstruction';
  const currentRel = context.pageContext.site.serverRelativeUrl || '';
  if (String(currentRel).toLowerCase() !== String(preconRel).toLowerCase()) {
    throw new Error('Provisioning is only allowed on the Preconstruction site.');
  }
  // Enable versioning for change tracking
  try { await list.update({ EnableVersioning: true, EnableMinorVersions: false }); } catch {}
  // Rename Title display name
  try {
    await list.fields.getByInternalNameOrTitle('Title').update({ Title: schema.titleDisplayName });
  } catch { /* ignore */ }
  // Apply fields (add missing only)
  for (const f of schema.fields) {
    try {
      await list.fields.getByInternalNameOrTitle(f.internalName)();
      // Exists: optionally update choice set if Choice
      if (f.type === 'Choice' && f.choices && f.choices.length) {
        try {
          await list.fields.getByInternalNameOrTitle(f.internalName).update({ Choices: { results: f.choices } });
        } catch { /* ignore */ }
      }
    } catch {
      // Add new field via XML so we control internal name
      const xml = fieldXml(f);
      await list.fields.createFieldAsXml(xml);
    }
  }
  // Index fields
  for (const iname of schema.indexed) {
    try {
      await list.fields.getByInternalNameOrTitle(iname).setIndexed(true);
    } catch { /* ignore */ }
  }
  return schema;
}


