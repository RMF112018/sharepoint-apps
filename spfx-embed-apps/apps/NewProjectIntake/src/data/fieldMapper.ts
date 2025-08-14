import type { BaseComponentContext } from '@microsoft/sp-component-base';
import fieldCatalog from '../config/fieldCatalog.json';
import { ensureSharePointUserIds, type PersonSuggestion } from '../utils/people';
import registry from '../config/fieldMappingRegistry.json';

export type CatalogField = {
  field_key: string;
  label: string;
  group: string;
  suggested_type: 'Text' | 'Number' | 'DateTime' | 'Choice' | 'Boolean' | 'Multiline' | 'Person';
  required?: boolean;
  choice_options?: string[];
  multi_select?: boolean;
};

export type FormValue = string | number | boolean | PersonSuggestion[] | undefined;
export type FormState = Record<string, FormValue>;

export interface SelectExpand {
  select: string[];
  expand: string[];
}

function toInternalName(key: string): string { return key.replace(/[^A-Za-z0-9_]/g, ''); }

function isJobNameLabel(label: string): boolean { return /job\s*name/i.test(label); }
// keep helpers available for future mapping expansions
// reserved helpers for future use
// function isJobNumberLabel(_label: string): boolean { return /job\s*number/i.test(_label); }
// function isProposalDueLabel(_label: string): boolean { return /proposal\s*due/i.test(_label); }

export function getCatalog(): CatalogField[] {
  return (fieldCatalog as unknown as CatalogField[]);
}

export interface FieldDestination {
  field_key: string;
  canonical_list: string;
  site_list: string;
  column_internal_name: string;
  type: CatalogField['suggested_type'];
  is_canonical: boolean;
  is_editable_on_project_site: boolean;
  validation_rules?: Record<string, unknown>;
}

export interface FieldMappingRegistry {
  schemaVersion: string;
  canonicalSiteRelativeUrl: string;
  lists: { canonicalFactsListTitle: string; projectSiteFactsListTitle: string };
  mappings: FieldDestination[];
}

export function getFieldMappingRegistry(): FieldMappingRegistry {
  return registry as unknown as FieldMappingRegistry;
}

export function resolveDestination(fieldKey: string): FieldDestination | undefined {
  const reg = getFieldMappingRegistry();
  for (const m of reg.mappings) {
    if (m.field_key === fieldKey) return m;
  }
  return undefined;
}

export function buildSelectExpand(catalog: CatalogField[]): SelectExpand {
  const select = ['Id', 'Title'];
  const expand: string[] = [];
  for (const f of catalog) {
    const iname = toInternalName(f.field_key);
    switch (f.suggested_type) {
      case 'Person':
        select.push(`${iname}/Id`, `${iname}/Title`, `${iname}/EMail`, `${iname}/UserPrincipalName`);
        expand.push(iname);
        break;
      default:
        select.push(iname);
    }
  }
  return { select, expand };
}

export async function formToSpPayload(context: BaseComponentContext, catalog: CatalogField[], form: FormState): Promise<Record<string, any>> {
  const payload: Record<string, any> = {};
  // Title mapping from Job Name
  let jobNameField: CatalogField | undefined;
  for (const f of catalog) { if (isJobNameLabel(f.label)) { jobNameField = f; break; } }
  if (jobNameField) { payload['Title'] = String(form[jobNameField.field_key] ?? ''); }

  for (const f of catalog) {
    const iname = toInternalName(f.field_key);
    const v = form[f.field_key];
    if (v === undefined) continue;
    switch (f.suggested_type) {
      case 'Text':
      case 'Multiline':
      case 'Choice':
        payload[iname] = v == null ? '' : String(v as string);
        break;
      case 'Number': {
        const n = typeof v === 'number' ? v : parseFloat(String(v));
        if (!isNaN(n as any)) payload[iname] = n; // else skip
        break;
      }
      case 'Boolean':
        payload[iname] = !!v;
        break;
      case 'DateTime': {
        const s = String(v ?? '');
        if (s) payload[iname] = s; // assume ISO/local datetime string compatible with SP
        break;
      }
      case 'Person': {
        const arr = Array.isArray(v) ? v as PersonSuggestion[] : [];
        const ids = await ensureSharePointUserIds(context, arr.map(p => p.email || p.upn || p.key).filter(Boolean) as string[]);
        const key = `${iname}Id`;
        if (f.multi_select) payload[key] = { results: ids };
        else payload[key] = ids[0];
        break;
      }
      default:
        // ignore unknown types defensively
        break;
    }
  }
  return payload;
}

export function spItemToForm(catalog: CatalogField[], item: any): FormState {
  const form: FormState = {};
  let jobNameField: CatalogField | undefined;
  for (const f of catalog) { if (isJobNameLabel(f.label)) { jobNameField = f; break; } }
  if (jobNameField) form[jobNameField.field_key] = item['Title'];
  for (const f of catalog) {
    const iname = toInternalName(f.field_key);
    const raw = item[iname];
    switch (f.suggested_type) {
      case 'Person': {
        const p = item[iname];
        if (Array.isArray(p)) {
          form[f.field_key] = p.map((u: any) => ({ key: String(u.Id), name: u.Title, email: u.EMail, upn: u.UserPrincipalName })) as any;
        } else if (p && typeof p === 'object') {
          form[f.field_key] = [{ key: String(p.Id), name: p.Title, email: p.EMail, upn: p.UserPrincipalName }] as any;
        } else {
          form[f.field_key] = [];
        }
        break;
      }
      case 'Boolean':
        form[f.field_key] = !!raw;
        break;
      case 'Number':
        form[f.field_key] = typeof raw === 'number' ? raw : (raw ? parseFloat(String(raw)) : undefined);
        break;
      default:
        form[f.field_key] = raw as any;
    }
  }
  return form;
}


