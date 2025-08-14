import type { BaseComponentContext } from '@microsoft/sp-component-base';
import { createSharePointClient } from '@hbi/sp-client';
import { v4 as uuidv4 } from 'uuid';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

export interface ProjectRegistryItem {
  ProjectId: string; // GUID
  ProjectNumber: string;
  ProjectName: string;
  TargetSiteUrl: string;
  Status: string;
  IntakeItemId: number;
  CreatedBy: string;
  CreatedUtc: string;
  CorrelationId: string;
  SchemaVersion: string;
  LastSyncStatus?: 'InSync' | 'Pending' | 'Error';
  LastSyncError?: string;
}

export interface UpsertRegistryParams extends Partial<ProjectRegistryItem> {
  listTitle?: string; // defaults to 'Project Registry'
}

export function newProjectId(): string { return uuidv4(); }

export async function upsertProjectRegistry(context: BaseComponentContext, item: UpsertRegistryParams): Promise<number> {
  const sp = createSharePointClient({ spfxContext: context });
  const listTitle = item.listTitle || 'Project Registry';
  const registry = sp.web.lists.getByTitle(listTitle);
  const filter = `ProjectId eq '${item.ProjectId}'`;
  const existing = await registry.items.filter(filter).top(1)();
  const payload: any = {
    ProjectId: item.ProjectId,
    ProjectNumber: item.ProjectNumber || '',
    ProjectName: item.ProjectName || '',
    TargetSiteUrl: item.TargetSiteUrl || '',
    Status: item.Status || 'Active',
    IntakeItemId: item.IntakeItemId || 0,
    CreatedBy: item.CreatedBy || context.pageContext.user.displayName,
    CreatedUtc: item.CreatedUtc || new Date().toISOString(),
    CorrelationId: item.CorrelationId || '',
    SchemaVersion: item.SchemaVersion || '1',
    LastSyncStatus: item.LastSyncStatus || undefined,
    LastSyncError: item.LastSyncError || ''
  };
  if (existing.length) {
    await registry.items.getById(existing[0].Id).update(payload);
    return existing[0].Id;
  } else {
    const res = await registry.items.add(payload);
    return res.data.Id;
  }
}

export interface ProjectFactsItem {
  ProjectId: string; // GUID, PK
  // plus flattened/normalized fields from intake
  SchemaVersion?: string;
}

export interface UpsertFactsParams {
  listTitle: string; // 'Project Facts (Canonical)' or local mirror
  siteUrl?: string; // if cross-site write needed later
  ProjectId: string;
  fields: Record<string, any>;
}

export async function upsertProjectFacts(context: BaseComponentContext, params: UpsertFactsParams): Promise<number> {
  const { listTitle, ProjectId, fields } = params;
  const sp = createSharePointClient({ spfxContext: context });
  const list = sp.web.lists.getByTitle(listTitle);
  try { await list.update({ EnableVersioning: true, EnableMinorVersions: false }); } catch {}
  const existing = await list.items.filter(`ProjectId eq '${ProjectId}'`).top(1)();
  const payload = { ProjectId, ...fields } as any;
  if (existing.length) {
    const item = list.items.getById(existing[0].Id);
    const meta = await item.select('Id', 'OData__UIVersionString')();
    await item.update(payload, meta['odata.etag']);
    return meta.Id;
  } else {
    const res = await list.items.add(payload);
    return res.data.Id;
  }
}

export interface AuditEntry {
  ProjectId: string;
  FieldKeys: string[];
  OldJson?: Record<string, unknown>;
  NewJson: Record<string, unknown>;
  SourceUser?: string;
  SyncId?: string;
  Timestamp?: string;
}

export async function writeAuditEntry(context: BaseComponentContext, auditListTitle: string, entry: AuditEntry): Promise<void> {
  const sp = createSharePointClient({ spfxContext: context });
  const list = sp.web.lists.getByTitle(auditListTitle);
  try { await list.update({ EnableVersioning: true, EnableMinorVersions: false }); } catch {}
  await list.items.add({
    Title: entry.ProjectId,
    ProjectId: entry.ProjectId,
    FieldKeys: (entry.FieldKeys || []).join(','),
    OldJson: JSON.stringify(entry.OldJson || {}),
    NewJson: JSON.stringify(entry.NewJson || {}),
    SourceUser: entry.SourceUser || '',
    SyncId: entry.SyncId || '',
    Timestamp: entry.Timestamp || new Date().toISOString()
  } as any);
}

export interface EnqueueSyncParams {
  ProjectId: string;
  ChangedFields: string[];
  SourceSiteUrl: string;
  CorrelationId?: string;
}

export async function enqueueFactsSync(context: BaseComponentContext, params: EnqueueSyncParams, registryListTitle: string): Promise<number> {
  const sp = createSharePointClient({ spfxContext: context });
  const queue = sp.web.lists.getByTitle(registryListTitle);
  const res = await queue.items.add({
    Title: params.ProjectId,
    ProjectId: params.ProjectId,
    ChangedJson: JSON.stringify({ fields: params.ChangedFields, source: params.SourceSiteUrl }),
    CorrelationId: params.CorrelationId || ''
  } as any);
  return res.data.Id;
}

export interface SyncFailureLogParams {
  ProjectId: string;
  Reason: string;
  Details?: string;
  RegistryListTitle?: string; // default 'Project Registry'
}

export async function recordSyncFailure(context: BaseComponentContext, p: SyncFailureLogParams): Promise<void> {
  const sp = createSharePointClient({ spfxContext: context });
  const listTitle = p.RegistryListTitle || 'Project Registry';
  const reg = sp.web.lists.getByTitle(listTitle);
  const existing = await reg.items.filter(`ProjectId eq '${p.ProjectId}'`).top(1)();
  if (existing.length) {
    await reg.items.getById(existing[0].Id).update({ LastSyncStatus: 'Error', LastSyncError: `${p.Reason}${p.Details ? ': ' + p.Details : ''}` } as any);
  }
}

export async function seedProjectSiteFacts(context: BaseComponentContext, siteListTitle: string, ProjectId: string, fields: Record<string, any>): Promise<number> {
  const sp = createSharePointClient({ spfxContext: context });
  const list = sp.web.lists.getByTitle(siteListTitle);
  const existing = await list.items.filter(`ProjectId eq '${ProjectId}'`).top(1)();
  const payload = { ProjectId, SchemaVersion: '1', ...fields } as any;
  if (existing.length) {
    const item = list.items.getById(existing[0].Id);
    const meta = await item.select('Id', 'OData__UIVersionString')();
    await item.update(payload, meta['odata.etag']);
    return meta.Id;
  } else {
    const res = await list.items.add(payload);
    return res.data.Id;
  }
}


