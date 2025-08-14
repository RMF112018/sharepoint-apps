import type { BaseComponentContext } from '@microsoft/sp-component-base';
import { v4 as uuidv4 } from 'uuid';
import syncConfig from '../config/sync.json';

export interface OutboxJob {
  SyncId: string;
  ProjectId: string;
  ChangedFields: Record<string, any>;
  SourceSiteUrl: string;
  Timestamp: string;
  Action?: 'push' | 'pull';
}

function storageKey(context: BaseComponentContext): string {
  const user = context.pageContext.user.loginName || 'anon';
  return `npi-outbox-${user}`;
}

export function enqueueOutbox(context: BaseComponentContext, job: Omit<OutboxJob, 'SyncId' | 'Timestamp'>): OutboxJob {
  const full: OutboxJob = { ...job, SyncId: uuidv4(), Timestamp: new Date().toISOString() };
  try {
    const raw = window.localStorage.getItem(storageKey(context));
    const arr: OutboxJob[] = raw ? JSON.parse(raw) : [];
    arr.push(full);
    window.localStorage.setItem(storageKey(context), JSON.stringify(arr));
  } catch {}
  return full;
}

export function readOutbox(context: BaseComponentContext): OutboxJob[] {
  try {
    const raw = window.localStorage.getItem(storageKey(context));
    return raw ? JSON.parse(raw) : [];
  } catch { return []; }
}

export function removeFromOutbox(context: BaseComponentContext, syncId: string): void {
  try {
    const raw = window.localStorage.getItem(storageKey(context));
    const arr: OutboxJob[] = raw ? JSON.parse(raw) : [];
    const filtered = arr.filter(j => j.SyncId !== syncId);
    window.localStorage.setItem(storageKey(context), JSON.stringify(filtered));
  } catch {}
}

export interface SyncResult { ok: boolean; status?: number; error?: string }

export async function tryServerSync(job: OutboxJob): Promise<SyncResult> {
  if (!syncConfig.endpointUrl) return { ok: false, error: 'no-endpoint' };
  try {
    const res = await fetch(syncConfig.endpointUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        ...(syncConfig.apiKey ? { 'x-api-key': syncConfig.apiKey } : {})
      },
      body: JSON.stringify(job)
    });
    if (!res.ok) {
      let text = '';
      try { text = await res.text(); } catch {}
      return { ok: false, status: res.status, error: text };
    }
    return { ok: true, status: res.status };
  } catch (e: any) { return { ok: false, error: (e && e.message) || 'network-error' }; }
}

export async function drainOutbox(context: BaseComponentContext): Promise<SyncResult> {
  const jobs = readOutbox(context);
  let lastError: SyncResult | undefined;
  for (const j of jobs) {
    const res = await tryServerSync(j);
    if (res.ok) {
      removeFromOutbox(context, j.SyncId);
    } else {
      lastError = res;
    }
  }
  return lastError ? lastError : { ok: true };
}


