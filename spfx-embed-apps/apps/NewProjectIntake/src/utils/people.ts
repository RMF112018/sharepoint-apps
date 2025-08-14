import type { BaseComponentContext } from '@microsoft/sp-component-base';
import { createSharePointClient } from '@hbi/sp-client';

export interface PersonSuggestion {
  key: string;
  name: string;
  email?: string;
  upn?: string;
  photoUrl?: string;
}

function storageKey(context: BaseComponentContext): string {
  const user = context.pageContext?.user?.loginName || 'anon';
  const site = context.pageContext?.site?.id?.toString() || 'site';
  return `npi-recent-people-${site}-${user}`;
}

export async function getRecentPeople(context: BaseComponentContext): Promise<PersonSuggestion[]> {
  try {
    const raw = typeof window !== 'undefined' ? window.localStorage.getItem(storageKey(context)) : null;
    if (raw) {
      const parsed = JSON.parse(raw) as PersonSuggestion[];
      if (Array.isArray(parsed)) return parsed.slice(0, 8);
    }
  } catch {}
  try {
    const client = await (context.msGraphClientFactory as any).getClient('3');
    const res = await client.api('/me/people').top(8).get();
    const items: PersonSuggestion[] = (res.value ?? []).map((p: any) => ({
      key: p.userPrincipalName || p.scoredEmailAddresses?.[0]?.address || p.id,
      name: p.displayName || p.scoredEmailAddresses?.[0]?.address || '',
      email: p.scoredEmailAddresses?.[0]?.address,
      upn: p.userPrincipalName
    }));
    cacheRecentPeople(context, items);
    return items;
  } catch {
    return [];
  }
}

function cacheRecentPeople(context: BaseComponentContext, people: PersonSuggestion[]): void {
  try {
    if (typeof window === 'undefined') return;
    window.localStorage.setItem(storageKey(context), JSON.stringify(people.slice(0, 20)));
  } catch {}
}

export async function searchPeople(context: BaseComponentContext, query: string): Promise<PersonSuggestion[]> {
  const q = query.trim();
  if (!q) return getRecentPeople(context);
  try {
    const client = await (context.msGraphClientFactory as any).getClient('3');
    const res = await client.api('/me/people').search(q).top(10).get();
    const items: PersonSuggestion[] = (res.value ?? []).map((p: any) => ({
      key: p.userPrincipalName || p.scoredEmailAddresses?.[0]?.address || p.id,
      name: p.displayName || p.scoredEmailAddresses?.[0]?.address || '',
      email: p.scoredEmailAddresses?.[0]?.address,
      upn: p.userPrincipalName
    }));
    return items;
  } catch {
    const sp = createSharePointClient({ spfxContext: context });
    const safe = q.replace(/'/g, "''");
    const users = await sp.web.siteUsers
      .filter(`substringof('${safe}',Title) or substringof('${safe}',Email) or substringof('${safe}',LoginName)`).top(10)();
    return users.map((u: any) => ({ key: u.LoginName, name: u.Title, email: u.Email, upn: u.LoginName }));
  }
}

export async function ensureSharePointUserIds(context: BaseComponentContext, upnsOrEmails: string[]): Promise<number[]> {
  const sp = createSharePointClient({ spfxContext: context });
  const ids: number[] = [];
  for (const id of upnsOrEmails) {
    try {
      const ensured: any = await sp.web.ensureUser(id);
      const spId: number | undefined = ensured?.data?.Id ?? ensured?.Id;
      if (typeof spId === 'number') ids.push(spId);
    } catch {}
  }
  return ids;
}


