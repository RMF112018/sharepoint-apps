export type Delta = Record<string, unknown>;

function isEqual(a: unknown, b: unknown): boolean {
  if (Array.isArray(a) && Array.isArray(b)) {
    if (a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++) {
      if (!isEqual(a[i], b[i])) return false;
    }
    return true;
  }
  return a === b;
}

export function computeDelta(previous: Record<string, unknown>, current: Record<string, unknown>, includeKeys?: string[]): Delta {
  const keys = includeKeys ?? Array.from(new Set([...Object.keys(previous || {}), ...Object.keys(current || {})]));
  const delta: Delta = {};
  for (const k of keys) {
    const prev = previous ? previous[k] : undefined;
    const cur = current ? current[k] : undefined;
    if (!isEqual(prev, cur)) {
      delta[k] = cur;
    }
  }
  return delta;
}

export function chunkDelta(delta: Delta, chunkSize = 100): Delta[] {
  const entries: Array<[string, unknown]> = [];
  for (const k in delta) { if (Object.prototype.hasOwnProperty.call(delta, k)) entries.push([k, (delta as any)[k]]); }
  const chunks: Delta[] = [];
  for (let i = 0; i < entries.length; i += chunkSize) {
    const slice = entries.slice(i, i + chunkSize);
    const piece: Delta = {};
    for (const [k, v] of slice) piece[k] = v;
    chunks.push(piece);
  }
  return chunks.length ? chunks : [delta];
}


