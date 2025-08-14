export type TelemetryEventName =
  | 'load'
  | 'submit_success'
  | 'submit_error'
  | 'validation_error'
  | 'localSaved'
  | 'syncQueued'
  | 'syncStart'
  | 'syncSuccess'
  | 'syncError'
  | 'conflictDetected'
  | 'resolved';

export interface TelemetryEvent {
  name: TelemetryEventName;
  atMs: number;
  durationMs?: number;
  details?: Record<string, string | number | boolean | undefined>;
  correlationId: string;
}

function generateCorrelationId(): string {
  if (typeof crypto !== 'undefined' && 'getRandomValues' in crypto) {
    const buf = new Uint8Array(16);
    crypto.getRandomValues(buf);
    // RFC4122 v4
    buf[6] = (buf[6] & 0x0f) | 0x40;
    buf[8] = (buf[8] & 0x3f) | 0x80;
    const hex: string[] = [];
    const toHex2 = (v: number): string => {
      const s = v.toString(16);
      return s.length === 1 ? '0' + s : s;
    };
    for (let i = 0; i < buf.length; i++) {
      const v: number = buf[i];
      hex.push(toHex2(v));
    }
    return `${hex[0]}${hex[1]}${hex[2]}${hex[3]}-${hex[4]}${hex[5]}-${hex[6]}${hex[7]}-${hex[8]}${hex[9]}-${hex[10]}${hex[11]}${hex[12]}${hex[13]}${hex[14]}${hex[15]}`;
  }
  return `${Date.now().toString(16)}-${Math.random().toString(16).slice(2, 10)}`;
}

function getOrCreateCorrelationId(): string {
  try {
    const key = 'hbi_correlationId';
    const existing = sessionStorage.getItem(key);
    if (existing) return existing;
    const id = generateCorrelationId();
    sessionStorage.setItem(key, id);
    return id;
  } catch {
    return generateCorrelationId();
  }
}

class TelemetryClient {
  private readonly correlationId: string;

  public constructor() {
    this.correlationId = getOrCreateCorrelationId();
  }

  public getCorrelationId(): string {
    return this.correlationId;
  }

  public track(name: TelemetryEventName, details?: TelemetryEvent['details'], durationMs?: number): void {
    const ev: TelemetryEvent = {
      name,
      atMs: Date.now(),
      durationMs,
      details,
      correlationId: this.correlationId
    };
    this.send(ev);
  }

  private send(ev: TelemetryEvent): void {
    try {
      if (typeof navigator !== 'undefined' && 'sendBeacon' in navigator) {
        const blob = new Blob([JSON.stringify(ev)], { type: 'application/json' });
        // Placeholder endpoint. Replace with real endpoint when available.
        navigator.sendBeacon('/telemetry', blob);
      } else {
        // Fallback to console for dev/test. No PII included.
        // eslint-disable-next-line no-console
        console.debug('[telemetry]', ev);
      }
    } catch {
      // swallow
    }
  }
}

export const telemetry = new TelemetryClient();

export function nowMs(): number {
  try {
    return (typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now();
  } catch {
    return Date.now();
  }
}

// Lightweight client-side metrics storage (non-PII)
const DURATIONS_KEY = 'npi-sync-durations';
const FAILURES_KEY = 'npi-sync-failures';

export function recordSyncDuration(ms: number): void {
  try {
    const raw = localStorage.getItem(DURATIONS_KEY);
    const arr: number[] = raw ? JSON.parse(raw) : [];
    arr.push(ms);
    while (arr.length > 200) arr.shift();
    localStorage.setItem(DURATIONS_KEY, JSON.stringify(arr));
  } catch {}
}

export function getSyncPercentiles(): { p50?: number; p95?: number } {
  try {
    const raw = localStorage.getItem(DURATIONS_KEY);
    const arr: number[] = raw ? JSON.parse(raw) : [];
    if (arr.length === 0) return {};
    const sorted = [...arr].sort((a, b) => a - b);
    const p = (q: number) => sorted[Math.min(sorted.length - 1, Math.floor(q * sorted.length))];
    return { p50: p(0.5), p95: p(0.95) };
  } catch { return {}; }
}

export interface FailureSample { at: string; status?: number; reason?: string }

export function recordSyncFailureSample(sample: FailureSample): void {
  try {
    const raw = localStorage.getItem(FAILURES_KEY);
    const arr: FailureSample[] = raw ? JSON.parse(raw) : [];
    arr.push(sample);
    while (arr.length > 100) arr.shift();
    localStorage.setItem(FAILURES_KEY, JSON.stringify(arr));
  } catch {}
}

export function getRecentFailures(): FailureSample[] {
  try {
    const raw = localStorage.getItem(FAILURES_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch { return []; }
}


