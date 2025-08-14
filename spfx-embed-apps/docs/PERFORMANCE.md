# Performance Guidelines

Targets: TTI ≤ 2s on modern pages; bundle ≤ 250KB gzip per app (core UI).

## Data fetching
- Always use  and  with minimal fields.
- Use  on indexed columns; avoid client-side filtering for large lists.
- Page results (top/skip); do not render >5k items; virtualize long lists.
- Batch requests when performing >5 operations.

## Rendering
- Lazy-load heavy components (charts, editors) and optional panels.
- Debounce input handlers ≥150ms to reduce re-renders.
- Cancel in-flight requests on unmount to avoid stale updates.

## Bundling
- Code-split optional routes/panels.
- Keep shared logic in libs/*; avoid duplicating heavy deps per app.
- Use the provided CI budget checks (size-limit/webpack hints) if configured.

## Checklist
- Queries use  /  minimal fields?
- Server-side  and paging used on large lists?
- Heavy components lazy-loaded?
- Input handlers debounced?
- Requests aborted on unmount?
