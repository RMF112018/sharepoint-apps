# Accessibility (A11y) Guidelines

Applies to all apps in apps/* . Aligns with WCAG 2.1 AA and the repo SPFx ruleset.

## Core requirements
- Keyboard: All interactive elements must be reachable and operable via Tab / Shift+Tab and Enter/Space.
- Focus: Always render a visible focus ring. Manage focus on dialogs/panels (focus first interactive control) and return focus on close.
- ARIA:
  - Use native elements where possible (button, a, input). If custom, add proper role, aria-* props, and keyboard handlers.
  - Provide aria-label or aria-labelledby for icons-only buttons.
  - Use aria-live (polite/assertive) regions to announce async status updates.
- Contrast: Ensure color contrast ≥ 4.5:1 for text and 3:1 for UI components.
- Forms: Associate labels with inputs. Surface errors inline and to screen readers (e.g., aria-invalid, aria-describedby).
- Images: Provide meaningful alt text; use empty alt for decorative images.

## Patterns
- Announce async states with a visually hidden aria-live="polite" region; update text on load/success/error.
- Manage focus for popups/panels: trap focus within, restore focus to the trigger on close.
- Disable controls while loading to prevent duplicate submissions; ensure disabled state is conveyed (aria-disabled).

## Lint & testing
- Lint: eslint-plugin-jsx-a11y is enabled repo-wide.
- Tests: Use @testing-library/react + @testing-library/jest-dom to assert roles, names, and focus behavior.

## Review checklist
- Tab order logical? Focus visible at all times?
- Keyboard equivalents exist?
- ARIA props valid and necessary? No redundant roles on native elements?
- Live regions used for async updates?
- Labels and errors properly announced?
