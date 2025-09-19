# My Tools Workspace

This repository hosts a collection of Microsoft 365 helper utilities that share a unified look and feel. Use the guidelines below when extending the workspace with new tools so that every experience feels cohesive.

## Landing Page Layout
- Cards should be wide rectangles that highlight the tool name, a short description, and structured “Data” and “Features” callouts.
- Use the existing `tool-card`, `meta-grid`, and `meta-item` classes from `shared/home.css` to keep the layout consistent.
- Each card should include a primary button that links to the tool’s index page.

## Tool Page Conventions
- Every tool page lives in its own folder and loads the shared styles from `../shared/styles.css`.
- The header must include a **Home** button linking back to `/index.html` as well as the dark-mode toggle button.
- The left-hand panel should follow this order:
  1. **Microsoft Graph Token** collapsible section (collapsed by default).
  2. **Load Data** collapsible section (collapsed by default).
  3. Any additional filters, controls, or output sections.
- Use the shared `AppUI.setSectionCollapsed` helper to collapse sections on page load.
- Whenever a workflow completes a major action (e.g., data fetch or upload), auto-collapse the load section so the user can focus on downstream controls.

## Shared Resources
- Global styles live in `shared/styles.css`; landing-page-specific styles live in `shared/home.css`.
- Token handling is centralized in `shared/graph-token.js`. Always reuse this module to persist and broadcast token updates between tools.
- Any reusable UI helpers should be added to `shared/ui.js` to avoid duplication.

## Development Tips
- Keep new tool directories self-contained with their own `index.html`, supporting scripts, and styles.
- Favor semantic HTML and accessible patterns (e.g., labelled inputs, descriptive button text).
- Test new tools with both light and dark mode enabled to ensure adequate contrast.

Following these conventions will ensure new utilities feel native to the workspace and can inherit improvements made to the shared components.
