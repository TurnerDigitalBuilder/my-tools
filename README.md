# My Tools Workspace

This repository hosts a suite of Microsoft 365 helper utilities that intentionally share a unified layout, color system, and workflow. Use the guidance below when updating existing tools or introducing new experiences so that everything in the workspace feels cohesive.

## Available Tools & Templates
- **Organization License Viewer** (`/org-license-viewer`): Visualize reporting lines, highlight ChatGPT license coverage, and explore team metrics with collapsible chart controls.
- **Profile Data Appender** (`/user-data-appender`): Upload CSV/Excel files, append Microsoft 365 profile attributes, and export enriched lists with progress and status messaging.
- **Column Comparison Tool** (`/column-compare`): Compare columns from two CSV/Excel files to surface overlaps, gaps, and duplicates with configurable matching rules. Export filtered CSV or Excel files for shared values or rows unique to each source. This tool is fully offline and intentionally omits the Microsoft Graph token panel.
- **Tool Starter Template** (`/tool-template`): A baseline layout with Microsoft Graph token handling, collapsible input/settings panels, and a neutral main canvas ready for custom visualizations.

## Shared Design System
### Layout Shell
- Every tool page loads `../shared/styles.css` and uses the shared header plus a two-column body (`.setup-panel` on the left, `.viz-container` on the right).
- The header always shows the tool title, a **Home** button linking back to `/index.html`, and the dark-mode toggle.
- The left panel is fixed at 380px wide, scrolls independently, and should keep its structure lightweight so the main canvas remains the visual focus.

### Setup Panel Standards
1. **Microsoft Graph Token** (only include when the tool makes Graph requests)
   - Tools that do not need Microsoft Graph should skip the token panel entirely to keep the setup area focused on required inputs.
   - When a tool does rely on Graph, include an input with `id="graphToken"`, a hint element with `id="tokenStatus"`, and call `GraphTokenManager.initialize({ onTokenChange: token => AppUI.updateTokenStatus('tokenStatus', token) });` inside your tool script.
   - Collapse the token section on page load with `AppUI.setSectionCollapsed('tokenSection', true);` so it stays out of the way after configuration.
2. **Load Data / Inputs** (collapsed on load)
   - Pair file pickers, textarea inputs, or API triggers with contextual helper text.
   - After a successful load, call `AppUI.setSectionCollapsed('loadSection', true);` so users can focus on downstream controls.
3. **Additional Controls**
   - Organize filters, toggles, and advanced settings into logical sections. Use `.wizard-step` blocks for consistent spacing.

### Content Area Patterns
- The shared `.viz-container` provides a default 2rem gutter, scroll behavior, and neutral background. Override `padding` locally when a tool needs a full-bleed canvas (see the license viewer and data appender for examples).
- Reuse existing utility classes: `.summary-grid`, `.results-table`, `.status-message`, and `.button-row` already handle spacing and responsive behavior.
- Keep headline hierarchy consistent (`h2` for primary sections, `h3` inside panels) and prefer semantic tables or lists over ad-hoc div grids.

### Color & Spacing Tokens
`shared/styles.css` exposes a small set of design tokens—reuse them rather than inventing new hex values:
- Surfaces: `--surface`, `--surface-elevated`, `--surface-alt`.
- Text: `--text-strong` (body copy) and `--text-subtle` (supporting text).
- Borders & Shadows: `--border`, `--border-subtle`, `--shadow`, `--shadow-sm`, `--shadow-xs`.
- Accents: `--primary`, `--primary-dark`, `--accent-muted`, `--accent-strong`, plus `--success`, `--warning`, and `--danger` for statuses.

Light and dark themes automatically remap those variables, so leaning on the shared tokens keeps every tool visually aligned and accessible.

## Microsoft Graph Token Workflow
- For tools that call Microsoft Graph, bootstrap token handling inside your tool script (not from inline HTML) using `GraphTokenManager.initialize`.
- Update all dependent controls inside the `onTokenChange` callback. Graph-enabled tools should:
  - Call `AppUI.updateTokenStatus` to refresh the hint under the token input.
  - Store the trimmed token in module state for validation and API calls.
  - Trigger any derived UI (e.g., detected tenant domain, enabled/disabled buttons).
- When a major workflow finishes (data fetched, comparison completed, etc.), collapse earlier setup sections via `AppUI.setSectionCollapsed` to guide the user forward.

## Landing Page Standards
- Tool cards live on `/index.html` and use the shared `tool-card`, `meta-grid`, and `meta-item` classes from `shared/home.css`.
- Each card lists the tool name, a concise description, “Data” and “Features” callouts, and a primary button that opens the tool.
- The landing page token card is the single place to seed the shared Microsoft Graph token—new tools should rely on that same storage.
- If a new tool does not require Graph data, do not surface an extra token panel; the landing page storage is enough for other tools that need it.

## Template as a Starting Point
The `/tool-template` folder demonstrates the preferred wiring:
- Collapsible token, input, and settings panels that already call `GraphTokenManager.initialize` and `AppUI.updateTokenStatus` when your feature needs Graph.
- Sample controls, toggle rows, and status messaging patterns ready to swap for real logic.
- A neutral `.template-state` section that makes it easy to drop in charts, tables, or cards.
Start new tools by copying this folder and replacing the placeholder logic while keeping the structure intact.

## Development Practices
- Keep each tool self-contained (HTML, JS, CSS). Shared utilities belong in `shared/` so improvements cascade everywhere.
- Favor semantic HTML, labelled controls, and aria-friendly status regions (`role="status"`, `aria-live="polite"`).
- Test light and dark modes, plus edge cases like empty data sets, to verify contrast and layout resilience.
- Update this README whenever a new tool ships or shared conventions evolve so future work stays aligned.
