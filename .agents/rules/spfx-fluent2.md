---
trigger: always_on
---

---
trigger: always_on
description: Enterprise SPFx development standards, Fluent UI 2 Griffel rules, and token-efficient MCP governance
---

# SharePoint Framework (SPFx) & Fluent UI 2 Standards

## Code Intelligence via CodeGraph MCP
- You MUST use `codegraph-spfx` (`codegraph_get_call_graph`, `codegraph_find_references`, `codegraph_get_symbol_info`) for symbol tracing, dependency mapping, and consumer lookups.
- Native multi-file regex grep and sequential file scans are strictly FORBIDDEN.
- AST Two-Step: Query `codegraph_get_symbol_info` first to verify signatures before pulling caller graphs or full component files into context.

## Property Pane Architecture & Control Tiering
- SPFx Base Tier: Use `@microsoft/sp-property-pane` for standard out-of-the-box property pane fields (`PropertyPaneTextField`, `PropertyPaneToggle`, `PropertyPaneDropdown`). Configuration flows strictly down via typed interfaces (`I<WebPartName>Props`). React components must never directly mutate `this.properties`.
- Specialised Property Fields Tier: Use `@pnp/spfx-property-controls` for rich SharePoint picker controls (`PropertyFieldFilePicker`, `PropertyFieldCollectionData`, `PropertyFieldPeoplePicker`).
- Custom Pane Surfaces Tier: Use `@fluentui/react-components` (v9) with Griffel `makeStyles` and `<Portal>` mounting (`z-index: 1000000`, `tokens.shadow16`) for bespoke canvas overlays, custom editors, and flyouts.

## Fluent UI 2 (v9) & Griffel Styling
- Component Libraries: Strictly use `@fluentui/react-components` (v9) and `@fluentui/react-icons`. Legacy `@fluentui/react` (v8) or `office-ui-fabric-react` imports are prohibited.
- Styling Engine: Use `makeStyles`, `shorthands`, and semantic `tokens` from `@fluentui/react-components`. Raw CSS/SCSS modules and hardcoded hex colours are forbidden.
- Portal Rule (Anti-Clipping): All floating toolbars, contextual popovers, dropdowns, and flyouts MUST render inside a Fluent UI `<Portal>` or `<Popover>` mounted at `z-index: 1000000` with `tokens.shadow16` elevation.
- Selection Preservation: Toolbar action buttons inside rich-text floating surfaces must include `onMouseDown={(e) => e.preventDefault()}` so editor text selection is not blurred.
- Themes: Root React tree must be wrapped in `<FluentProvider>`, bridging SharePoint's `this.context.palette` or `ThemeProvider` to `webLightTheme` or `webDarkTheme`.

## SPFx Context & Integration
- Context Isolation: Never pass `WebPartContext` directly to presentational child components. Pass only strongly typed callbacks or scoped service interfaces.
- Asset Normalisation: Always use server-relative URLs for site assets and document libraries; never hardcode absolute tenant URLs.
- Always initialize PnPjs via the SPFx factory: `spfi().using(SPFx(this.context))`.
- Access Microsoft Graph strictly via `this.context.msGraphClientFactory.getClient('3')`.
- Maintain `feature.JSON` (matrix), `ProjectStructure.JSON`, and `test_plan.md` using non-destructive replacement (`replace_file_content`).
- Record non-trivial architecture or layout updates in `DECISIONS.md`.

## Defensive Engineering & Async UX
- Resource & Listener Cleanup: Always clear timers, remove window/document event listeners (`resize`, `scroll`, `blur`, `pointerdown`), and pass `AbortController.signal` to pending PnPjs/Graph fetch requests inside `useEffect` cleanup returns.
- Null Safety & Data Resilience: Never assume SharePoint payload completeness. Enforce optional chaining (`?.`), nullish coalescing (`??`), and strict runtime schema guards on all list items, user profile properties, and PropertyPane inputs before rendering.
- Async UX States: Every asynchronous operation must account for three explicit UI states: loading (using Fluent UI 2 `<Spinner />` or `<Skeleton />`), graceful error boundaries with actionable messaging, and zero-data empty states.
- Long-Running Task Safety: Maintain and preserve the 'Active Handshake' and 'Stay on this Tab' prompts in any long-running synchronization, batch upload, or migration UI code.

## Bundle & Import Performance
- Strictly avoid barrel imports from icon packages. Import icons individually (e.g., `import { EditRegular, DeleteRegular } from '@fluentui/react-icons'`).
- Enforce dynamic code splitting using `React.lazy()` and `<Suspense>` for heavy modal pickers, custom code editors, and secondary property pane panels.

## Token & Context Optimisation
- Fast Type Verification: Validate code edits using `npx tsc --noEmit` instead of full Gulp builds. If running Gulp, filter output to errors only (`| findstr /i "error"`).
- Interface Scoping: Query `codegraph-spfx` (`codegraph_get_symbol_info`) for interface definitions rather than reading entire parent components.
- Scoped Everything Queries: Always prefix `everything_search` queries with `path:D:\Playbook\spfx-fullwidth-container`.