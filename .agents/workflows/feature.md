---
description: Implement new SPFx capability, UI component, or property pane binding
---

# New SPFx Feature Implementation ($1)

Target: Build the requested feature for `$1` adhering to Microsoft and Fluent UI 2 standards.

1. **Contract & Dependency Discovery (CodeGraph MCP):**
   - Call `codegraph_get_symbol_info` on relevant parent components or data contracts.
   - If introducing shared interfaces or callbacks, check naming parity across existing models.
   - For file resolution across the project, use `everything_search` scoped with `path:D:\Playbook\spfx-fullwidth-container`.

2. **Property Pane & UI Architecture Selection:**
   Classify UI controls into their mandatory layer before writing code:
   - **SPFx Base Tier:** Use `@microsoft/sp-property-pane` for standard out-of-the-box fields (`PropertyPaneTextField`, `PropertyPaneToggle`, `PropertyPaneDropdown`).
   - **Specialised Property Fields Tier:** Use `@pnp/spfx-property-controls` for rich SharePoint picker controls (`PropertyFieldFilePicker`, `PropertyFieldCollectionData`, `PropertyFieldPeoplePicker`).
   - **Custom Pane Surfaces Tier:** Use `@fluentui/react-components` (v9) with Griffel `makeStyles` and `<Portal>` mounting (`z-index: 1000000`, `tokens.shadow16`) for bespoke canvas overlays, custom editors, and flyouts.

3. **Single-Pass Context Ingestion:**
   - Ingest the target parent or host file in a single `view_file` pass (do not chunk).
   - Verify context isolation: presentational children must receive callbacks or scoped interfaces, never raw `WebPartContext`.

4. **Feature Construction & Fluent UI 2 Parity:**
   - Implement the feature using `@fluentui/react-components` (v9) tokens and Griffel `makeStyles`.
   - Wrap floating overlays, dropdowns, pickers, and toolbars in `<Portal>` mounted at `z-index: 1000000`.
   - Include loading (`<Spinner />` or `<Skeleton />`), error fallback, and zero-data states for async data paths.
   - Preserve all existing code, comments, and safety prompts ("Active Handshake", "Stay on this Tab").
   - Output the complete, un-truncated module.

5. **Registries & Build Verification:**
   - Run `npx tsc --noEmit` to confirm zero TypeScript compilation errors.
   - Surgically update `feature.JSON` (matrix) and `ProjectStructure.JSON` using `replace_file_content`.
   - Document layout/architectural choices in `DECISIONS.md`.