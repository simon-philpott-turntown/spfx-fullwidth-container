---
description: Research a technical topic or feature scope and produce a structured implementation report for review
---

# Feature Research & Implementation Report ($1)

Subject / Scope: Research `$1` across the codebase and external design standards. Produce an architectural evaluation and implementation plan for user review without modifying code files.

1. **AST & Interface Discovery (CodeGraph MCP):**
   - Call `codegraph_get_symbol_info` on relevant interfaces, parent components, and utility contracts linked to `$1`.
   - Call `codegraph_find_references` to identify all dependent call sites and consumers that would be affected by the feature.
   - Query `everything_search` with `path:D:\Playbook\spfx-fullwidth-container` to map related components, state hooks, or service layers.

2. **Target Context Ingestion:**
   - Ingest candidate parent or sibling files using single-pass `view_file` reads (no window slicing).
   - Inspect existing state flows, property pane bindings (`this.properties`), and context providers.

3. **Architectural Evaluation & Design Mapping:**
   - **Control Architecture Tiering:**
     - *SPFx Base:* Map native property pane fields to `@microsoft/sp-property-pane`.
     - *Specialised Fields:* Check if complex SharePoint assets (e.g., site assets, documents) leverage `@pnp/spfx-property-controls` (e.g., `PropertyFieldFilePicker`).
     - *Custom Surfaces:* Identify bespoke property controls or in-canvas edit surfaces requiring `@fluentui/react-components` (v9) Griffel tokens and `<Portal>` mounting.
   - **Data Access:** Determine whether data flows via PnPjs (`spfi().using(SPFx(this.context))`) or Microsoft Graph (`msGraphClientFactory.getClient('3')`).
   - **Context Isolation:** Define the typed interfaces or callback props to pass down so presentational components avoid direct `WebPartContext` coupling.
   - **Defensive Strategy:** Identify required async states (loading, error boundaries, empty collections) and unmount cleanups (`AbortController`).

4. **Deliverable (Structured Implementation Report):**
   Output a formal report formatted as follows:
   - **Executive Summary & Scope:** What is being added and what is explicitly out of scope.
   - **Property Control Strategy:** Explicit classification across `@microsoft/sp-property-pane`, `@pnp/spfx-property-controls`, or custom Fluent UI 2 `<Portal>` surfaces.
   - **Impacted Files & Dependencies:** Affected components, registries (`feature.JSON`, `ProjectStructure.JSON`), and new modules.
   - **Component & Contract Design:** Proposed TypeScript interfaces, prop shapes, and state flow diagrams.
   - **Edge Cases & Failure Modes:** Async timeouts, missing SharePoint permissions, null data schemas.
   - **Step-by-Step Execution Plan:** Ordered sequence of implementation steps ready for approval.