---
description: Validate implemented features against a defined test scope, type integrity, and defensive design standards
---

# Feature Validation & Testing ($1)

Target & Validation Scope: Evaluate `$1` against the workspace rules and any explicit scope criteria provided in the prompt.

1. **Static Type & Compilation Audit:**
   - Run `npx tsc --noEmit` to verify type safety and interface adherence across the workspace.
   - If any errors exist, output the exact compiler errors and abort further testing until resolved.

2. **AST Contract Verification (CodeGraph MCP):**
   - Query `codegraph_get_symbol_info` on the target component or service.
   - Verify that all exported signatures match the consumer expectations found via `codegraph_find_references`.
   - Confirm context isolation: verify that child components do not take `WebPartContext` as a direct prop.

3. **Control Tier & Design Invariant Audit:**
   Inspect the target file in a single `view_file` pass to verify:
   - **Property Field Governance:** Standard inputs use `@microsoft/sp-property-pane`, rich fields use `@pnp/spfx-property-controls`, and bespoke flyouts use `@fluentui/react-components` (v9) Griffel tokens.
   - **Portal Stacking:** Floating surfaces, popovers, and pickers are wrapped in `<Portal>` with `z-index: 1000000`.
   - **Lifecycle Cleanup:** All `useEffect` hooks with event listeners (`resize`, `scroll`, `pointerdown`) or timers return an explicit cleanup function.
   - **Request Abort:** All async SharePoint/Graph fetch operations bind to an `AbortController.signal`.
   - **Null & Undefined Safety:** Optional chaining (`?.`) and nullish coalescing (`??`) are applied to all dynamic payload properties.
   - **Async UI Triad:** Explicit UI coverage is implemented for loading (`<Spinner />` or `<Skeleton />`), graceful error fallbacks, and zero-data states.
   - **Code Integrity:** Confirm no commented-out code, existing features, or safety warnings ("Active Handshake") were accidentally removed.

4. **Validation Scope Review & Pass/Fail Matrix:**
   - Evaluate each criterion defined in the user's validation scope.
   - Output a clean markdown table showing: `Criterion`, `Status` (PASS / FAIL / WARN), and `Evidence / File Location`.
   - If any item fails, provide the specific remediation instructions.