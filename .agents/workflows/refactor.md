---
description: Fast SPFx refactor with CodeGraph AST mapping and property tiering
---

# SPFx Component Refactor ($1)

Target: Apply this checklist to `$1` (or the component specified by the user).

1. **AST & Impact Analysis (CodeGraph MCP):**
   - Call `codegraph_get_symbol_info` on `$1`.
   - If props or interfaces change, call `codegraph_find_references` to identify consumers.

2. **Single-Pass Read:**
   - Ingest the file for `$1` in a single `view_file` operation (no chunking).

3. **Architecture & Property Tier Alignment:**
   - Verify that configuration UI maps to the correct tier:
     - Base controls: `@microsoft/sp-property-pane`.
     - Specialised controls: `@pnp/spfx-property-controls`.
     - Custom flyouts/editors: `@fluentui/react-components` (v9) with Griffel `makeStyles` and `<Portal>` (`z-index: 1000000`).
   - Preserve all existing code, handlers, and safety prompts.
   - Output the complete, un-truncated module.

4. **Fast Validation:**
   - Verify typing with `npx tsc --noEmit`.
   - Record non-trivial decisions in `DECISIONS.md`.