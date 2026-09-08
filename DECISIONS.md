# Architectural Decisions Log

## DEC-001: Interactive Filter Buttons & Customisable Process Model (FEAT-031)

### Context
Users require the ability to:
1. Provide button-based filtering where pill buttons with configurable titles act as card filters (matching mockups with pills such as "Programme advisory", "Cost and commercial management", "Sustainability", etc.).
2. Build customisable process models / chevron stage flow diagrams (e.g., Stage 1: Shape, Stage 2: Plan, Stage 3: Source, Stage 4: Deliver, Stage 5: Realise) with configurable step count, titles, descriptions, and capability badges.
3. Steps must support two click behaviours: navigating to a URL or filtering cards down.

### Decision
- **Composable Sub-Item Integration**: Extended `ICardItemType` with `'filterButtons'` and `'processModel'`, allowing authors to add them via [ToolboxModal.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/ToolboxModal.tsx) and configure them in [CardItemPropertyEditor.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/CardItemPropertyEditor.tsx).
- **Design Alignment**: Built [FilterButtonsRenderer.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/FilterButtonsRenderer.tsx) and [ProcessModelRenderer.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/ProcessModelRenderer.tsx) strictly with Fluent UI 2 Griffel tokens and Turner & Townsend palette (`#1E4479` deep blue active fill, `#001436` hover state, `#F8F6F2` stage containers).
- **Decoupled Card Filtering Dispatch**: Dispatched custom events (`dashboard:card-filter-apply`) consumed by both [TabsContainer.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/TabsContainer.tsx) and [AccordionContainer.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/AccordionContainer.tsx) to filter card items cleanly without requiring prop drilling through presentational subcomponents.

## DEC-002: Transparent Card Layout Container (FEAT-032)

### Context
Users requested the ability to disable visual card chrome (card backgrounds, borders, shadows, elevation) so that the card acts purely as a layout container slot on the section canvas, seamlessly hosting elements such as filter button rows, process models, or media without a boxed card frame.

## DEC-003: Composable Header & Section Top Content, Alignment Controls & Portal Mounting (FIX-001)

### Context
1. Authors required the ability to add composable content items (text, buttons, filter buttons, process models, media, etc.) directly into the web part top header (above sections) and at the top of each section (above the cards grid).
2. Content sections and search bar needed alignment options (Left, Centre, Right).
3. The tabs/accordion layout mode switcher needed to be permanently anchored to the top right of the web part header.
4. The blue outline/border in Edit Mode needed to be removed.
5. In-place property dialogs ([CardEditDialog.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/CardEditDialog.tsx) and [SectionEditDialog.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/SectionEditDialog.tsx)) were being clipped by the right-hand SharePoint page rail and overlapping card contents.

### Decision
- **Composable Content Areas**: Created [ComposableContentSection.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/ComposableContentSection.tsx), allowing authors to insert and edit any of the 15 composable item types at both the top of the web part (`headerContentItems`) and top of sections (`section.topContentItems`).
- **Alignment Controls**:
  - Added `alignment?: 'left' | 'center' | 'right'` to `ICardItem` and exposed an alignment dropdown in [CardItemPropertyEditor.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/CardItemPropertyEditor.tsx).
  - Applied flex justification and text alignment across [BlockRenderer.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/BlockRenderer.tsx), [FilterButtonsRenderer.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/FilterButtonsRenderer.tsx), and [ProcessModelRenderer.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/ProcessModelRenderer.tsx).
  - Added `searchAlignment` property in Property Pane and aligned search and term filter dropdowns accordingly.
- **Top-Right Mode Switcher**: Restructured [FullWidthContainer.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/FullWidthContainer.tsx) header with flex row layout, keeping title/subtitle on the left and the tabs/accordion switcher pinned to the top right.
- **Removed Blue Border**: Suppressed dashed blue border on `editBanner` and neutralised canvas selection styling on container root.
- **Fluent UI 2 Portal & Responsive Side Panels**: Mounted both `CardEditDialog` and `SectionEditDialog` in `<Portal>` at `z-index: 1000000`, added right margin of `48px` to clear the SharePoint page editing rail, set sticky footers, and used auto-wrapping two-column grids so action buttons and controls are never clipped.

