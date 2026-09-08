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

### Decision
- Added `transparentCard?: boolean` to `IContentBlock` data contract in [IContainerModels.ts](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/models/IContainerModels.ts).
- Exposed the setting as a toggle in both the in-place [CardEditDialog.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/CardEditDialog.tsx) ("Transparent card (layout container only)") and the SPFx Property Pane sidebar in [FullWidthContainerWebPart.ts](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/FullWidthContainerWebPart.ts).
- In [BlockRenderer.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/BlockRenderer.tsx), when `transparentCard` is enabled:
  - Clears `border`, `box-shadow`, and `background-color` in published mode.
  - In Edit Mode, renders an unobtrusive dashed indicator (`rgba(0, 144, 220, 0.45)`) so authors can still easily hover, drag resize boundaries, click the card settings toolbar, and use (+) insertion bars.
  - Removes outer container padding on inner composable items for flush layout alignment.

