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

---

## DEC-003: Search Bar Permanent Top-Right Relocation & Editable Placeholder (FEAT-035)

### Context
The search input was previously positioned inside a mid-page `searchAndFiltersRow`, causing it to disappear when section-level content pushed the layout down. Authors also had no way to change the default placeholder text (`Filter items, tags, GBP...`) without a code change.

### Decision
- **Relocation**: The `<Input>` is moved out of the mid-page row and into `headerTopRight` — the same flex container that holds the Tabs/Accordion mode switcher — so it remains permanently visible regardless of content scroll position.
- **Editable Placeholder**: In edit mode, an `EditRegular` icon button is rendered alongside the search input. Clicking it opens a Fluent UI 2 `<Popover>` containing a labelled `<Input>` pre-filled with the current placeholder. Pressing Enter commits the change via the `onSearchPlaceholderChange` callback; pressing Escape reverts.
- **Prop Surface**: `searchPlaceholder` and `onSearchPlaceholderChange` are added to `IFullWidthContainerProps`. The web part root passes both from `this.properties.searchPlaceholder`.
- **Preview Parity**: `preview/index.html` mirrors this by adding `searchPlaceholder` to the state object and rendering a `window.prompt`-based edit path in edit mode.

---

## DEC-004: Dropdown as a Composable Content Item (FEAT-036)

### Context
Authors need a native dropdown control that can be placed as a composable item inside cards, sections, or the dashboard header to drive card-level filtering — the same pattern established by `filterButtons` and `processModel`. The existing Term Store filter bar operates at a global scope; this new item operates at the card/section scope.

### Decision
- **Model Extension**: `ICardItem` gains `dropdownLabel`, `dropdownPlaceholder`, `dropdownTermSetName`, `dropdownOptions: Array<{ label, value }>`, and `selectedDropdownValue` fields in [IContainerModels.ts](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/models/IContainerModels.ts).
- **Rendering**: The Fluent UI 2 `<Dropdown>` + `<Option>` components are used in both [ComposableContentSection.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/ComposableContentSection.tsx) and [BlockRenderer.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/BlockRenderer.tsx). On selection change, a `dashboard:card-filter-apply` custom event is dispatched with the `optionValue`.
- **Property Editor**: `CardItemPropertyEditor.tsx` handles `case 'dropdown'` with label, placeholder, term set name, and a reorderable static options list (Add / Up / Down / Delete).
- **Filter-Host Safety**: Both `TabsContainer.tsx` and `AccordionContainer.tsx` extend their `isFilterHost` guard to include `it.type === 'dropdown'`, preventing the hosting card from being hidden by its own filter event.
- **Static Options as Primary / Term Set as Future Enhancement**: The term set name field is stored and serialised, but live term population from the SharePoint taxonomy API is deferred to a future sprint. Static options serve as the primary data source.

---

## DEC-005: Capabilities Card Content Part Architecture (FEAT-039)

### Context
Authors require a structured 'Capabilities' content part that can be added to any card (or header/section top slot) to display programme capabilities:
1. Dynamic header reading "Capabilities applied here X - [secondary label]" where X strictly reflects the live count of capabilities added.
2. An editable section name and optional secondary description note.
3. Multiple capability mini-cards within the parent card, each displaying an editable title, contextual subtitle line, configurable tags/badges, and an action link ("Open the standard →").
4. Ability to add unlimited capabilities, reorder them with Up/Down controls, and remove them dynamically.

### Decision
- **Model Extension**: Added `ICapabilityTag`, `ICapabilityItem`, and `'capabilities'` to `ICardItemType` in [IContainerModels.ts](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/models/IContainerModels.ts). Stored `capabilitiesSectionLabel`, `capabilitiesSecondaryLabel`, and `capabilities: ICapabilityItem[]` directly on `ICardItem`.
- **Component Separation**: Implemented [CapabilitiesRenderer.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/CapabilitiesRenderer.tsx) using 100% Fluent UI 2 Griffel tokens (`tokens.colorNeutralBackground1`, `tokens.colorBrandForeground1`, `tokens.shadow2`). Mini-cards use a responsive flex-wrap layout (`flex: 1 1 210px`, `maxWidth: 300px`) that naturally wraps within any card grid column span.
- **Semantic Tag Color Mapping**: Built `getTagStyle()` utility mapping common capability tag types (`templates` -> soft cyan/blue `#CCE9F8`, `mandatory` -> soft coral/red `#FDE7E9`, `insight` -> soft mint/green `#DFF6DD`, `learning`/`notes` -> warm wheat `#FFF4CE`, `infra` -> TT Mushroom `#F2EEE7`) matching corporate Turner & Townsend brand standards.
- **Deep Property Editor**: Built rich configuration UI inside [CardItemPropertyEditor.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/CardItemPropertyEditor.tsx) with a live banner previewing the derived count header, Add/Move/Delete capability cards, and inline tag management.
- **Insertion & Rendering Parity**: Integrated into [ToolboxModal.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/ToolboxModal.tsx) (with `BoardRegular` icon), [BlockRenderer.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/BlockRenderer.tsx), and [ComposableContentSection.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/ComposableContentSection.tsx).
