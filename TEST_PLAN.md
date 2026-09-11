# SPFx Full-Width Container Web Part — Comprehensive Test Plan

## Executive Summary
This document establishes the end-to-end test plan and verification matrix for all **38 registered features** of the **SPFx Full-Width Container Web Part** (`spfx-fullwidth-container`), spanning layout engines, canvas editing, taxonomy integrations, backup pipelines, and defensive engineering standards.

---

## 1. Feature Verification Matrix (All 38 Features)

| Feature ID | Feature Name | Primary Component / File | Verification Method | Target Status |
| :--- | :--- | :--- | :--- | :--- |
| **FEAT-001** | FullBleedLayout | [FullWidthContainerWebPart.manifest.json](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/FullWidthContainerWebPart.manifest.json) | Inspect modern full-width column placement | PASS |
| **FEAT-002** | DualLayoutModes | [FullWidthContainer.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/FullWidthContainer.tsx) | Switch between Tabs and Accordion views | PASS |
| **FEAT-003** | ChildContentBlocks | [BlockRenderer.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/BlockRenderer.tsx) | Verify Card, Metric, Embed, RichText, QuickLinks | PASS |
| **FEAT-004** | FluentUI2DesignSystem | [themeBridge.ts](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/utils/themeBridge.ts) | Test dark/light mode palette inheritance | PASS |
| **FEAT-005** | PropertyPaneCustomisation | [FullWidthContainerWebPart.ts](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/FullWidthContainerWebPart.ts) | Verify 3-page property pane controls & presets | PASS |
| **FEAT-006** | InlineCanvasAndCardEditing | [RichTextEditable.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/RichTextEditable.tsx) | Inline edit titles, descriptions, and metrics | PASS |
| **FEAT-007** | ComposableInnerCardsAndHoverToolbox | [ToolboxModal.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/ToolboxModal.tsx) | Add items inside cards via (+) insertion bar | PASS |
| **FEAT-008** | SharePointGlobalTermStoreIntegration | [TaxonomyService.ts](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/services/TaxonomyService.ts) | Query Graph API termStore & render tag pills | PASS |
| **FEAT-009** | LiveDataAPIFields | [LiveDataService.ts](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/services/LiveDataService.ts) | Fetch live JSON data with interval refresh & GBP | PASS |
| **FEAT-010** | CardGridSpanningAndContainerLayout | [BlockRenderer.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/BlockRenderer.tsx) | Verify 1-4 column spans & 1-2 row spans | PASS |
| **FEAT-011** | IndependentCardHeightAndEqualRowScaling | [BlockRenderer.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/BlockRenderer.tsx) | Test 'auto' (fit-content) vs 'equal' height | PASS |
| **FEAT-012** | FluentUI2SkeletonShimmers | [LiveDataRenderer.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/LiveDataRenderer.tsx) | Inspect Skeleton shimmer during async fetch | PASS |
| **FEAT-013** | FluentUI2ToastNotifications | `preview/index.html` | Trigger save/resize toasts | PASS |
| **FEAT-014** | DashboardJSONPortability | [FullWidthContainerWebPart.ts](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/FullWidthContainerWebPart.ts) | Export and import dashboard JSON templates | PASS |
| **FEAT-015** | TurnerAndTownsendBrandPalette | [FloatingTextToolbar.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/FloatingTextToolbar.tsx) | Verify TT July 2025 colors and tints | PASS |
| **FEAT-016** | LivePhysicalCardResizePreviewAndDragIsolation | [dragIsolation.ts](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/utils/dragIsolation.ts) | Resize card boundaries without triggering SP drag | PASS |
| **FEAT-017** | MiniCollapsibleCardTags | [BlockRenderer.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/BlockRenderer.tsx) | Toggle collapsible tags badge on cards | PASS |
| **FEAT-018** | VisualBrandColorPopovers | [BrandColorPickerPopover.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/BrandColorPickerPopover.tsx) | Select colors with active selection rings | PASS |
| **FEAT-019** | SharePointDocumentLibraryBackupAndRestore | [DashboardStorageService.ts](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/services/DashboardStorageService.ts) | Backup snapshots to SharePoint 'Dashboards' | PASS |
| **FEAT-020** | CustomBrandSVGIconsAndDynamicStyling | [CustomSvgIconRegistry.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/CustomSvgIconRegistry.tsx) | Render custom corporate SVG glyphs | PASS |
| **FEAT-021** | PropertyPaneAccordionAndVisualIconField | [PropertyPaneIconField.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/PropertyPaneIconField.tsx) | Test collapsible property pane accordion & icon | PASS |
| **FEAT-022** | SharePointDocumentLibraryDirectRESTFix | [DashboardStorageService.ts](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/services/DashboardStorageService.ts) | Validate unencoded slashes in REST paths | PASS |
| **FEAT-023** | UniversalSvgIconStylingAndTransparentColorPicker | [BrandColorPickerPopover.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/BrandColorPickerPopover.tsx) | Select transparent checkerboard background | PASS |
| **FEAT-024** | SectionActionIconsAndUnblurredCanvasPanels | [SectionEditDialog.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/SectionEditDialog.tsx) | Edit section in non-blocking backdrop panel | PASS |
| **FEAT-025** | ContextualDefaultColorOptionsAcrossPickers | [BrandColorPickerPopover.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/BrandColorPickerPopover.tsx) | Test 'Default' swatch representation per context | PASS |
| **FEAT-026** | DualModeRichTextFormatting | [RichTextEditable.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/RichTextEditable.tsx) | Highlighted selection vs full block formatting | PASS |
| **FEAT-027** | ResilientSharePointLibrarySnapshotUpload | [DashboardStorageService.ts](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/services/DashboardStorageService.ts) | Test fallback to root library folder | PASS |
| **FEAT-028** | UnclippedFloatingTextToolbarAndCardTypography | [FloatingTextToolbar.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/FloatingTextToolbar.tsx) | Check toolbar popover positioning above target | PASS |
| **FEAT-029** | ContextualItemPropertyEditor | [CardItemPropertyEditor.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/CardItemPropertyEditor.tsx) | Edit all 15 item types with live sync | PASS |
| **FEAT-030** | OptionalComposableFilterDropdowns | [ComposableContentSection.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/ComposableContentSection.tsx) | Optional composable filter dropdowns in cards and header slots (not permanent) | PASS |
| **FEAT-031** | FilterButtonsAndProcessModel | [FilterButtonsRenderer.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/FilterButtonsRenderer.tsx) | Test pill buttons & chevron stages filtering | PASS |
| **FEAT-032** | TransparentCardLayoutContainer | [BlockRenderer.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/BlockRenderer.tsx) | Render borderless/backgroundless card slot | PASS |
| **FEAT-033** | ComposableHeaderAndSectionContent | [ComposableContentSection.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/ComposableContentSection.tsx) | Top-of-webpart & top-of-section content slots | PASS |
| **FEAT-034** | TextToolbarFocusDismissalAndPreservation | [RichTextEditable.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/RichTextEditable.tsx) | Ensure formatting ribbon disappears on blur | PASS |
| **FEAT-035** | SearchBarTopRightRelocationWithEditablePlaceholder | [FullWidthContainer.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/FullWidthContainer.tsx) | Verify search pinned top-right; inline Popover placeholder edit in edit mode | PASS |
| **FEAT-036** | DropdownFilterComposableContentItem | [ComposableContentSection.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/ComposableContentSection.tsx) | Insert dropdown item, configure options, verify filter dispatch | PASS |
| **FEAT-037** | UserProfileDiagnosticsInspection | [FullWidthContainerWebPart.ts](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/FullWidthContainerWebPart.ts) | Click author-only user profile icon in edit banner to inspect pageContext.user | PASS |
| **FEAT-038** | ProcessModelConfigurationOverhaulAndComposableDropdownFilterPane | [CardItemPropertyEditor.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/CardItemPropertyEditor.tsx) | Test pinned modal footer, rich text stage description, filter/link mode switch, human title, and taxonomy dropdown editor | PASS |

---

## 2. Detailed Test Scenarios

### Test Suite 1: Layout & Mode Switching (FEAT-001, FEAT-002, FEAT-033)
1. **Tabs vs Accordion Toggle**:
   - Verify layout toggle button is permanently anchored to the top-right corner.
   - Click "Accordion": confirm cards transition into collapsible section panels.
   - Click "Tabs": confirm cards transition into tabbed view with active underline/pill indicator.
2. **Top Content Slots**:
   - In Edit Mode, verify the `(+)` insertion bar appears above sections (web part header).
   - In Edit Mode, verify the `(+)` insertion bar appears inside each section above the card grid.
   - Add a Button and a Process Model to the web part header; confirm they render above the tabs.

### Test Suite 2: Composable Content Alignment (FEAT-031, FEAT-033)
1. **Alignment Options**:
   - Open properties for an inner item (Button, Text, Filter Buttons, Process Model).
   - Set alignment to `Left`: confirm element aligns to the left margin.
   - Set alignment to `Centre`: confirm element centers within the full container width.
   - Set alignment to `Right`: confirm element aligns to the right edge.
2. **Search Bar Alignment**:
   - In Property Pane Page 1, change Search Alignment from `Left` to `Centre` and `Right`.
   - Confirm search input and taxonomy filter bar re-align accordingly.

### Test Suite 3: Rich Text WYSIWYG & Focus Lifecycle (FEAT-006, FEAT-026, FEAT-034)
1. **Toolbar Appearance**:
   - Click into any container title, subtitle, or card description.
   - Verify floating ribbon appears anchored directly above the text box.
2. **Focus Dismissal**:
   - Click onto the dashboard canvas, an empty section area, or an adjacent card.
   - Confirm the floating ribbon disappears immediately without ghosting.
3. **Sub-Surface Interaction**:
   - Highlight text and click the Font Family dropdown or Color Picker popover in the ribbon.
   - Confirm selecting a font or color modifies the text without dismissing the toolbar prematurely.

### Test Suite 4: Taxonomy Filtering & Process Models (FEAT-008, FEAT-030, FEAT-031)
1. **Global Term Store Dropdowns**:
   - In Edit Mode, click `+ Add Term Filter` to configure a taxonomy dropdown (e.g., "Our Sectors").
   - Select a term (e.g., "Infrastructure"): confirm only cards tagged with "Infrastructure" are displayed.
   - Confirm active summary pill reads `✓ updated — showing results for Infrastructure`.
   - Click the dismiss `(x)` on the summary pill: confirm all cards are restored.
2. **Process Model Chevrons**:
   - Click on "Stage 1: Shape": confirm card filter event dispatches and matches cards.

### Test Suite 5: Cloud Backup, Restore & Resilience (FEAT-019, FEAT-022, FEAT-027)
1. **Document Library Provisioning**:
   - Click `💾 Save snapshot to Library`.
   - Confirm prompt appears if the `Dashboards` library does not exist, followed by clean upload.
2. **Direct REST Fallback**:
   - Confirm snapshot saves into `Dashboards/Backups/<Dashboard-Title>/` with timestamped filename.

### Test Suite 6: Search Bar — Relocation & Placeholder Editing (FEAT-035)
1. **Always-Visible Search & Height Parity**:
   - In both Read Mode and Edit Mode, confirm the search input is rendered inside `headerTopRight`, immediately to the left of the layout mode switcher.
   - Confirm the search bar has increased height (32px medium size with vertical padding), expanded width (220px to 320px), and matches the exact vertical height of the adjacent Tabs / Accordion switcher.
   - Confirm typing a query filters visible cards across both Tabs and Accordion views.
2. **Placeholder Edit Button**:
   - Enter Edit Mode; confirm a pencil (`EditRegular`) icon appears alongside the search input.
   - Click the pencil icon; confirm a Fluent UI 2 `Popover` opens containing a labelled `Input` field pre-filled with the current placeholder.
   - Change the placeholder text and press Enter; confirm the `Input` placeholder updates and the Popover closes.
   - Press Escape; confirm the Popover closes and the placeholder reverts to the previous value.
   - Exit to Read Mode; confirm the pencil icon is no longer visible.

### Test Suite 7: Dropdown Filter Composable Item (FEAT-036)
1. **Insertion via Toolbox**:
   - In Edit Mode, open the toolbox (+) and click **Filter Dropdown**.
   - Confirm a dropdown item is inserted at the correct position with default label "Filter by", placeholder "Select an option", and two starter options.
2. **Property Editing**:
   - Open the property editor for the dropdown item.
   - Update the Dropdown label, Placeholder text, and Term Set name fields; confirm values persist.
   - Add a third option via **Add Option**; confirm it appears in the list.
   - Use the Up/Down arrows to reorder options; confirm the list order updates correctly.
   - Delete an option; confirm the option is removed.
3. **Filter Dispatch**:
   - In view mode, select an option from the dropdown; confirm a `dashboard:card-filter-apply` event is dispatched with the correct `filterValue`.
   - Verify cards whose tags / filter values match the selection remain visible; others are hidden.
   - The card containing the dropdown itself must never be hidden (filter-host protection).
### Test Suite 8: User Profile Diagnostics Inspection (FEAT-037)
1. **Edit Banner User Profile Icon**:
   - In Edit Mode, verify the `PersonRegular` icon button is visible directly adjacent to `💾 Save snapshot to Library`.
   - In Read Mode, verify neither the button nor edit banner are rendered.
2. **Diagnostics Popover Display**:
   - Click the person icon button; confirm a Fluent UI 2 `Popover` opens cleanly beneath the button with elevation `tokens.shadow16`.
   - Verify properties extracted from `pageContext.user` are rendered with bold keys and grey values:
     - `displayName` (e.g. current user name)
     - `email` (e.g. user mail address)
     - `loginName` (e.g. user UPN or claims identifier)
     - `isAnonymousGuestUser` (false/true)
     - `isExternalGuestUser` (false/true)
   - Click outside the popover or on another canvas element; verify the popover dismisses cleanly without lingering.

### Test Suite 9: Process Model Property Overhaul & Composable Filter Pane (FEAT-038)
1. **Ergonomic Modal Surface & Pinned Footer**:
   - Click to edit any Process Model item.
   - Confirm the modal title displays `"Configure the process model"` instead of raw uppercase IDs.
   - Expand multiple stage accordions; verify that the dialog content scrolls smoothly as a single fluid surface without nested inner scrollbar traps or bottom clipping.
   - Verify that all stage fields ("Stage label", "Stage title", "Capabilities metric badge") have generous vertical spacing, distinct label separation, and caption margin preventing text crowding.
   - Verify the footer buttons (**Cancel** and **Apply Changes**) remain permanently pinned at the bottom with a clean border, eliminating button clipping/overflow.
2. **Rich Text Formatting for Stage Descriptions**:
   - Inside an expanded stage card, locate the description area with its micro-formatting toolbar.
   - Test **Bold** (`**text**`), **Italic** (`*text*`), and **Bullet list** (`• `) buttons on selected or unselected text.
   - Confirm formatted descriptions persist and render accurately across canvas process model stages.
3. **Interactivity & Action Segmented Switcher**:
   - In the stage editor, inspect the "INTERACTIVITY & ACTION" section.
   - Verify the segmented pill switcher clearly presents **[ 🔍 Filter Cards ]** and **[ 🔗 Open Web Link ]**.
   - When **Filter Cards** is selected, verify the guidance banner ("When clicked, filters the dashboard cards...") and the custom filter term input (with fallback to stage pill label).
   - When **Open Web Link** is selected, verify the target URL input with immediate navigation feedback.
4. **Composable Dropdown Properties Pane**:
   - Open properties for any composable Dropdown item.
   - Verify the dedicated property pane with Term Store Group dropdown, Term Set dropdown, unified Icon picker, custom labels, and reorderable fallback options.
5. **Interlocking Chevron Arrow Geometry & Flat End Caps**:
   - Inspect the Process Model component on the canvas and in the property editor mini preview.
   - Verify that Stage 1 has a flat vertical left end cap and an angled chevron arrow point on the right.
   - Verify intermediate stages (Stages 2, 3, 4) have a matching left chevron notch and a right arrow point, seamlessly interlocking without unwanted gaps.
   - Verify the final stage (Stage 5) has a left chevron notch and a flat vertical right end cap.
   - Verify that when an active stage is selected, its background turns deep Turner & Townsend navy (`#001436`), its indicator dot turns amber (`#EAA023`), and its text colors switch to crisp white and light blue.
   - Verify that on mobile/narrow viewports (<=820px), the component collapses into a clean vertical stacked process list without clipped polygons.

---

## 3. Defensive Engineering Invariants
- **Compilation Check**: `npx tsc --noEmit` must pass with code 0.
- **Event Cleanup**: Global pointer listeners in `RichTextEditable` must detach on unmount.
- **Context Isolation**: No direct `WebPartContext` injection into presentational child components.
- **Null Safety**: Optional chaining applied across all list items and term sets.
- **Dropdown Filter-Host Safety**: Cards containing a `'dropdown'` item must be exempt from filter-hiding logic in `TabsContainer.tsx` and `AccordionContainer.tsx` (guarded by `isFilterHost` check).
- **Dropdown Null Safety**: `item.dropdownOptions` must default to `[]` before `.map()` calls in all render paths (`ComposableContentSection`, `BlockRenderer`, `renderInnerItemPreview`).
- **Search Placeholder Persistence**: `searchPlaceholder` must survive React re-renders and be saved to web part properties via `onSearchPlaceholderChange` callback; it must not reset on layout mode switch.
- **User Profile Diagnostic Isolation**: `userProfileDetails` dictionary passes strictly read-only serializable string/boolean primitives extracted from `this.context.pageContext.user` to avoid circular references or context leakage.
- **Modal Dialog Flexbox Boundaries**: All property editors utilizing Fluent UI 2 `<DialogSurface>` must enforce `display: 'flex', flexDirection: 'column'` with scrolling contained exclusively in `<DialogContent>` (`overflowY: 'auto'`) and a rigid `flexShrink: 0` sticky footer to eliminate button clipping regardless of viewport height.
- **Composable Content Item Toolbar Elevation**: Item-level edit, move, and remove toolbars rendered in `ComposableContentSection` and `BlockRenderer` must be anchored with negative vertical offset (`top: -14px`, `right: 4px`) at `zIndex: 50` with solid opaque backgrounds and `tokens.shadow8` elevation to prevent occlusion by inner items such as clip-path process models or hero banners.


