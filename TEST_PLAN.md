# SPFx Full-Width Container Web Part — Comprehensive Test Plan

## Executive Summary
This document establishes the end-to-end test plan and verification matrix for all 34 registered features of the **SPFx Full-Width Container Web Part** (`spfx-fullwidth-container`), spanning layout engines, canvas editing, taxonomy integrations, backup pipelines, and defensive engineering standards.

---

## 1. Feature Verification Matrix (All 34 Features)

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
| **FEAT-030** | GlobalTermStoreDropdownFilters | [TermFilterBar.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/TermFilterBar.tsx) | Filter cards via taxonomy dropdowns & summary pill | PASS |
| **FEAT-031** | FilterButtonsAndProcessModel | [FilterButtonsRenderer.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/FilterButtonsRenderer.tsx) | Test pill buttons & chevron stages filtering | PASS |
| **FEAT-032** | TransparentCardLayoutContainer | [BlockRenderer.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/BlockRenderer.tsx) | Render borderless/backgroundless card slot | PASS |
| **FEAT-033** | ComposableHeaderAndSectionContent | [ComposableContentSection.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/ComposableContentSection.tsx) | Top-of-webpart & top-of-section content slots | PASS |
| **FEAT-034** | TextToolbarFocusDismissalAndPreservation | [RichTextEditable.tsx](file:///d:/Playbook/spfx-fullwidth-container/src/webparts/fullWidthContainer/components/RichTextEditable.tsx) | Ensure formatting ribbon disappears on blur | PASS |

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

---

## 3. Defensive Engineering Invariants
- **Compilation Check**: `npx tsc --noEmit` must pass with code 0.
- **Event Cleanup**: Global pointer listeners in `RichTextEditable` must detach on unmount.
- **Context Isolation**: No direct `WebPartContext` injection into presentational child components.
- **Null Safety**: Optional chaining applied across all list items and term sets.
