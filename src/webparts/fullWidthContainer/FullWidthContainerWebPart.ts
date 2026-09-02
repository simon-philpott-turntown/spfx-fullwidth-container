/**
 * @file FullWidthContainerWebPart.ts
 * @description SPFx entrypoint webpart for Full-Width Container solution.
 * Provides a comprehensive, interactive 3-page Property Pane sidebar
 * allowing page authors to fully customize layout, sections, badges, British metrics (£), and cards in Edit Mode.
 */

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneChoiceGroup,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneHorizontalRule,
  type IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import { ErrorBoundary } from './components/ErrorBoundary';
import { FullWidthContainer } from './components/FullWidthContainer';
import { IFullWidthContainerProps } from './components/IFullWidthContainerProps';
import {
  DEFAULT_CONTAINER_SECTIONS,
  PRESET_TEMPLATES,
  IContainerSection,
  IContentBlock,
  LayoutMode,
  ContainerStyle,
  BlockType
} from './models/IContainerModels';

export interface IFullWidthContainerWebPartProps {
  title: string;
  subtitle: string;
  layoutMode: LayoutMode;
  containerStyle: ContainerStyle;
  accentColor: string;
  enableAnimation: boolean;
  compactPadding: boolean;
  showSearch: boolean;
  gridColumns?: number;
  gridRows?: number;
  cardHeightMode?: 'auto' | 'equal';
  presetTemplate?: string;
  sectionsJson: string;

  // Active section editor state in Property Pane
  activeSectionIndex: number;
  sectionTitle: string;
  sectionBadge: string;
  sectionIcon: string;
  sectionDescription: string;

  // Active card / metric editor state in Property Pane
  activeBlockSectionIndex: number;
  activeBlockIndex: number;
  blockType: BlockType;
  blockTitle: string;
  blockDescription: string;
  blockBadge: string;
  blockIcon: string;
  blockHeightMode: 'default' | 'auto' | 'equal';
  blockMetricValue: string;
  blockMetricTrend: string;
  blockActionText: string;
  blockActionUrl: string;
  blockTags: string;
  importJsonRaw?: string;
}

export default class FullWidthContainerWebPart extends BaseClientSideWebPart<IFullWidthContainerWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _currentTheme: IReadonlyTheme | undefined;

  /**
   * Retrieves active sections array from canonical properties or initializes default.
   */
  private _getActiveSections(): IContainerSection[] {
    if (this.properties && this.properties.sectionsJson && this.properties.sectionsJson.trim()) {
      try {
        const parsed = JSON.parse(this.properties.sectionsJson);
        if (Array.isArray(parsed) && parsed.length > 0) {
          return parsed;
        }
      } catch {
        // Fallback if JSON parsing fails
      }
    }
    // Initialize default
    const defaults = DEFAULT_CONTAINER_SECTIONS;
    if (this.properties) {
      this.properties.sectionsJson = JSON.stringify(defaults);
    }
    return defaults;
  }

  /**
   * Serializes updated sections array to webpart properties.
   */
  private _saveSections(sections: IContainerSection[]): void {
    if (this.properties) {
      this.properties.sectionsJson = JSON.stringify(sections);
    }
    this.render();
  }

  protected onInit(): Promise<void> {
    console.log('[FullWidthContainerWebPart] Initialising full-width container web part v1.0.0...');
    return super.onInit();
  }

  protected get disableReactivePropertyChanges(): boolean {
    return false;
  }

  public render(): void {
    if (!this.domElement) {
      return;
    }

    try {
      const props = this.properties || ({} as Partial<IFullWidthContainerWebPartProps>);
      const activeSections = this._getActiveSections();
      const userDisplayName = this.context?.pageContext?.user?.displayName || 'User';

      const containerElement: React.ReactElement<IFullWidthContainerProps> = React.createElement(
        FullWidthContainer,
        {
          title: props.title || 'Container title',
          subtitle: props.subtitle !== undefined
            ? props.subtitle
            : 'Container subtitle or description',
          layoutMode: props.layoutMode || 'tabs',
          containerStyle: props.containerStyle || 'standard',
          accentColor: props.accentColor || '#0078d4',
          enableAnimation: props.enableAnimation !== false,
          compactPadding: !!props.compactPadding,
          showSearch: props.showSearch !== false,
          gridColumns: props.gridColumns,
          gridRows: props.gridRows,
          cardHeightMode: props.cardHeightMode || 'auto',
          sections: activeSections,
          isDarkTheme: this._isDarkTheme,
          userDisplayName: userDisplayName,
          spfxTheme: this._currentTheme,
          isEditMode: this.displayMode === DisplayMode.Edit,
          onOpenPropertyPane: () => {
            if (this.context && this.context.propertyPane) {
              this.context.propertyPane.open();
            }
          },
          onTitleChange: (newTitle: string) => {
            this.properties.title = newTitle;
            this.render();
          },
          onSubtitleChange: (newSubtitle: string) => {
            this.properties.subtitle = newSubtitle;
            this.render();
          },
          onUpdateBlock: (sectionId: string, blockId: string, updatedFields: Partial<IContentBlock>) => {
            const sections = this._getActiveSections();
            const sec = sections.find((s) => s.id === sectionId);
            if (sec && Array.isArray(sec.blocks)) {
              const blk = sec.blocks.find((b) => b.id === blockId);
              if (blk) {
                Object.assign(blk, updatedFields);
                this._saveSections(sections);
              }
            }
          },
          onDeleteBlock: (sectionId: string, blockId: string) => {
            const sections = this._getActiveSections();
            const sec = sections.find((s) => s.id === sectionId);
            if (sec && Array.isArray(sec.blocks)) {
              sec.blocks = sec.blocks.filter((b) => b.id !== blockId);
              this._saveSections(sections);
            }
          },
          onAddBlock: (sectionId: string) => {
            const sections = this._getActiveSections();
            const sec = sections.find((s) => s.id === sectionId);
            if (sec) {
              if (!Array.isArray(sec.blocks)) {
                sec.blocks = [];
              }
              sec.blocks.push({
                id: `blk-${Date.now()}`,
                type: 'card',
                title: 'Card title',
                description: 'Card summary',
                badge: 'Badge',
                iconName: 'BookAnswers',
                linkText: 'Action link',
                linkUrl: '#'
              });
              this._saveSections(sections);
            }
          },
          onEditBlockProperties: (sectionIndex: number, blockIndex: number) => {
            const sections = this._getActiveSections();
            this.properties.activeBlockSectionIndex = sectionIndex;
            this.properties.activeBlockIndex = blockIndex;
            this._syncBlockFields(sections, sectionIndex, blockIndex);
            if (this.context && this.context.propertyPane) {
              this.context.propertyPane.open();
            }
          },
          onAddSection: () => {
            this._handleAddSection();
          },
          onUpdateSection: (sectionId: string, updatedFields: Partial<IContainerSection>) => {
            const sections = this._getActiveSections();
            const sec = sections.find((s) => s.id === sectionId);
            if (sec) {
              Object.assign(sec, updatedFields);
              this._saveSections(sections);
            }
          }
        }
      );

      const rootElement = React.createElement(
        ErrorBoundary,
        { fallbackTitle: 'Full-width container' },
        containerElement
      );

      ReactDom.render(rootElement, this.domElement);
    } catch (err) {
      console.error('[FullWidthContainerWebPart] Critical Render Error:', err);
      const errorDetails = err instanceof Error ? (err.stack || err.message) : JSON.stringify(err);
      this.domElement.innerHTML = `
        <div style="padding: 24px; color: #a80000; background: #fde7e9; border: 2px solid #d13438; border-radius: 8px; font-family: Segoe UI, sans-serif;">
          <h3 style="margin-top:0; font-size: 18px;">⚠️ Web part render error (detailed diagnostics)</h3>
          <p style="font-size: 14px; margin-bottom: 12px;">An error occurred during web part execution:</p>
          <pre style="white-space: pre-wrap; word-break: break-all; background: #ffffff; padding: 12px; border: 1px solid #d13438; border-radius: 4px; font-size: 12px; color: #333333;">${errorDetails}</pre>
        </div>
      `;
    }
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._currentTheme = currentTheme;
    const theme = currentTheme as IReadonlyTheme & { isInverted?: boolean };
    this._isDarkTheme = !!theme.isInverted;

    // Only re-render if web part is already mounted and properties are initialized
    if (this.domElement && this.properties) {
      this.render();
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /**
   * Responds to field updates in the property pane sidebar in real-time.
   */
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    const sections = this._getActiveSections();

    // 1. Preset Template selection
    if (propertyPath === 'presetTemplate' && typeof newValue === 'string' && PRESET_TEMPLATES[newValue]) {
      const template = PRESET_TEMPLATES[newValue];
      this._saveSections(template.sections);
      this.properties.activeSectionIndex = 0;
      this.properties.activeBlockSectionIndex = 0;
      this.properties.activeBlockIndex = 0;
      this._syncSectionFields(template.sections, 0);
      this.context.propertyPane.refresh();
      return;
    }

    // 2. Section selector changed
    if (propertyPath === 'activeSectionIndex') {
      const idx = Number(newValue) || 0;
      this._syncSectionFields(sections, idx);
      this.context.propertyPane.refresh();
      return;
    }

    // 3. Section fields edited
    if (propertyPath.startsWith('section')) {
      const secIdx = this.properties.activeSectionIndex || 0;
      if (sections[secIdx]) {
        if (propertyPath === 'sectionTitle') sections[secIdx].title = String(newValue || '');
        if (propertyPath === 'sectionBadge') sections[secIdx].badge = String(newValue || '');
        if (propertyPath === 'sectionIcon') sections[secIdx].iconName = String(newValue || '');
        if (propertyPath === 'sectionDescription') sections[secIdx].description = String(newValue || '');
        this._saveSections(sections);
      }
      return;
    }

    // 4. Block target section changed
    if (propertyPath === 'activeBlockSectionIndex') {
      this.properties.activeBlockIndex = 0;
      const secIdx = Number(newValue) || 0;
      this._syncBlockFields(sections, secIdx, 0);
      this.context.propertyPane.refresh();
      return;
    }

    // 5. Block selector changed
    if (propertyPath === 'activeBlockIndex') {
      const secIdx = this.properties.activeBlockSectionIndex || 0;
      const blkIdx = Number(newValue) || 0;
      this._syncBlockFields(sections, secIdx, blkIdx);
      this.context.propertyPane.refresh();
      return;
    }

    // 6. Block fields edited
    if (propertyPath.startsWith('block')) {
      const secIdx = this.properties.activeBlockSectionIndex || 0;
      const blkIdx = this.properties.activeBlockIndex || 0;
      if (sections[secIdx] && sections[secIdx].blocks[blkIdx]) {
        const block = sections[secIdx].blocks[blkIdx];
        if (propertyPath === 'blockType') block.type = newValue as BlockType;
        if (propertyPath === 'blockTitle') block.title = String(newValue || '');
        if (propertyPath === 'blockDescription') block.description = String(newValue || '');
        if (propertyPath === 'blockBadge') block.badge = String(newValue || '');
        if (propertyPath === 'blockIcon') block.iconName = String(newValue || '');
        if (propertyPath === 'blockHeightMode') block.heightMode = newValue as 'default' | 'auto' | 'equal';
        if (propertyPath === 'blockMetricValue') block.metricValue = String(newValue || '');
        if (propertyPath === 'blockMetricTrend') block.metricTrend = String(newValue || '');
        if (propertyPath === 'blockActionText') block.linkText = String(newValue || '');
        if (propertyPath === 'blockActionUrl') block.linkUrl = String(newValue || '');
        if (propertyPath === 'blockTags') {
          block.tags = String(newValue || '').split(',').map(t => t.trim()).filter(Boolean);
        }
        this._saveSections(sections);
      }
    }
  }

  private _syncSectionFields(sections: IContainerSection[], index: number): void {
    const sec = sections[index] || sections[0];
    if (sec) {
      this.properties.sectionTitle = sec.title || '';
      this.properties.sectionBadge = sec.badge || '';
      this.properties.sectionIcon = sec.iconName || 'BookAnswers';
      this.properties.sectionDescription = sec.description || '';
    }
  }

  private _syncBlockFields(sections: IContainerSection[], secIndex: number, blockIndex: number): void {
    const sec = sections[secIndex] || sections[0];
    if (sec && sec.blocks && sec.blocks[blockIndex]) {
      const blk = sec.blocks[blockIndex];
      this.properties.blockType = blk.type || 'card';
      this.properties.blockTitle = blk.title || '';
      this.properties.blockDescription = blk.description || '';
      this.properties.blockBadge = blk.badge || '';
      this.properties.blockIcon = blk.iconName || 'BookAnswers';
      this.properties.blockHeightMode = blk.heightMode || 'default';
      this.properties.blockMetricValue = blk.metricValue || '';
      this.properties.blockMetricTrend = blk.metricTrend || '';
      this.properties.blockActionText = blk.linkText || '';
      this.properties.blockActionUrl = blk.linkUrl || '';
      this.properties.blockTags = blk.tags ? blk.tags.join(', ') : '';
    }
  }

  /**
   * Action handler: Adds a new section.
   */
  private _handleAddSection(): void {
    const sections = this._getActiveSections();
    const newIdx = sections.length + 1;
    const newSection: IContainerSection = {
      id: `sec-${Date.now()}`,
      title: `Section title ${newIdx}`,
      iconName: 'BookAnswers',
      badge: 'Badge',
      description: `Section summary ${newIdx}`,
      blocks: [
        {
          id: `blk-${Date.now()}`,
          type: 'card',
          title: 'Card title',
          description: 'Card summary',
          badge: 'Badge',
          iconName: 'BookAnswers',
          linkText: 'Action link',
          linkUrl: '#',
          tags: ['Tag 1']
        }
      ]
    };
    sections.push(newSection);
    const newIndex = sections.length - 1;
    this.properties.activeSectionIndex = newIndex;
    this._syncSectionFields(sections, newIndex);
    this._saveSections(sections);
    this.context.propertyPane.refresh();
  }

  /**
   * Action handler: Deletes the active section.
   */
  private _handleDeleteSection(): void {
    const sections = this._getActiveSections();
    if (sections.length <= 1) {
      return; // Keep at least one section
    }
    const idx = this.properties.activeSectionIndex || 0;
    sections.splice(idx, 1);
    const nextIdx = Math.max(0, idx - 1);
    this.properties.activeSectionIndex = nextIdx;
    this._syncSectionFields(sections, nextIdx);
    this._saveSections(sections);
    this.context.propertyPane.refresh();
  }

  /**
   * Action handler: Adds a new card/metric block to current section.
   */
  private _handleAddBlock(): void {
    const sections = this._getActiveSections();
    const secIdx = this.properties.activeBlockSectionIndex || 0;
    if (!sections[secIdx]) return;

    const newBlock: IContentBlock = {
      id: `blk-${Date.now()}`,
      type: 'card',
      title: 'Card title',
      description: 'Card summary',
      badge: 'Badge',
      iconName: 'BookAnswers',
      linkText: 'Action link',
      linkUrl: '#',
      tags: ['Tag 1']
    };

    sections[secIdx].blocks.push(newBlock);
    const newBlkIdx = sections[secIdx].blocks.length - 1;
    this.properties.activeBlockIndex = newBlkIdx;
    this._syncBlockFields(sections, secIdx, newBlkIdx);
    this._saveSections(sections);
    this.context.propertyPane.refresh();
  }

  /**
   * Action handler: Deletes the active card/metric block.
   */
  private _handleDeleteBlock(): void {
    const sections = this._getActiveSections();
    const secIdx = this.properties.activeBlockSectionIndex || 0;
    if (!sections[secIdx] || sections[secIdx].blocks.length <= 1) return;

    const blkIdx = this.properties.activeBlockIndex || 0;
    sections[secIdx].blocks.splice(blkIdx, 1);
    const nextBlkIdx = Math.max(0, blkIdx - 1);
    this.properties.activeBlockIndex = nextBlkIdx;
    this._syncBlockFields(sections, secIdx, nextBlkIdx);
    this._saveSections(sections);
    this.context.propertyPane.refresh();
  }

  /**
   * Action handler: Imports and validates Dashboard JSON template.
   */
  private _applyImportedJson(): void {
    if (!this.properties.importJsonRaw || this.properties.importJsonRaw.trim() === '') return;
    try {
      const parsed = JSON.parse(this.properties.importJsonRaw);
      if (parsed.sectionsJson) {
        this.properties.sectionsJson = parsed.sectionsJson;
      } else if (Array.isArray(parsed)) {
        this.properties.sectionsJson = JSON.stringify(parsed);
      } else if (parsed.sections && Array.isArray(parsed.sections)) {
        this.properties.sectionsJson = JSON.stringify(parsed.sections);
      }
      if (parsed.title) this.properties.title = parsed.title;
      if (parsed.subtitle) this.properties.subtitle = parsed.subtitle;
      if (parsed.layoutMode) this.properties.layoutMode = parsed.layoutMode;
      if (parsed.containerStyle) this.properties.containerStyle = parsed.containerStyle;
      if (parsed.gridColumns !== undefined) this.properties.gridColumns = parsed.gridColumns;
      if (parsed.gridRows !== undefined) this.properties.gridRows = parsed.gridRows;
      if (parsed.cardHeightMode) this.properties.cardHeightMode = parsed.cardHeightMode;

      this.render();
      this.context.propertyPane.refresh();
    } catch {
      // Invalid JSON syntax handled safely
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const sections = this._getActiveSections();
    const activeSecIdx = Math.min(this.properties.activeSectionIndex || 0, Math.max(0, sections.length - 1));
    const activeBlockSecIdx = Math.min(this.properties.activeBlockSectionIndex || 0, Math.max(0, sections.length - 1));
    const currentBlocks = sections[activeBlockSecIdx]?.blocks || [];
    const activeBlkIdx = Math.min(this.properties.activeBlockIndex || 0, Math.max(0, currentBlocks.length - 1));

    // Section dropdown options
    const sectionOptions: IPropertyPaneDropdownOption[] = sections.map((sec, i) => ({
      key: i,
      text: `${i + 1}. ${sec.title} (${sec.badge || 'Section'})`
    }));

    // Block dropdown options
    const blockOptions: IPropertyPaneDropdownOption[] = currentBlocks.map((blk, i) => ({
      key: i,
      text: `${i + 1}. ${blk.title} [${blk.type}]`
    }));

    const iconDropdownOptions: IPropertyPaneDropdownOption[] = [
      { key: 'BookAnswers', text: 'Book or documentation' },
      { key: 'Financial', text: 'Financial or pound (£)' },
      { key: 'AppIconDefault', text: 'Apps or tool grid' },
      { key: 'Shield', text: 'Shield or governance' },
      { key: 'ComplianceAudit', text: 'Audit or checklist' },
      { key: 'CheckList', text: 'Checklist or verified' },
      { key: 'TimelineProgress', text: 'Timeline or trending metrics' },
      { key: 'Calculator', text: 'Calculator or engineering' },
      { key: 'Lock', text: 'Lock or security' },
      { key: 'Globe', text: 'Globe or portal' },
      { key: 'DocumentManagement', text: 'Document or report' }
    ];

    return {
      pages: [
        // PAGE 1: Dashboard & Layout Options
        {
          header: {
            description: 'Page 1 of 3: Configure dashboard presentation, layout modes, and preset templates.'
          },
          groups: [
            {
              groupName: 'Dashboard presentation',
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Dashboard title',
                  value: this.properties.title || 'Dashboard title'
                }),
                PropertyPaneTextField('subtitle', {
                  label: 'Subtitle or guidance',
                  value: this.properties.subtitle || 'Dashboard subtitle or description'
                }),
                PropertyPaneChoiceGroup('layoutMode', {
                  label: 'Default layout mode',
                  options: [
                    { key: 'tabs', text: 'Tabbed view', iconProps: { officeFabricIconFontName: 'Tab' } },
                    { key: 'accordion', text: 'Accordion view', iconProps: { officeFabricIconFontName: 'GroupList' } }
                  ]
                }),
                PropertyPaneDropdown('containerStyle', {
                  label: 'Dashboard surface style',
                  selectedKey: this.properties.containerStyle || 'standard',
                  options: [
                    { key: 'standard', text: 'Standard Fluent 2 (solid layered surface)' },
                    { key: 'glassmorphism', text: 'Acrylic glassmorphism (frosted translucent)' },
                    { key: 'branded', text: 'Branded tint (inherited site palette)' },
                    { key: 'minimal', text: 'Minimal (clean transparent)' }
                  ]
                })
              ]
            },
            {
              groupName: 'Controls and interactivity',
              groupFields: [
                PropertyPaneToggle('showSearch', {
                  label: 'Enable real-time search and filter',
                  checked: this.properties.showSearch !== false
                }),
                PropertyPaneToggle('compactPadding', {
                  label: 'Compact vertical padding',
                  checked: !!this.properties.compactPadding
                }),
                PropertyPaneDropdown('gridColumns', {
                  label: 'Grid columns per row',
                  selectedKey: this.properties.gridColumns || 0,
                  options: [
                    { key: 0, text: 'Auto-fit responsive (dynamic)' },
                    { key: 1, text: '1 column (stacked full-width)' },
                    { key: 2, text: '2 columns (50% / 50%)' },
                    { key: 3, text: '3 columns (33% / 33% / 33%)' },
                    { key: 4, text: '4 columns (25% / 25% / 25% / 25%)' },
                    { key: 5, text: '5 columns' },
                    { key: 6, text: '6 columns' }
                  ]
                }),
                PropertyPaneDropdown('gridRows', {
                  label: 'Grid rows per container',
                  selectedKey: this.properties.gridRows || 0,
                  options: [
                    { key: 0, text: 'Auto or dynamic (unlimited)' },
                    { key: 1, text: '1 row (single row)' },
                    { key: 2, text: '2 rows' },
                    { key: 3, text: '3 rows' },
                    { key: 4, text: '4 rows' },
                    { key: 5, text: '5 rows' },
                    { key: 6, text: '6 rows' }
                  ]
                }),
                PropertyPaneDropdown('cardHeightMode', {
                  label: 'Card height alignment',
                  selectedKey: this.properties.cardHeightMode || 'auto',
                  options: [
                    { key: 'auto', text: 'Fit content height (independent)' },
                    { key: 'equal', text: 'Equal row height (match tallest)' }
                  ]
                })
              ]
            },
            {
              groupName: 'Preset templates',
              groupFields: [
                PropertyPaneDropdown('presetTemplate', {
                  label: 'Load ready-to-use template',
                  options: [
                    { key: 'commercial', text: 'Commercial and financial hub (GBP £)' },
                    { key: 'governance', text: 'Assurance and governance portal' },
                    { key: 'starter', text: 'Blank 2-section starter' }
                  ],
                  selectedKey: this.properties.presetTemplate || 'commercial'
                })
              ]
            },
            {
              groupName: 'Configuration JSON & Portability',
              groupFields: [
                PropertyPaneTextField('importJsonRaw', {
                  label: 'Dashboard JSON Template',
                  multiline: true,
                  rows: 3,
                  placeholder: 'Paste full dashboard JSON to import or view current structure...',
                  value: this.properties.importJsonRaw || ''
                }),
                PropertyPaneButton('applyImportJsonBtn', {
                  text: '📥 Import & Apply JSON',
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: () => this._applyImportedJson()
                })
              ]
            }
          ]
        },

        // PAGE 2: Manage Sections (Tabs / Accordions)
        {
          header: {
            description: 'Page 2 of 3: Add, edit, and rebrand content sections, badges, and Fluent 2 icons.'
          },
          groups: [
            {
              groupName: 'Select section to edit',
              groupFields: [
                PropertyPaneDropdown('activeSectionIndex', {
                  label: 'Active section',
                  options: sectionOptions,
                  selectedKey: activeSecIdx
                }),
                PropertyPaneHorizontalRule()
              ]
            },
            {
              groupName: 'Section details and metadata',
              groupFields: [
                PropertyPaneTextField('sectionTitle', {
                  label: 'Section title',
                  value: sections[activeSecIdx]?.title || ''
                }),
                PropertyPaneDropdown('sectionBadge', {
                  label: 'Section badge',
                  options: [
                    { key: 'Badge', text: 'Badge' },
                    { key: 'Core', text: 'Core' },
                    { key: 'Real-time', text: 'Real-time' },
                    { key: 'Interactive', text: 'Interactive' },
                    { key: 'Compliance', text: 'Compliance' },
                    { key: 'Finance', text: 'Finance' },
                    { key: 'Verified', text: 'Verified' }
                  ],
                  selectedKey: sections[activeSecIdx]?.badge || 'Badge'
                }),
                PropertyPaneDropdown('sectionIcon', {
                  label: 'Section icon',
                  options: iconDropdownOptions,
                  selectedKey: sections[activeSecIdx]?.iconName || 'BookAnswers'
                }),
                PropertyPaneTextField('sectionDescription', {
                  label: 'Section guidance or summary',
                  multiline: true,
                  rows: 2,
                  value: sections[activeSecIdx]?.description || ''
                })
              ]
            },
            {
              groupName: 'Section actions',
              groupFields: [
                PropertyPaneButton('btnAddSection', {
                  text: '+ Add new section',
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this._handleAddSection.bind(this)
                }),
                PropertyPaneButton('btnDeleteSection', {
                  text: 'Delete current section',
                  buttonType: PropertyPaneButtonType.Normal,
                  disabled: sections.length <= 1,
                  onClick: this._handleDeleteSection.bind(this)
                })
              ]
            }
          ]
        },

        // PAGE 3: Manage Content Cards & Metrics (£)
        {
          header: {
            description: 'Page 3 of 3: Add and configure individual cards, British financial metrics (£), and action links.'
          },
          groups: [
            {
              groupName: 'Target section and card selection',
              groupFields: [
                PropertyPaneDropdown('activeBlockSectionIndex', {
                  label: 'Target section',
                  options: sectionOptions,
                  selectedKey: activeBlockSecIdx
                }),
                PropertyPaneDropdown('activeBlockIndex', {
                  label: 'Select card or metric to edit',
                  options: blockOptions.length > 0 ? blockOptions : [{ key: 0, text: '(No items in section)' }],
                  selectedKey: activeBlkIdx,
                  disabled: currentBlocks.length === 0
                }),
                PropertyPaneHorizontalRule()
              ]
            },
            {
              groupName: 'Card or metric properties',
              groupFields: [
                PropertyPaneDropdown('blockType', {
                  label: 'Item type',
                  options: [
                    { key: 'card', text: 'Standard content card' },
                    { key: 'metric', text: 'British metric stat (£)' },
                    { key: 'embed', text: 'Embedded tool or portal' }
                  ],
                  selectedKey: currentBlocks[activeBlkIdx]?.type || 'card'
                }),
                PropertyPaneTextField('blockTitle', {
                  label: 'Item title',
                  value: currentBlocks[activeBlkIdx]?.title || ''
                }),
                PropertyPaneDropdown('blockIcon', {
                  label: 'Item icon',
                  options: iconDropdownOptions,
                  selectedKey: currentBlocks[activeBlkIdx]?.iconName || 'BookAnswers'
                }),
                PropertyPaneDropdown('blockHeightMode', {
                  label: 'Card height behavior',
                  options: [
                    { key: 'default', text: 'Inherit container setting' },
                    { key: 'auto', text: 'Fit content height (independent)' },
                    { key: 'equal', text: 'Equal row height (match tallest)' }
                  ],
                  selectedKey: currentBlocks[activeBlkIdx]?.heightMode || 'default'
                }),
                PropertyPaneTextField('blockDescription', {
                  label: 'Card summary or description',
                  multiline: true,
                  rows: 2,
                  value: currentBlocks[activeBlkIdx]?.description || ''
                }),
                PropertyPaneTextField('blockBadge', {
                  label: 'Badge label',
                  value: currentBlocks[activeBlkIdx]?.badge || ''
                }),
                PropertyPaneTextField('blockMetricValue', {
                  label: 'Metric value in GBP (£)',
                  value: currentBlocks[activeBlkIdx]?.metricValue || ''
                }),
                PropertyPaneTextField('blockMetricTrend', {
                  label: 'Metric trend (e.g. +10% vs target)',
                  value: currentBlocks[activeBlkIdx]?.metricTrend || ''
                }),
                PropertyPaneTextField('blockActionText', {
                  label: 'Action button label',
                  value: currentBlocks[activeBlkIdx]?.linkText || ''
                }),
                PropertyPaneTextField('blockActionUrl', {
                  label: 'Action destination URL',
                  value: currentBlocks[activeBlkIdx]?.linkUrl || ''
                }),
                PropertyPaneTextField('blockTags', {
                  label: 'Tags (comma-separated for search filtering)',
                  value: currentBlocks[activeBlkIdx]?.tags ? currentBlocks[activeBlkIdx]?.tags.join(', ') : ''
                })
              ]
            },
            {
              groupName: 'Card actions',
              groupFields: [
                PropertyPaneButton('btnAddBlock', {
                  text: '+ Add card to section',
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this._handleAddBlock.bind(this)
                }),
                PropertyPaneButton('btnDeleteBlock', {
                  text: 'Delete selected card',
                  buttonType: PropertyPaneButtonType.Normal,
                  disabled: currentBlocks.length <= 1,
                  onClick: this._handleDeleteBlock.bind(this)
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
