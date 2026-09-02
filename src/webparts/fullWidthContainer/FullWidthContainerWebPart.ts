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
  presetTemplate: string;
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
  blockMetricValue: string;
  blockMetricTrend: string;
  blockActionText: string;
  blockActionUrl: string;
  blockTags: string;
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
    console.log('[FullWidthContainerWebPart] Initializing Full-Width Container Web Part v1.0.0...');
    return super.onInit();
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
          title: props.title || 'Interactive Content Hub',
          subtitle: props.subtitle !== undefined
            ? props.subtitle
            : 'Centralised operational methodologies, metrics, and resources',
          layoutMode: props.layoutMode || 'tabs',
          containerStyle: props.containerStyle || 'glassmorphism',
          accentColor: props.accentColor || '#0078d4',
          enableAnimation: props.enableAnimation !== false,
          compactPadding: !!props.compactPadding,
          showSearch: props.showSearch !== false,
          sections: activeSections,
          isDarkTheme: this._isDarkTheme,
          userDisplayName: userDisplayName,
          spfxTheme: this._currentTheme,
          isEditMode: this.displayMode === DisplayMode.Edit,
          onOpenPropertyPane: () => {
            if (this.context && this.context.propertyPane) {
              this.context.propertyPane.open();
            }
          }
        }
      );

      const rootElement = React.createElement(
        ErrorBoundary,
        { fallbackTitle: 'Full-Width Container' },
        containerElement
      );

      ReactDom.render(rootElement, this.domElement);
    } catch (err) {
      console.error('[FullWidthContainerWebPart] Critical Render Error:', err);
      const errorDetails = err instanceof Error ? (err.stack || err.message) : JSON.stringify(err);
      this.domElement.innerHTML = `
        <div style="padding: 24px; color: #a80000; background: #fde7e9; border: 2px solid #d13438; border-radius: 8px; font-family: Segoe UI, sans-serif;">
          <h3 style="margin-top:0; font-size: 18px;">⚠️ Web Part Render Error (Detailed Diagnostics)</h3>
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
      this.properties.sectionIcon = sec.iconName || 'Document';
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
      this.properties.blockMetricValue = blk.metricValue || '';
      this.properties.blockMetricTrend = blk.metricTrend || '';
      this.properties.blockActionText = blk.linkText || '';
      this.properties.blockActionUrl = blk.linkUrl || '';
      this.properties.blockTags = blk.tags ? blk.tags.join(', ') : '';
    }
  }

  /**
   * Action handler: Adds a new Section.
   */
  private _handleAddSection(): void {
    const sections = this._getActiveSections();
    const newSection: IContainerSection = {
      id: `sec-${Date.now()}`,
      title: `New Section ${sections.length + 1}`,
      iconName: 'Document',
      badge: 'Core',
      description: 'Section description and instructions',
      blocks: [
        {
          id: `blk-${Date.now()}`,
          type: 'card',
          title: 'Sample Card Item',
          description: 'Customise this item using Page 3 in the properties sidebar.',
          badge: 'New',
          linkText: 'Learn More',
          linkUrl: '#'
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
   * Action handler: Deletes the active Section.
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
   * Action handler: Adds a new Card/Metric block to current section.
   */
  private _handleAddBlock(): void {
    const sections = this._getActiveSections();
    const secIdx = this.properties.activeBlockSectionIndex || 0;
    if (!sections[secIdx]) return;

    const newBlock: IContentBlock = {
      id: `blk-${Date.now()}`,
      type: 'card',
      title: `New Card ${sections[secIdx].blocks.length + 1}`,
      description: 'Enter your card description here in the properties sidebar.',
      badge: 'Interactive',
      linkText: 'View Details',
      linkUrl: '#',
      tags: ['New']
    };

    sections[secIdx].blocks.push(newBlock);
    const newBlkIdx = sections[secIdx].blocks.length - 1;
    this.properties.activeBlockIndex = newBlkIdx;
    this._syncBlockFields(sections, secIdx, newBlkIdx);
    this._saveSections(sections);
    this.context.propertyPane.refresh();
  }

  /**
   * Action handler: Deletes the active Card/Metric block.
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

    return {
      pages: [
        // PAGE 1: Container & Layout Options
        {
          header: {
            description: 'Page 1 of 3: Configure Container presentation, layout modes, and preset templates.'
          },
          groups: [
            {
              groupName: 'Container Presentation',
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Container Title',
                  value: this.properties.title || 'Interactive Content Hub'
                }),
                PropertyPaneTextField('subtitle', {
                  label: 'Subtitle / Guidance',
                  value: this.properties.subtitle || 'Centralised operational methodologies, metrics, and resources'
                }),
                PropertyPaneChoiceGroup('layoutMode', {
                  label: 'Default Layout Mode',
                  options: [
                    { key: 'tabs', text: 'Tabbed View', iconProps: { officeFabricIconFontName: 'Tab' } },
                    { key: 'accordion', text: 'Accordion View', iconProps: { officeFabricIconFontName: 'GroupList' } }
                  ]
                })
              ]
            },
            {
              groupName: 'Controls & Interactivity',
              groupFields: [
                PropertyPaneToggle('showSearch', {
                  label: 'Enable Real-time Search / Filter',
                  checked: this.properties.showSearch !== false
                }),
                PropertyPaneToggle('compactPadding', {
                  label: 'Compact Vertical Padding',
                  checked: !!this.properties.compactPadding
                })
              ]
            },
            {
              groupName: 'Preset Templates',
              groupFields: [
                PropertyPaneDropdown('presetTemplate', {
                  label: 'Load Ready-to-Use Template',
                  options: [
                    { key: 'commercial', text: 'Commercial & Financial Hub (GBP £)' },
                    { key: 'governance', text: 'Assurance & Governance Portal' },
                    { key: 'starter', text: 'Blank 2-Section Starter' }
                  ],
                  selectedKey: this.properties.presetTemplate || 'commercial'
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
              groupName: 'Select Section to Edit',
              groupFields: [
                PropertyPaneDropdown('activeSectionIndex', {
                  label: 'Active Section',
                  options: sectionOptions,
                  selectedKey: activeSecIdx
                }),
                PropertyPaneHorizontalRule()
              ]
            },
            {
              groupName: 'Section Details & Metadata',
              groupFields: [
                PropertyPaneTextField('sectionTitle', {
                  label: 'Section Title',
                  value: sections[activeSecIdx]?.title || ''
                }),
                PropertyPaneDropdown('sectionBadge', {
                  label: 'Section Type / Badge',
                  options: [
                    { key: 'Core', text: 'Core (Standard)' },
                    { key: 'Real-time', text: 'Real-time (Live Feed)' },
                    { key: 'Interactive', text: 'Interactive (Tool / Portal)' },
                    { key: 'Compliance', text: 'Compliance / Mandatory' },
                    { key: 'Finance', text: 'Finance / Metrics (£)' },
                    { key: 'Verified', text: 'Verified Standards' }
                  ],
                  selectedKey: sections[activeSecIdx]?.badge || 'Core'
                }),
                PropertyPaneDropdown('sectionIcon', {
                  label: 'Section Fluent 2 Icon',
                  options: [
                    { key: 'BookAnswers', text: 'Book / Playbook' },
                    { key: 'Financial', text: 'Financial / Pound (£)' },
                    { key: 'AppIconDefault', text: 'App / Live Portal' },
                    { key: 'Shield', text: 'Shield / Security' },
                    { key: 'ComplianceAudit', text: 'Audit / Checklist' },
                    { key: 'DocumentManagement', text: 'Document Management' },
                    { key: 'TimelineProgress', text: 'Timeline / Progress' }
                  ],
                  selectedKey: sections[activeSecIdx]?.iconName || 'BookAnswers'
                }),
                PropertyPaneTextField('sectionDescription', {
                  label: 'Section Guidance / Subtitle',
                  multiline: true,
                  rows: 2,
                  value: sections[activeSecIdx]?.description || ''
                })
              ]
            },
            {
              groupName: 'Section Actions',
              groupFields: [
                PropertyPaneButton('btnAddSection', {
                  text: '+ Add New Section',
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this._handleAddSection.bind(this)
                }),
                PropertyPaneButton('btnDeleteSection', {
                  text: 'Delete Current Section',
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
              groupName: 'Target Section & Card Selection',
              groupFields: [
                PropertyPaneDropdown('activeBlockSectionIndex', {
                  label: 'Target Section',
                  options: sectionOptions,
                  selectedKey: activeBlockSecIdx
                }),
                PropertyPaneDropdown('activeBlockIndex', {
                  label: 'Select Card / Metric to Edit',
                  options: blockOptions.length > 0 ? blockOptions : [{ key: 0, text: '(No items in section)' }],
                  selectedKey: activeBlkIdx,
                  disabled: currentBlocks.length === 0
                }),
                PropertyPaneHorizontalRule()
              ]
            },
            {
              groupName: 'Card / Metric Properties',
              groupFields: [
                PropertyPaneDropdown('blockType', {
                  label: 'Item Type',
                  options: [
                    { key: 'card', text: 'Standard Card with Link' },
                    { key: 'metric', text: 'Financial Metric Stat (£)' },
                    { key: 'embed', text: 'Embedded Tool / Portal' }
                  ],
                  selectedKey: currentBlocks[activeBlkIdx]?.type || 'card'
                }),
                PropertyPaneTextField('blockTitle', {
                  label: 'Item Title',
                  value: currentBlocks[activeBlkIdx]?.title || ''
                }),
                PropertyPaneTextField('blockDescription', {
                  label: 'Description / Summary',
                  multiline: true,
                  rows: 2,
                  value: currentBlocks[activeBlkIdx]?.description || ''
                }),
                PropertyPaneTextField('blockBadge', {
                  label: 'Badge Label (e.g. P1 Mandatory, Live, Compliance)',
                  value: currentBlocks[activeBlkIdx]?.badge || ''
                }),
                PropertyPaneTextField('blockMetricValue', {
                  label: 'Metric Value (For Metric Type, e.g. £1,420,000, 94.8%)',
                  value: currentBlocks[activeBlkIdx]?.metricValue || ''
                }),
                PropertyPaneTextField('blockMetricTrend', {
                  label: 'Metric Trend (e.g. +18.4% ahead of target)',
                  value: currentBlocks[activeBlkIdx]?.metricTrend || ''
                }),
                PropertyPaneTextField('blockActionText', {
                  label: 'Action Button Label (e.g. Open Playbook, Launch)',
                  value: currentBlocks[activeBlkIdx]?.linkText || ''
                }),
                PropertyPaneTextField('blockActionUrl', {
                  label: 'Action Destination URL (e.g. https://...)',
                  value: currentBlocks[activeBlkIdx]?.linkUrl || ''
                }),
                PropertyPaneTextField('blockTags', {
                  label: 'Tags (comma-separated for search filtering)',
                  value: currentBlocks[activeBlkIdx]?.tags ? currentBlocks[activeBlkIdx]?.tags.join(', ') : ''
                })
              ]
            },
            {
              groupName: 'Card Actions',
              groupFields: [
                PropertyPaneButton('btnAddBlock', {
                  text: '+ Add Card to Section',
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this._handleAddBlock.bind(this)
                }),
                PropertyPaneButton('btnDeleteBlock', {
                  text: 'Delete Selected Card',
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
