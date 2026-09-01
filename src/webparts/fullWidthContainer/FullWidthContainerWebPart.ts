/**
 * @file FullWidthContainerWebPart.ts
 * @description SPFx entrypoint webpart for Full-Width Container solution.
 */

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import { FullWidthContainer } from './components/FullWidthContainer';
import { IFullWidthContainerProps } from './components/IFullWidthContainerProps';
import {
  DEFAULT_CONTAINER_SECTIONS,
  IContainerSection,
  LayoutMode,
  ContainerStyle
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
  customSectionsJson?: string;
}

export default class FullWidthContainerWebPart extends BaseClientSideWebPart<IFullWidthContainerWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _currentTheme: IReadonlyTheme | undefined;

  public render(): void {
    let activeSections: IContainerSection[] = DEFAULT_CONTAINER_SECTIONS;

    // Optional custom JSON override from Property Pane
    if (this.properties.customSectionsJson && this.properties.customSectionsJson.trim()) {
      try {
        const parsed = JSON.parse(this.properties.customSectionsJson);
        if (Array.isArray(parsed) && parsed.length > 0) {
          activeSections = parsed;
        }
      } catch (err) {
        // Fall back to default sections if custom JSON is invalid
      }
    }

    const element: React.ReactElement<IFullWidthContainerProps> = React.createElement(
      FullWidthContainer,
      {
        title: this.properties.title || 'Interactive Content Hub',
        subtitle: this.properties.subtitle || 'Centralised operational methodologies, metrics, and resources',
        layoutMode: this.properties.layoutMode || 'tabs',
        containerStyle: this.properties.containerStyle || 'glassmorphism',
        accentColor: this.properties.accentColor || '#0078d4',
        enableAnimation: this.properties.enableAnimation !== false,
        compactPadding: !!this.properties.compactPadding,
        showSearch: this.properties.showSearch !== false,
        sections: activeSections,
        isDarkTheme: this._isDarkTheme,
        userDisplayName: this.context.pageContext.user.displayName,
        spfxTheme: this._currentTheme
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._currentTheme = currentTheme;
    const theme = currentTheme as IReadonlyTheme & { isInverted?: boolean };
    this._isDarkTheme = !!theme.isInverted;
    const semanticColors = currentTheme.semanticColors;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

    this.render();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Configure your full-width container layout, styles, and content blocks.'
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
                  label: 'Subtitle / Description',
                  value: this.properties.subtitle || 'Centralised operational methodologies, metrics, and resources'
                }),
                PropertyPaneChoiceGroup('layoutMode', {
                  label: 'Default Layout Mode',
                  options: [
                    { key: 'tabs', text: 'Tabbed View', iconProps: { officeFabricIconFontName: 'Tab' } },
                    { key: 'accordion', text: 'Accordion View', iconProps: { officeFabricIconFontName: 'GroupList' } }
                  ]
                }),
                PropertyPaneDropdown('containerStyle', {
                  label: 'Visual Style',
                  options: [
                    { key: 'glassmorphism', text: 'Glassmorphism (Frosted Glass)' },
                    { key: 'gradient', text: 'Gradient Glow' },
                    { key: 'minimal', text: 'Minimalist Flat' }
                  ],
                  selectedKey: this.properties.containerStyle || 'glassmorphism'
                })
              ]
            },
            {
              groupName: 'Display & Interactivity',
              groupFields: [
                PropertyPaneToggle('showSearch', {
                  label: 'Enable Content Filter / Search Bar',
                  checked: this.properties.showSearch !== false
                }),
                PropertyPaneToggle('enableAnimation', {
                  label: 'Enable Micro-animations & Transitions',
                  checked: this.properties.enableAnimation !== false
                }),
                PropertyPaneToggle('compactPadding', {
                  label: 'Compact Vertical Padding',
                  checked: !!this.properties.compactPadding
                })
              ]
            },
            {
              groupName: 'Advanced Data Configuration',
              groupFields: [
                PropertyPaneTextField('customSectionsJson', {
                  label: 'Custom Sections & Cards JSON (Optional)',
                  multiline: true,
                  rows: 6,
                  placeholder: 'Paste custom JSON array of IContainerSection[] to override defaults'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
