/**
 * @file PropertyPaneIconField.tsx
 * @description Custom SPFx Property Pane field that displays an icon preview (24px) alongside its title,
 * and launches the complete 229-icon Fluent UI 2 & Turner & Townsend corporate SVG modal.
 */

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { IPropertyPaneField, PropertyPaneFieldType } from '@microsoft/sp-property-pane';
import { FluentProvider, webLightTheme, Button, Caption1, Subtitle2 } from '@fluentui/react-components';
import { renderUnifiedIcon } from './CustomSvgIconRegistry';
import { FluentIconPicker } from './FluentIconPicker';

export interface IPropertyPaneIconFieldProps {
  label: string;
  selectedIconKey: string;
  onSelectIcon: (iconKey: string) => void;
  buttonLabel?: string;
}

const IconFieldControl: React.FC<IPropertyPaneIconFieldProps> = ({
  label,
  selectedIconKey,
  onSelectIcon,
  buttonLabel = 'Change icon...'
}) => {
  const [isOpen, setIsOpen] = React.useState<boolean>(false);
  const [currentIcon, setCurrentIcon] = React.useState<string>(selectedIconKey || 'BookAnswers');

  React.useEffect(() => {
    setCurrentIcon(selectedIconKey || 'BookAnswers');
  }, [selectedIconKey]);

  const handleSelect = (iconKey: string): void => {
    setCurrentIcon(iconKey);
    onSelectIcon(iconKey);
    setIsOpen(false);
  };

  return (
    <div style={{ marginBottom: '12px' }}>
      <label
        style={{
          display: 'block',
          fontSize: '14px',
          fontWeight: 600,
          marginBottom: '6px',
          color: '#323130'
        }}
      >
        {label}
      </label>

      <div
        style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          padding: '8px 12px',
          backgroundColor: '#ffffff',
          border: '1px solid #d1d1d1',
          borderRadius: '4px',
          gap: '8px'
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', gap: '10px', overflow: 'hidden' }}>
          <div
            style={{
              width: '32px',
              height: '32px',
              borderRadius: '4px',
              backgroundColor: '#f3f2f1',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              fontSize: '20px',
              color: '#1e4479',
              flexShrink: 0
            }}
          >
            {renderUnifiedIcon(currentIcon, '#1e4479', '20px')}
          </div>
          <div style={{ overflow: 'hidden' }}>
            <Subtitle2
              style={{
                display: 'block',
                fontSize: '13px',
                fontWeight: 600,
                color: '#242424',
                whiteSpace: 'nowrap',
                overflow: 'hidden',
                textOverflow: 'ellipsis'
              }}
            >
              {currentIcon}
            </Subtitle2>
            <Caption1 style={{ color: '#616161', fontSize: '11px' }}>
              {currentIcon.startsWith('svg-') ? 'SVG Icon' : 'Standard Icon'}
            </Caption1>
          </div>
        </div>

        <Button
          size="small"
          appearance="outline"
          onClick={() => setIsOpen(true)}
          style={{ flexShrink: 0 }}
        >
          {buttonLabel}
        </Button>
      </div>

      <FluentIconPicker
        isOpen={isOpen}
        selectedIconKey={currentIcon}
        onSelectIcon={handleSelect}
        onDismiss={() => setIsOpen(false)}
      />
    </div>
  );
};

/**
 * Factory creating the custom property pane field.
 */
export function PropertyPaneIconField(
  key: string,
  props: IPropertyPaneIconFieldProps
): IPropertyPaneField<any> {
  return {
    type: PropertyPaneFieldType.Custom,
    targetProperty: key,
    properties: {
      key,
      onRender: (domElement: HTMLElement) => {
        ReactDom.render(
          React.createElement(
            FluentProvider,
            { theme: webLightTheme },
            React.createElement(IconFieldControl, props)
          ),
          domElement
        );
      },
      onDispose: (domElement: HTMLElement) => {
        ReactDom.unmountComponentAtNode(domElement);
      }
    }
  };
}
