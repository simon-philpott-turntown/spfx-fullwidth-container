/**
 * @file FloatingTextToolbar.tsx
 * @description Floating rich text formatting toolbar modeled after the native SharePoint Text Web Part ribbon.
 * Follows the Microsoft Fluent 2 Design System and inherits site theme palette and typography.
 */

import * as React from 'react';
import {
  Button,
  makeStyles,
  shorthands,
  tokens,
  Divider,
  Popover,
  PopoverSurface,
  PopoverTrigger,
  Caption1
} from '@fluentui/react-components';
import {
  TextBoldRegular,
  TextItalicRegular,
  TextUnderlineRegular,
  TextColorRegular,
  HighlightRegular,
  LinkRegular,
  TextClearFormattingRegular,
  TextAlignLeftRegular,
  TextAlignCenterRegular,
  TextAlignRightRegular,
  TextBulletListRegular,
  TextNumberListLtrRegular,
  LineHorizontal1Regular,
  TextStrikethroughRegular,
  CodeRegular
} from '@fluentui/react-icons';

const useStyles = makeStyles({
  toolbarContainer: {
    display: 'inline-flex',
    alignItems: 'center',
    ...shorthands.padding('3px', '6px'),
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke1),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    boxShadow: tokens.shadow16,
    ...shorthands.gap('2px'),
    zIndex: 100,
    marginBottom: tokens.spacingVerticalXS,
    flexWrap: 'wrap',
    userSelect: 'none'
  },
  dragHandle: {
    cursor: 'grab',
    display: 'flex',
    alignItems: 'center',
    color: tokens.colorNeutralForeground3,
    paddingRight: '4px'
  },
  btnActive: {
    backgroundColor: tokens.colorNeutralBackground1Selected,
    color: tokens.colorBrandForeground1
  },
  colorGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(5, 1fr)',
    ...shorthands.gap('6px'),
    ...shorthands.padding('8px')
  },
  colorSwatch: {
    width: '24px',
    height: '24px',
    ...shorthands.borderRadius('4px'),
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    cursor: 'pointer',
    transitionProperty: 'transform, border-color',
    transitionDuration: '100ms',
    ':hover': {
      transform: 'scale(1.15)',
      ...shorthands.borderColor(tokens.colorBrandStroke1)
    }
  },
  fontSizeBtn: {
    minWidth: '32px',
    fontSize: '0.8rem',
    fontWeight: tokens.fontWeightSemibold,
    padding: '0 4px',
    height: '26px'
  }
});

export interface IFloatingTextToolbarProps {
  onFormat?: (command: string, value?: string) => void;
  isBold?: boolean;
  isItalic?: boolean;
  isUnderline?: boolean;
  isStrikethrough?: boolean;
}

export const SITE_THEME_COLORS = [
  { name: 'Brand Primary', hex: '#0078d4' },
  { name: 'Dark Theme', hex: '#106ebe' },
  { name: 'Light Brand', hex: '#2b88d8' },
  { name: 'Dark Neutral', hex: '#242424' },
  { name: 'Neutral Secondary', hex: '#616161' },
  { name: 'Success Green', hex: '#107c10' },
  { name: 'Warning Gold', hex: '#d83b01' },
  { name: 'Error Crimson', hex: '#a80000' },
  { name: 'Royal Blue', hex: '#004e8c' },
  { name: 'Deep Purple', hex: '#5c2d91' }
];

export const HIGHLIGHT_COLORS = [
  { name: 'Yellow Highlight', hex: '#fff176' },
  { name: 'Green Highlight', hex: '#a5d6a7' },
  { name: 'Cyan Highlight', hex: '#80deea' },
  { name: 'Lavender Highlight', hex: '#d1c4e9' },
  { name: 'Peach Highlight', hex: '#ffccbc' }
];

export const FloatingTextToolbar: React.FC<IFloatingTextToolbarProps> = ({
  onFormat,
  isBold = false,
  isItalic = false,
  isUnderline = false,
  isStrikethrough = false
}) => {
  const styles = useStyles();

  const handleCommand = (cmd: string, val?: string): void => {
    if (onFormat) {
      onFormat(cmd, val);
    } else {
      document.execCommand(cmd, false, val);
    }
  };

  /**
   * Prevents mouse down default action so text selection inside contentEditable is NOT blurred or cleared.
   */
  const handleToolbarMouseDown = (e: React.MouseEvent): void => {
    e.preventDefault();
    e.stopPropagation();
  };

  return (
    <div
      className={styles.toolbarContainer}
      onMouseDown={handleToolbarMouseDown}
      onClick={(e) => e.stopPropagation()}
    >
      {/* Drag handle */}
      <div className={styles.dragHandle} title="Formatting Tools">
        <LineHorizontal1Regular fontSize={14} />
      </div>

      {/* Bold, Italic, Underline, Strikethrough */}
      <Button
        size="small"
        appearance="subtle"
        className={isBold ? styles.btnActive : ''}
        icon={<TextBoldRegular />}
        title="Bold (Ctrl+B)"
        onMouseDown={handleToolbarMouseDown}
        onClick={() => handleCommand('bold')}
      />
      <Button
        size="small"
        appearance="subtle"
        className={isItalic ? styles.btnActive : ''}
        icon={<TextItalicRegular />}
        title="Italic (Ctrl+I)"
        onMouseDown={handleToolbarMouseDown}
        onClick={() => handleCommand('italic')}
      />
      <Button
        size="small"
        appearance="subtle"
        className={isUnderline ? styles.btnActive : ''}
        icon={<TextUnderlineRegular />}
        title="Underline (Ctrl+U)"
        onMouseDown={handleToolbarMouseDown}
        onClick={() => handleCommand('underline')}
      />
      <Button
        size="small"
        appearance="subtle"
        className={isStrikethrough ? styles.btnActive : ''}
        icon={<TextStrikethroughRegular />}
        title="Strikethrough"
        onMouseDown={handleToolbarMouseDown}
        onClick={() => handleCommand('strikeThrough')}
      />
      <Button
        size="small"
        appearance="subtle"
        icon={<CodeRegular />}
        title="Code inline"
        onMouseDown={handleToolbarMouseDown}
        onClick={() => handleCommand('formatBlock', '<pre>')}
      />

      <Divider vertical style={{ height: '20px', margin: '0 4px' }} />

      {/* Font Theme Color Picker Popover */}
      <Popover positioning="below">
        <PopoverTrigger disableButtonEnhancement>
          <Button
            size="small"
            appearance="subtle"
            icon={<TextColorRegular />}
            title="Brand & Theme Font Colour"
            onMouseDown={handleToolbarMouseDown}
          />
        </PopoverTrigger>
        <PopoverSurface onMouseDown={handleToolbarMouseDown} style={{ padding: '8px' }}>
          <Caption1 style={{ fontWeight: 600, paddingLeft: '8px', display: 'block' }}>
            Site Theme Palette
          </Caption1>
          <div className={styles.colorGrid}>
            {SITE_THEME_COLORS.map((c) => (
              <div
                key={c.name}
                className={styles.colorSwatch}
                style={{ backgroundColor: c.hex }}
                title={c.name}
                onMouseDown={handleToolbarMouseDown}
                onClick={() => handleCommand('foreColor', c.hex)}
              />
            ))}
          </div>
        </PopoverSurface>
      </Popover>

      {/* Highlight Color Picker Popover */}
      <Popover positioning="below">
        <PopoverTrigger disableButtonEnhancement>
          <Button
            size="small"
            appearance="subtle"
            icon={<HighlightRegular />}
            title="Highlight Colour"
            onMouseDown={handleToolbarMouseDown}
          />
        </PopoverTrigger>
        <PopoverSurface onMouseDown={handleToolbarMouseDown} style={{ padding: '8px' }}>
          <Caption1 style={{ fontWeight: 600, paddingLeft: '8px', display: 'block' }}>
            Highlight Background
          </Caption1>
          <div className={styles.colorGrid}>
            {HIGHLIGHT_COLORS.map((c) => (
              <div
                key={c.name}
                className={styles.colorSwatch}
                style={{ backgroundColor: c.hex }}
                title={c.name}
                onMouseDown={handleToolbarMouseDown}
                onClick={() => handleCommand('hiliteColor', c.hex)}
              />
            ))}
          </div>
        </PopoverSurface>
      </Popover>

      {/* Insert Link */}
      <Button
        size="small"
        appearance="subtle"
        icon={<LinkRegular />}
        title="Insert Hyperlink"
        onMouseDown={handleToolbarMouseDown}
        onClick={() => {
          const url = window.prompt('Enter destination URL:');
          if (url) handleCommand('createLink', url);
        }}
      />

      {/* Clear Formatting */}
      <Button
        size="small"
        appearance="subtle"
        icon={<TextClearFormattingRegular />}
        title="Clear Formatting"
        onMouseDown={handleToolbarMouseDown}
        onClick={() => handleCommand('removeFormat')}
      />

      <Divider vertical style={{ height: '20px', margin: '0 4px' }} />

      {/* Alignment */}
      <Button
        size="small"
        appearance="subtle"
        icon={<TextAlignLeftRegular />}
        title="Align Left"
        onMouseDown={handleToolbarMouseDown}
        onClick={() => handleCommand('justifyLeft')}
      />
      <Button
        size="small"
        appearance="subtle"
        icon={<TextAlignCenterRegular />}
        title="Align Centre"
        onMouseDown={handleToolbarMouseDown}
        onClick={() => handleCommand('justifyCenter')}
      />
      <Button
        size="small"
        appearance="subtle"
        icon={<TextAlignRightRegular />}
        title="Align Right"
        onMouseDown={handleToolbarMouseDown}
        onClick={() => handleCommand('justifyRight')}
      />

      <Divider vertical style={{ height: '20px', margin: '0 4px' }} />

      {/* Lists */}
      <Button
        size="small"
        appearance="subtle"
        icon={<TextBulletListRegular />}
        title="Bullet List"
        onMouseDown={handleToolbarMouseDown}
        onClick={() => handleCommand('insertUnorderedList')}
      />
      <Button
        size="small"
        appearance="subtle"
        icon={<TextNumberListLtrRegular />}
        title="Numbered List"
        onMouseDown={handleToolbarMouseDown}
        onClick={() => handleCommand('insertOrderedList')}
      />
    </div>
  );
};
