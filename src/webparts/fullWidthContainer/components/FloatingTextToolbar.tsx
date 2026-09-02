/**
 * @file FloatingTextToolbar.tsx
 * @description Floating rich text formatting toolbar modeled after the native SharePoint Text Web Part ribbon.
 * Features Text Style (H1, H2, H3, H4, Body), Font Family selector, comprehensive SharePoint Site Theme Palette,
 * highlight backgrounds, list formatting, and text alignment.
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
  CodeRegular,
  TextFontRegular,
  TextHeader1Regular,
  ChevronDownRegular
} from '@fluentui/react-icons';

const useStyles = makeStyles({
  toolbarContainer: {
    display: 'inline-flex',
    alignItems: 'center',
    ...shorthands.padding('3px', '6px'),
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke1),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    boxShadow: tokens.shadow28,
    ...shorthands.gap('2px'),
    zIndex: 1000,
    marginBottom: tokens.spacingVerticalXS,
    flexWrap: 'nowrap',
    whiteSpace: 'nowrap',
    width: 'max-content',
    minWidth: 'max-content',
    maxWidth: 'none',
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
  dropdownBtn: {
    height: '26px',
    fontSize: '0.8rem',
    fontWeight: tokens.fontWeightSemibold,
    padding: '0 6px',
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px'
  },
  menuSurface: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
    padding: '6px',
    minWidth: '150px',
    maxHeight: '260px',
    overflowY: 'auto'
  },
  menuItem: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: '6px 10px',
    borderRadius: tokens.borderRadiusSmall,
    cursor: 'pointer',
    fontSize: '0.85rem',
    transitionProperty: 'background-color, color',
    transitionDuration: '100ms',
    ':hover': {
      backgroundColor: tokens.colorBrandBackground2,
      color: tokens.colorBrandForeground1
    }
  },
  colorGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(6, 1fr)',
    ...shorthands.gap('6px'),
    ...shorthands.padding('8px')
  },
  colorSectionHeader: {
    fontSize: '0.72rem',
    fontWeight: 600,
    color: tokens.colorNeutralForeground3,
    padding: '4px 8px 2px 8px',
    textTransform: 'uppercase',
    letterSpacing: '0.5px'
  },
  colorSwatch: {
    width: '22px',
    height: '22px',
    ...shorthands.borderRadius('4px'),
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    cursor: 'pointer',
    transitionProperty: 'transform, border-color',
    transitionDuration: '100ms',
    ':hover': {
      transform: 'scale(1.18)',
      ...shorthands.borderColor(tokens.colorBrandStroke1)
    }
  }
});

export interface IFloatingTextToolbarProps {
  onFormat?: (command: string, value?: string) => void;
  isBold?: boolean;
  isItalic?: boolean;
  isUnderline?: boolean;
  isStrikethrough?: boolean;
}

/**
 * Full SharePoint Online Site Theme & Semantic Palette
 */
export const SITE_THEME_COLORS = [
  // Primary Site Theme Tints
  { name: 'Theme Darker', hex: '#004578', group: 'theme' },
  { name: 'Theme Dark', hex: '#005a9e', group: 'theme' },
  { name: 'Brand Primary', hex: '#0078d4', group: 'theme' },
  { name: 'Theme Light', hex: '#2b88d8', group: 'theme' },
  { name: 'Theme Lighter', hex: '#c7e0f4', group: 'theme' },
  { name: 'Theme Tint', hex: '#eff6fc', group: 'theme' },

  // Neutral & Content Palette
  { name: 'Black / Dark 1', hex: '#242424', group: 'neutral' },
  { name: 'Neutral Dark', hex: '#323130', group: 'neutral' },
  { name: 'Neutral Primary', hex: '#424242', group: 'neutral' },
  { name: 'Neutral Secondary', hex: '#605e5c', group: 'neutral' },
  { name: 'Neutral Tertiary', hex: '#a19f9d', group: 'neutral' },
  { name: 'Neutral Light / White', hex: '#ffffff', group: 'neutral' },

  // SharePoint Status & Accent Colors
  { name: 'Success Green', hex: '#107c10', group: 'accents' },
  { name: 'Warning Gold', hex: '#d83b01', group: 'accents' },
  { name: 'Error Crimson', hex: '#a80000', group: 'accents' },
  { name: 'Royal Blue', hex: '#004e8c', group: 'accents' },
  { name: 'Deep Purple', hex: '#5c2d91', group: 'accents' },
  { name: 'Teal Accent', hex: '#038387', group: 'accents' }
];

export const HIGHLIGHT_COLORS = [
  { name: 'Bright Yellow', hex: '#fff100' },
  { name: 'Pastel Green', hex: '#92c353' },
  { name: 'Sky Cyan', hex: '#00bcf2' },
  { name: 'Soft Lavender', hex: '#b4a0ff' },
  { name: 'Warm Peach', hex: '#ffb900' },
  { name: 'Light Rose', hex: '#e3008c' }
];

export const FONT_FAMILIES = [
  { name: 'Segoe UI (Standard)', value: 'Segoe UI, -apple-system, BlinkMacSystemFont, Roboto, sans-serif' },
  { name: 'Aptos (Modern 365)', value: 'Aptos, Segoe UI, sans-serif' },
  { name: 'Arial', value: 'Arial, Helvetica, sans-serif' },
  { name: 'Calibri', value: 'Calibri, Candara, Segoe, sans-serif' },
  { name: 'Georgia', value: 'Georgia, Cambria, serif' },
  { name: 'Tahoma', value: 'Tahoma, Verdana, Segoe, sans-serif' },
  { name: 'Trebuchet MS', value: 'Trebuchet MS, Lucida Sans Unicode, sans-serif' },
  { name: 'Times New Roman', value: 'Times New Roman, Times, serif' },
  { name: 'Courier New (Code)', value: 'Courier New, Courier, monospace' }
];

export const TEXT_STYLES = [
  { label: 'Header 1', tag: '<h1>', style: { fontSize: '1.75rem', fontWeight: 700 } },
  { label: 'Header 2', tag: '<h2>', style: { fontSize: '1.35rem', fontWeight: 600 } },
  { label: 'Header 3', tag: '<h3>', style: { fontSize: '1.15rem', fontWeight: 600 } },
  { label: 'Header 4', tag: '<h4>', style: { fontSize: '1rem', fontWeight: 600 } },
  { label: 'Body text', tag: '<p>', style: { fontSize: '0.9rem', fontWeight: 400 } },
  { label: 'Quote', tag: '<blockquote>', style: { fontStyle: 'italic', borderLeft: '3px solid #0078d4', paddingLeft: '8px' } },
  { label: 'Code block', tag: '<pre>', style: { fontFamily: 'monospace', fontSize: '0.85rem' } }
];

export const FloatingTextToolbar: React.FC<IFloatingTextToolbarProps> = ({
  onFormat,
  isBold = false,
  isItalic = false,
  isUnderline = false,
  isStrikethrough = false
}) => {
  const styles = useStyles();
  const [selectedStyleLabel, setSelectedStyleLabel] = React.useState<string>('Style');
  const [selectedFontLabel, setSelectedFontLabel] = React.useState<string>('Font');

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
      <div className={styles.dragHandle} title="Formatting Ribbon">
        <LineHorizontal1Regular fontSize={14} />
      </div>

      {/* Text Style / Type Dropdown (H1, H2, H3, H4, Body) */}
      <Popover positioning="below">
        <PopoverTrigger disableButtonEnhancement>
          <Button
            size="small"
            appearance="subtle"
            className={styles.dropdownBtn}
            icon={<TextHeader1Regular />}
            title="Text Style (Heading 1, 2, 3, Body)"
            onMouseDown={handleToolbarMouseDown}
          >
            {selectedStyleLabel}
            <ChevronDownRegular fontSize={10} />
          </Button>
        </PopoverTrigger>
        <PopoverSurface onMouseDown={handleToolbarMouseDown} className={styles.menuSurface}>
          <Caption1 style={{ fontWeight: 600, padding: '2px 8px', color: tokens.colorNeutralForeground3 }}>
            Text Styles
          </Caption1>
          {TEXT_STYLES.map((st) => (
            <div
              key={st.label}
              className={styles.menuItem}
              onMouseDown={handleToolbarMouseDown}
              onClick={() => {
                handleCommand('formatBlock', st.tag);
                setSelectedStyleLabel(st.label);
              }}
            >
              <span style={st.style}>{st.label}</span>
            </div>
          ))}
        </PopoverSurface>
      </Popover>

      {/* Font Family Selector Dropdown */}
      <Popover positioning="below">
        <PopoverTrigger disableButtonEnhancement>
          <Button
            size="small"
            appearance="subtle"
            className={styles.dropdownBtn}
            icon={<TextFontRegular />}
            title="Font Family"
            onMouseDown={handleToolbarMouseDown}
          >
            {selectedFontLabel}
            <ChevronDownRegular fontSize={10} />
          </Button>
        </PopoverTrigger>
        <PopoverSurface onMouseDown={handleToolbarMouseDown} className={styles.menuSurface}>
          <Caption1 style={{ fontWeight: 600, padding: '2px 8px', color: tokens.colorNeutralForeground3 }}>
            Font Families
          </Caption1>
          {FONT_FAMILIES.map((font) => (
            <div
              key={font.name}
              className={styles.menuItem}
              onMouseDown={handleToolbarMouseDown}
              onClick={() => {
                handleCommand('fontName', font.value);
                setSelectedFontLabel(font.name.split(' ')[0]);
              }}
            >
              <span style={{ fontFamily: font.value }}>{font.name}</span>
            </div>
          ))}
        </PopoverSurface>
      </Popover>

      <Divider vertical style={{ height: '18px', margin: '0 2px' }} />

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
        title="Inline code snippet"
        onMouseDown={handleToolbarMouseDown}
        onClick={() => handleCommand('formatBlock', '<pre>')}
      />

      <Divider vertical style={{ height: '18px', margin: '0 2px' }} />

      {/* Full SharePoint Theme Color Picker Popover */}
      <Popover positioning="below">
        <PopoverTrigger disableButtonEnhancement>
          <Button
            size="small"
            appearance="subtle"
            icon={<TextColorRegular />}
            title="SharePoint Site Theme Colour Palette"
            onMouseDown={handleToolbarMouseDown}
          />
        </PopoverTrigger>
        <PopoverSurface onMouseDown={handleToolbarMouseDown} style={{ padding: '6px', minWidth: '200px' }}>
          <div className={styles.colorSectionHeader}>Site Theme Colours</div>
          <div className={styles.colorGrid}>
            {SITE_THEME_COLORS.filter((c) => c.group === 'theme').map((c) => (
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

          <div className={styles.colorSectionHeader}>Neutral & Text Shades</div>
          <div className={styles.colorGrid}>
            {SITE_THEME_COLORS.filter((c) => c.group === 'neutral').map((c) => (
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

          <div className={styles.colorSectionHeader}>Standard Accents</div>
          <div className={styles.colorGrid}>
            {SITE_THEME_COLORS.filter((c) => c.group === 'accents').map((c) => (
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

      {/* Highlight Background Color Picker Popover */}
      <Popover positioning="below">
        <PopoverTrigger disableButtonEnhancement>
          <Button
            size="small"
            appearance="subtle"
            icon={<HighlightRegular />}
            title="Highlight Background Colour"
            onMouseDown={handleToolbarMouseDown}
          />
        </PopoverTrigger>
        <PopoverSurface onMouseDown={handleToolbarMouseDown} style={{ padding: '6px', minWidth: '180px' }}>
          <div className={styles.colorSectionHeader}>Highlight Backgrounds</div>
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

      <Divider vertical style={{ height: '18px', margin: '0 2px' }} />

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

      <Divider vertical style={{ height: '18px', margin: '0 2px' }} />

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
