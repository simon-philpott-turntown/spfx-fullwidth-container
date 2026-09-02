/**
 * @file FloatingTextToolbar.tsx
 * @description Floating rich formatting toolbar modeled after the native SharePoint Text Web Part ribbon.
 * Follows the Microsoft Fluent 2 Design System.
 */

import * as React from 'react';
import {
  Button,
  makeStyles,
  shorthands,
  tokens,
  Divider,
  Dropdown,
  Option
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
  TextNumberListRegular,
  LineHorizontal1Regular
} from '@fluentui/react-icons';

const useStyles = makeStyles({
  toolbarContainer: {
    display: 'inline-flex',
    alignItems: 'center',
    ...shorthands.padding('3px', '6px'),
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke1),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    boxShadow: tokens.shadow8,
    ...shorthands.gap('2px'),
    zIndex: 100,
    marginBottom: tokens.spacingVerticalXS,
    flexWrap: 'wrap'
  },
  dragHandle: {
    cursor: 'grab',
    display: 'flex',
    alignItems: 'center',
    color: tokens.colorNeutralForeground3,
    paddingRight: '4px'
  },
  styleDropdown: {
    minWidth: '95px',
    height: '28px'
  },
  sizeDropdown: {
    minWidth: '60px',
    height: '28px'
  },
  btnActive: {
    backgroundColor: tokens.colorNeutralBackground1Selected,
    color: tokens.colorBrandForeground1
  }
});

export interface IFloatingTextToolbarProps {
  onFormat?: (command: string, value?: string) => void;
  currentStyle?: string;
  currentSize?: string;
  isBold?: boolean;
  isItalic?: boolean;
  isUnderline?: boolean;
}

export const FloatingTextToolbar: React.FC<IFloatingTextToolbarProps> = ({
  onFormat,
  currentStyle = 'Normal',
  currentSize = '16',
  isBold = false,
  isItalic = false,
  isUnderline = false
}) => {
  const styles = useStyles();

  const handleCommand = (cmd: string, val?: string): void => {
    if (onFormat) {
      onFormat(cmd, val);
    }
  };

  return (
    <div className={styles.toolbarContainer} onClick={(e) => e.stopPropagation()}>
      {/* Drag handle */}
      <div className={styles.dragHandle} title="Formatting Tools">
        <LineHorizontal1Regular fontSize={14} />
      </div>

      {/* Style Dropdown */}
      <Dropdown
        className={styles.styleDropdown}
        size="small"
        value={currentStyle}
        onOptionSelect={(e, data) => handleCommand('formatBlock', data.optionValue)}
      >
        <Option value="Normal">Normal</Option>
        <Option value="Heading 1">Heading 1</Option>
        <Option value="Heading 2">Heading 2</Option>
        <Option value="Heading 3">Heading 3</Option>
        <Option value="Pull quote">Pull quote</Option>
      </Dropdown>

      {/* Size Dropdown */}
      <Dropdown
        className={styles.sizeDropdown}
        size="small"
        value={currentSize}
        onOptionSelect={(e, data) => handleCommand('fontSize', data.optionValue)}
      >
        <Option value="24">24</Option>
        <Option value="20">20</Option>
        <Option value="18">18</Option>
        <Option value="16">16</Option>
        <Option value="14">14</Option>
        <Option value="12">12</Option>
      </Dropdown>

      <Divider vertical style={{ height: '20px', margin: '0 4px' }} />

      {/* Bold, Italic, Underline */}
      <Button
        size="small"
        appearance="subtle"
        className={isBold ? styles.btnActive : ''}
        icon={<TextBoldRegular />}
        title="Bold (Ctrl+B)"
        onClick={() => handleCommand('bold')}
      />
      <Button
        size="small"
        appearance="subtle"
        className={isItalic ? styles.btnActive : ''}
        icon={<TextItalicRegular />}
        title="Italic (Ctrl+I)"
        onClick={() => handleCommand('italic')}
      />
      <Button
        size="small"
        appearance="subtle"
        className={isUnderline ? styles.btnActive : ''}
        icon={<TextUnderlineRegular />}
        title="Underline (Ctrl+U)"
        onClick={() => handleCommand('underline')}
      />

      {/* Color & Highlight */}
      <Button
        size="small"
        appearance="subtle"
        icon={<TextColorRegular />}
        title="Font Colour"
        onClick={() => handleCommand('foreColor')}
      />
      <Button
        size="small"
        appearance="subtle"
        icon={<HighlightRegular />}
        title="Highlight"
        onClick={() => handleCommand('hiliteColor')}
      />
      <Button
        size="small"
        appearance="subtle"
        icon={<LinkRegular />}
        title="Insert Link"
        onClick={() => handleCommand('createLink')}
      />
      <Button
        size="small"
        appearance="subtle"
        icon={<TextClearFormattingRegular />}
        title="Clear Formatting"
        onClick={() => handleCommand('removeFormat')}
      />

      <Divider vertical style={{ height: '20px', margin: '0 4px' }} />

      {/* Alignment */}
      <Button
        size="small"
        appearance="subtle"
        icon={<TextAlignLeftRegular />}
        title="Align Left"
        onClick={() => handleCommand('justifyLeft')}
      />
      <Button
        size="small"
        appearance="subtle"
        icon={<TextAlignCenterRegular />}
        title="Align Centre"
        onClick={() => handleCommand('justifyCenter')}
      />
      <Button
        size="small"
        appearance="subtle"
        icon={<TextAlignRightRegular />}
        title="Align Right"
        onClick={() => handleCommand('justifyRight')}
      />

      <Divider vertical style={{ height: '20px', margin: '0 4px' }} />

      {/* Lists */}
      <Button
        size="small"
        appearance="subtle"
        icon={<TextBulletListRegular />}
        title="Bullet List"
        onClick={() => handleCommand('insertUnorderedList')}
      />
      <Button
        size="small"
        appearance="subtle"
        icon={<TextNumberListRegular />}
        title="Numbered List"
        onClick={() => handleCommand('insertOrderedList')}
      />
    </div>
  );
};
