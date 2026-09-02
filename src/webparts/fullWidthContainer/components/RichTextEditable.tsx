/**
 * @file RichTextEditable.tsx
 * @description Inline WYSIWYG rich text editor with non-collapsing Floating Toolbar.
 * Inherits SharePoint site theme colors, brand fonts, and typography tokens.
 */

import * as React from 'react';
import { makeStyles, shorthands, tokens } from '@fluentui/react-components';
import { FloatingTextToolbar } from './FloatingTextToolbar';

const useStyles = makeStyles({
  wrapper: {
    position: 'relative',
    width: '100%'
  },
  toolbarWrapper: {
    position: 'absolute',
    bottom: 'calc(100% + 4px)',
    left: 0,
    marginBottom: '2px',
    zIndex: 1000,
    width: 'max-content',
    minWidth: 'max-content',
    maxWidth: 'none',
    whiteSpace: 'nowrap'
  },
  editable: {
    outlineStyle: 'none',
    width: '100%',
    boxSizing: 'border-box',
    transitionProperty: 'box-shadow, border-color, background-color',
    transitionDuration: '150ms',
    ':empty::before': {
      content: 'attr(data-placeholder)',
      color: tokens.colorNeutralForeground4,
      pointerEvents: 'none'
    }
  },
  editableActive: {
    ...shorthands.borderRadius(tokens.borderRadiusSmall),
    ...shorthands.border('1px', 'dashed', tokens.colorBrandStroke1),
    backgroundColor: tokens.colorNeutralBackground1Hover,
    ...shorthands.padding('2px', '4px')
  }
});

export interface IRichTextEditableProps {
  html: string;
  isEditMode: boolean;
  onChange: (newHtml: string) => void;
  placeholder?: string;
  className?: string;
  style?: React.CSSProperties;
  tag?: 'h1' | 'h2' | 'h3' | 'p' | 'div' | 'span';
}

export const RichTextEditable: React.FC<IRichTextEditableProps> = ({
  html,
  isEditMode,
  onChange,
  placeholder = 'Type text...',
  className = '',
  style = {},
  tag = 'div'
}) => {
  const styles = useStyles();
  const elementRef = React.useRef<HTMLDivElement | null>(null);
  const [isFocused, setIsFocused] = React.useState<boolean>(false);
  const [savedRange, setSavedRange] = React.useState<Range | null>(null);

  // Sync incoming HTML to DOM when not actively typing
  React.useEffect(() => {
    if (elementRef.current && !isFocused) {
      if (elementRef.current.innerHTML !== (html || '')) {
        elementRef.current.innerHTML = html || '';
      }
    }
  }, [html, isFocused]);

  const handleSelectionSave = (): void => {
    if (typeof window !== 'undefined') {
      const sel = window.getSelection();
      if (sel && sel.rangeCount > 0 && elementRef.current) {
        const range = sel.getRangeAt(0);
        if (elementRef.current.contains(range.commonAncestorContainer)) {
          setSavedRange(range.cloneRange());
        }
      }
    }
  };

  const handleInput = (): void => {
    if (elementRef.current) {
      const newHtml = elementRef.current.innerHTML;
      onChange(newHtml);
    }
  };

  const handleFormat = (command: string, value?: string): void => {
    if (elementRef.current) {
      elementRef.current.focus();
      if (savedRange && typeof window !== 'undefined') {
        const sel = window.getSelection();
        if (sel) {
          sel.removeAllRanges();
          sel.addRange(savedRange);
        }
      }
      document.execCommand(command, false, value);
      handleInput();
      handleSelectionSave();
    }
  };

  const handleFocus = (): void => {
    setIsFocused(true);
    handleSelectionSave();
  };

  const handleBlur = (e: React.FocusEvent): void => {
    // If blur was caused by clicking inside the floating toolbar, do not close
    if (e.relatedTarget && (e.relatedTarget as HTMLElement).closest('.floating-toolbar-container')) {
      return;
    }
    setTimeout(() => {
      setIsFocused(false);
      if (elementRef.current) {
        onChange(elementRef.current.innerHTML);
      }
    }, 200);
  };

  const commonStyle: React.CSSProperties = {
    fontFamily: tokens.fontFamilyBase,
    color: tokens.colorNeutralForeground1,
    ...style
  };

  if (!isEditMode) {
    return React.createElement(tag, {
      className,
      style: commonStyle,
      dangerouslySetInnerHTML: { __html: html || placeholder }
    });
  }

  return (
    <div className={styles.wrapper}>
      {isFocused && (
        <div className={styles.toolbarWrapper} onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); }}>
          <FloatingTextToolbar onFormat={handleFormat} />
        </div>
      )}
      {React.createElement(tag, {
        ref: elementRef,
        contentEditable: true,
        suppressContentEditableWarning: true,
        className: `${styles.editable} ${isFocused ? styles.editableActive : ''} ${className}`,
        'data-placeholder': placeholder,
        style: commonStyle,
        onFocus: handleFocus,
        onBlur: handleBlur,
        onInput: handleInput,
        onKeyUp: handleSelectionSave,
        onMouseUp: handleSelectionSave
      })}
    </div>
  );
};
