/**
 * @file RichTextEditable.tsx
 * @description Inline WYSIWYG rich text editor with non-collapsing Floating Toolbar.
 * Inherits SharePoint site theme colors, brand fonts, and typography tokens.
 * Only displays floating formatting toolbar when text element is actively focused or selected.
 */

import * as React from 'react';
import { makeStyles, shorthands, tokens, Popover, PopoverSurface } from '@fluentui/react-components';
import { FloatingTextToolbar } from './FloatingTextToolbar';
import { suppressSharePointWebPartDrag } from '../utils/dragIsolation';

const useStyles = makeStyles({
  wrapper: {
    position: 'relative',
    width: '100%'
  },
  toolbarSurface: {
    ...shorthands.padding('0px'),
    ...shorthands.border('none'),
    backgroundColor: 'transparent',
    boxShadow: 'none',
    overflow: 'visible',
    zIndex: 1000001
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
  const popoverSurfaceRef = React.useRef<HTMLDivElement | null>(null);
  const blurTimeoutRef = React.useRef<number | null>(null);
  const isInteractingWithToolbarRef = React.useRef<boolean>(false);
  const [isFocused, setIsFocused] = React.useState<boolean>(false);
  const [savedRange, setSavedRange] = React.useState<Range | null>(null);

  // Clear pending blur timeouts on unmount
  React.useEffect(() => {
    return () => {
      if (blurTimeoutRef.current) {
        window.clearTimeout(blurTimeoutRef.current);
      }
    };
  }, []);

  // Global document click listener: if user clicks outside this editable element and its floating toolbar, hide toolbar
  React.useEffect(() => {
    if (!isFocused) return;

    const handleDocumentPointerDown = (event: PointerEvent): void => {
      const target = event.target as Node | null;
      if (!target) return;

      // Check if click is inside editable content element
      if (elementRef.current && elementRef.current.contains(target)) {
        return;
      }

      // Check if click is inside the toolbar popover surface
      if (popoverSurfaceRef.current && popoverSurfaceRef.current.contains(target)) {
        return;
      }

      // Check if click is inside any child Popover dropdown (e.g. Font family, styles, color picker)
      if ((target as HTMLElement).closest && (target as HTMLElement).closest('.floating-toolbar-container')) {
        return;
      }

      // Click occurred completely outside: dismiss toolbar immediately
      isInteractingWithToolbarRef.current = false;
      if (blurTimeoutRef.current) {
        window.clearTimeout(blurTimeoutRef.current);
        blurTimeoutRef.current = null;
      }
      suppressSharePointWebPartDrag(false, elementRef.current || undefined);
      setIsFocused(false);
      if (elementRef.current) {
        onChange(elementRef.current.innerHTML);
      }
    };

    document.addEventListener('pointerdown', handleDocumentPointerDown, true);
    return () => {
      document.removeEventListener('pointerdown', handleDocumentPointerDown, true);
    };
  }, [isFocused, onChange]);

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

  const [blockOverrideStyle, setBlockOverrideStyle] = React.useState<React.CSSProperties>({});

  const handleFormat = (command: string, value?: string): void => {
    if (elementRef.current) {
      elementRef.current.focus();

      // Check if we have an active non-collapsed selection inside this element
      let hasTextSelection = false;
      if (typeof window !== 'undefined') {
        const sel = window.getSelection();
        if (savedRange && sel) {
          sel.removeAllRanges();
          sel.addRange(savedRange);
        }

        const currentSel = window.getSelection();
        if (
          currentSel &&
          !currentSel.isCollapsed &&
          currentSel.rangeCount > 0 &&
          elementRef.current.contains(currentSel.getRangeAt(0).commonAncestorContainer) &&
          currentSel.toString().length > 0
        ) {
          hasTextSelection = true;
        }
      }

      // Enable CSS style output where supported
      try {
        document.execCommand('styleWithCSS', false, 'true');
      } catch {
        // Fallback for older browsers
      }

      if (hasTextSelection) {
        // User highlighted specific text: format only that highlighted selection
        document.execCommand(command, false, value);
      } else {
        // No specific text selected: format the whole text block using DOM selectNodeContents so formatting is saved in HTML
        const fullRange = document.createRange();
        fullRange.selectNodeContents(elementRef.current);
        const sel = window.getSelection();
        if (sel) {
          sel.removeAllRanges();
          sel.addRange(fullRange);
        }

        if (command === 'fontName' && value) {
          document.execCommand('fontName', false, value);
          elementRef.current.style.fontFamily = value;
          setBlockOverrideStyle((prev) => ({ ...prev, fontFamily: value }));
        } else if (command === 'foreColor' && value) {
          document.execCommand('foreColor', false, value);
          elementRef.current.style.color = value;
          setBlockOverrideStyle((prev) => ({ ...prev, color: value }));
        } else if (command === 'hiliteColor' && value) {
          document.execCommand('hiliteColor', false, value);
          elementRef.current.style.backgroundColor = value;
          setBlockOverrideStyle((prev) => ({ ...prev, backgroundColor: value }));
        } else if (command === 'removeFormat') {
          elementRef.current.style.fontFamily = '';
          elementRef.current.style.color = '';
          elementRef.current.style.backgroundColor = '';
          setBlockOverrideStyle({});
          document.execCommand('removeFormat', false);
        } else {
          document.execCommand(command, false, value);
        }
      }

      handleInput();
      handleSelectionSave();
    }
  };

  const handleFocus = (): void => {
    if (blurTimeoutRef.current) {
      window.clearTimeout(blurTimeoutRef.current);
      blurTimeoutRef.current = null;
    }
    setIsFocused(true);
    handleSelectionSave();
    suppressSharePointWebPartDrag(true, elementRef.current || undefined);
  };

  const handleBlur = (e: React.FocusEvent): void => {
    const relatedTarget = e.relatedTarget as HTMLElement | null;

    // If focus explicitly moved to the floating toolbar or any of its children / dropdown popovers, stay focused
    if (
      relatedTarget &&
      (relatedTarget.closest('.floating-toolbar-container') ||
        relatedTarget.closest('.fui-floating-toolbar') ||
        (popoverSurfaceRef.current && popoverSurfaceRef.current.contains(relatedTarget)))
    ) {
      return;
    }

    if (blurTimeoutRef.current) {
      window.clearTimeout(blurTimeoutRef.current);
    }

    blurTimeoutRef.current = window.setTimeout(() => {
      // Check activeElement to see if focus is inside toolbar
      const activeEl = document.activeElement as HTMLElement | null;
      if (
        activeEl &&
        (activeEl.closest('.floating-toolbar-container') ||
          activeEl.closest('.fui-floating-toolbar') ||
          (popoverSurfaceRef.current && popoverSurfaceRef.current.contains(activeEl)))
      ) {
        return;
      }

      isInteractingWithToolbarRef.current = false;
      suppressSharePointWebPartDrag(false, elementRef.current || undefined);
      setIsFocused(false);
      if (elementRef.current) {
        onChange(elementRef.current.innerHTML);
      }
    }, 150);
  };

  const commonStyle: React.CSSProperties = {
    fontFamily: tokens.fontFamilyBase,
    color: tokens.colorNeutralForeground1,
    ...style,
    ...blockOverrideStyle
  };

  if (!isEditMode) {
    return React.createElement(tag, {
      className,
      style: commonStyle,
      dangerouslySetInnerHTML: { __html: html || placeholder }
    });
  }

  return (
    <div
      className={styles.wrapper}
      draggable={false}
      onDragStart={(e) => { e.preventDefault(); e.stopPropagation(); }}
    >
      <Popover
        open={isFocused}
        trapFocus={false}
        positioning={{
          target: elementRef.current || undefined,
          position: 'above',
          align: 'start',
          offset: 8
        }}
      >
        <PopoverSurface
          ref={popoverSurfaceRef}
          className={styles.toolbarSurface}
          onMouseEnter={() => {
            isInteractingWithToolbarRef.current = true;
          }}
          onMouseLeave={() => {
            isInteractingWithToolbarRef.current = false;
          }}
          onPointerDown={() => {
            isInteractingWithToolbarRef.current = true;
          }}
          onPointerUp={() => {
            setTimeout(() => {
              isInteractingWithToolbarRef.current = false;
            }, 150);
          }}
          onMouseDown={(e) => {
            e.preventDefault();
            e.stopPropagation();
            isInteractingWithToolbarRef.current = true;
          }}
        >
          <FloatingTextToolbar onFormat={handleFormat} />
        </PopoverSurface>
      </Popover>
      {React.createElement(tag, {
        ref: elementRef,
        contentEditable: true,
        suppressContentEditableWarning: true,
        draggable: false,
        className: `${styles.editable} ${isFocused ? styles.editableActive : ''} ${className}`,
        'data-placeholder': placeholder,
        style: commonStyle,
        onFocus: handleFocus,
        onBlur: handleBlur,
        onInput: handleInput,
        onKeyUp: handleSelectionSave,
        onMouseUp: handleSelectionSave,
        onMouseDown: (e: React.MouseEvent) => {
          e.stopPropagation();
        },
        onPointerDown: (e: React.PointerEvent) => {
          e.stopPropagation();
        },
        onDragStart: (e: React.DragEvent) => {
          e.preventDefault();
          e.stopPropagation();
        }
      })}
    </div>
  );
};
