/**
 * @file BlockRenderer.tsx
 * @description Renders individual child content blocks using Fluent UI 2 (Card, Badge, Button, Text, Icons, Composable Sub-elements).
 * Supports rich text formatting on titles, descriptions, and text elements with site theme and brand font inheritance.
 * Follows the Microsoft Fluent 2 Design System (https://fluent2.microsoft.design/).
 */

import * as React from 'react';
import { IContentBlock, ICardItem, ICardItemType, ITermStoreTag } from '../models/IContainerModels';
import {
  Card,
  CardHeader,
  CardFooter,
  Badge,
  Button,
  Title3,
  Subtitle2,
  Body1,
  Caption1,
  makeStyles,
  shorthands,
  tokens,
  Divider
} from '@fluentui/react-components';
import {
  OpenRegular,
  ArrowTrendingLinesRegular,
  EditRegular,
  DeleteRegular,
  CursorClickRegular,
  MegaphoneRegular,
  DismissRegular,
  TagRegular,
  ChevronRightRegular,
  ChevronDownRegular
} from '@fluentui/react-icons';
import { CardEditDialog, renderFluentIconPreview } from './CardEditDialog';
import { RichTextEditable } from './RichTextEditable';
import { InsertionBar } from './InsertionBar';
import { TermStorePicker } from './TermStorePicker';
import { LiveDataRenderer } from './LiveDataRenderer';
import { suppressSharePointWebPartDrag } from '../utils/dragIsolation';

const useStyles = makeStyles({
  cardWrapper: {
    position: 'relative',
    width: '100%',
    display: 'flex',
    flexDirection: 'column'
  },
  card: {
    width: '100%',
    height: '100%',
    boxSizing: 'border-box',
    display: 'flex',
    flexDirection: 'column',
    justifyContent: 'space-between',
    transitionProperty: 'transform, box-shadow, border-color',
    transitionDuration: '200ms',
    transitionTimingFunction: 'cubic-bezier(0.4, 0, 0.2, 1)',
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke1),
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    backgroundColor: tokens.colorNeutralBackground1,
    ':hover': {
      transform: 'translateY(-2px)',
      boxShadow: tokens.shadow8,
      ...shorthands.borderColor(tokens.colorBrandStroke1)
    }
  },
  cardEditMode: {
    ...shorthands.border('1px', 'dashed', tokens.colorBrandStroke2)
  },
  cardToolbar: {
    position: 'absolute',
    top: '8px',
    right: '8px',
    zIndex: 10,
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap(tokens.spacingHorizontalXXS),
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.borderRadius(tokens.borderRadiusSmall),
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    boxShadow: tokens.shadow4,
    ...shorthands.padding('2px', '4px')
  },
  metricCard: {
    width: '100%',
    height: '100%',
    boxSizing: 'border-box',
    position: 'relative',
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke1),
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.padding(tokens.spacingHorizontalM),
    transitionProperty: 'transform, box-shadow, border-color',
    transitionDuration: '200ms',
    ':hover': {
      transform: 'translateY(-2px)',
      boxShadow: tokens.shadow8,
      ...shorthands.borderColor(tokens.colorBrandStroke1)
    }
  },
  metricValue: {
    color: tokens.colorBrandForeground1,
    fontSize: '2rem',
    lineHeight: '2.25rem',
    fontWeight: tokens.fontWeightBold,
    marginTop: tokens.spacingVerticalXS,
    marginBottom: tokens.spacingVerticalXXS
  },
  metricTrendRow: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap(tokens.spacingHorizontalXS),
    marginBottom: tokens.spacingVerticalS
  },
  tagsWrapper: {
    display: 'flex',
    flexWrap: 'wrap',
    ...shorthands.gap(tokens.spacingHorizontalXS),
    marginTop: tokens.spacingVerticalS
  },
  embedContainer: {
    width: '100%',
    height: '240px',
    ...shorthands.border('none'),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorNeutralBackground3
  },
  iconBox: {
    width: '40px',
    height: '40px',
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground2,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: '20px'
  },
  badge: {
    fontSize: '0.75rem',
    fontWeight: tokens.fontWeightSemibold
  },
  inlineInput: {
    width: '100%',
    border: 'none',
    backgroundColor: 'transparent',
    outline: 'none',
    fontFamily: 'inherit',
    ...shorthands.padding('2px', '4px'),
    ...shorthands.borderRadius(tokens.borderRadiusSmall),
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover
    },
    '&:focus': {
      backgroundColor: tokens.colorNeutralBackground1,
      boxShadow: `0 0 0 1.5px ${tokens.colorBrandStroke1}`
    }
  },
  innerItemsContainer: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
    ...shorthands.padding('12px', '16px')
  },
  ctaBox: {
    ...shorthands.padding('12px'),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorBrandBackground2,
    ...shorthands.border('1px', 'solid', tokens.colorBrandStroke2),
    display: 'flex',
    flexDirection: 'column',
    gap: '6px'
  },
  heroBox: {
    ...shorthands.padding('16px'),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    background: `linear-gradient(135deg, ${tokens.colorBrandBackground2} 0%, ${tokens.colorNeutralBackground2} 100%)`,
    ...shorthands.border('1px', 'solid', tokens.colorBrandStroke2),
    display: 'flex',
    flexDirection: 'column',
    gap: '8px'
  },
  galleryGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(120px, 1fr))',
    gap: '8px',
    marginTop: '6px'
  },
  galleryImg: {
    width: '100%',
    height: '90px',
    objectFit: 'cover',
    ...shorthands.borderRadius(tokens.borderRadiusSmall)
  },
  quickLinksRow: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '6px'
  },
  resizeHandleRight: {
    position: 'absolute',
    top: '15%',
    right: '-5px',
    width: '10px',
    height: '70%',
    cursor: 'ew-resize',
    zIndex: 35,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    ':hover': {
      backgroundColor: tokens.colorBrandBackground2,
      ...shorthands.borderRadius('4px')
    }
  },
  resizeHandleBottom: {
    position: 'absolute',
    bottom: '-5px',
    left: '15%',
    width: '70%',
    height: '10px',
    cursor: 'ns-resize',
    zIndex: 35,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    ':hover': {
      backgroundColor: tokens.colorBrandBackground2,
      ...shorthands.borderRadius('4px')
    }
  },
  resizeHandleCorner: {
    position: 'absolute',
    bottom: '-6px',
    right: '-6px',
    width: '16px',
    height: '16px',
    cursor: 'nwse-resize',
    zIndex: 40,
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.border('1.5px', 'solid', tokens.colorBrandStroke1),
    ...shorthands.borderRadius('4px'),
    boxShadow: tokens.shadow4,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    ':hover': {
      backgroundColor: tokens.colorBrandBackground,
      ...shorthands.borderColor(tokens.colorBrandBackground)
    }
  },
  resizeLiveBadge: {
    position: 'absolute',
    bottom: '8px',
    left: '8px',
    zIndex: 50,
    pointerEvents: 'none'
  },
  miniTagsContainer: {
    width: '100%',
    marginTop: '8px',
    boxSizing: 'border-box'
  },
  miniTagsTile: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    width: '100%',
    padding: '5px 8px',
    ...shorthands.borderRadius(tokens.borderRadiusSmall),
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    backgroundColor: tokens.colorNeutralBackground2,
    cursor: 'pointer',
    outlineStyle: 'none',
    boxSizing: 'border-box',
    transition: 'all 0.15s ease',
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground2Hover,
      ...shorthands.borderColor(tokens.colorNeutralStroke1)
    }
  },
  miniTagsPanel: {
    marginTop: '6px',
    padding: '4px 2px',
    display: 'flex',
    flexDirection: 'column',
    gap: '6px'
  }
});

export interface IBlockRendererProps {
  block: IContentBlock;
  containerGridColumns?: number;
  containerGridRows?: number;
  containerCardHeightMode?: 'auto' | 'equal';
  isEditMode?: boolean;
  onEditProperties?: () => void;
  onDelete?: () => void;
  onUpdate?: (updatedFields: Partial<IContentBlock>) => void;
}

/**
 * Calculates whether a color has low perceptual luminance to enforce WCAG text contrast.
 */
const isDarkColor = (color?: string): boolean => {
  if (!color || color === 'transparent') return false;
  const hex = color.replace('#', '').trim();
  if (hex.length === 3) {
    const r = parseInt(hex[0] + hex[0], 16);
    const g = parseInt(hex[1] + hex[1], 16);
    const b = parseInt(hex[2] + hex[2], 16);
    return (r * 299 + g * 587 + b * 114) / 1000 < 135;
  }
  if (hex.length === 6) {
    const r = parseInt(hex.substring(0, 2), 16);
    const g = parseInt(hex.substring(2, 4), 16);
    const b = parseInt(hex.substring(4, 6), 16);
    return (r * 299 + g * 587 + b * 114) / 1000 < 135;
  }
  return false;
};

export interface ICardTagsCollapsibleProps {
  termStoreTags?: ITermStoreTag[];
  tags?: string[];
  isEditMode?: boolean;
  onUpdateTermStoreTags?: (tags: ITermStoreTag[]) => void;
  isDarkBg?: boolean;
}

export const CardTagsCollapsible: React.FC<ICardTagsCollapsibleProps> = ({
  termStoreTags,
  tags,
  isEditMode = false,
  onUpdateTermStoreTags,
  isDarkBg = false
}) => {
  const [isOpen, setIsOpen] = React.useState<boolean>(false);
  const styles = useStyles();

  const termCount = termStoreTags ? termStoreTags.length : 0;
  const standardCount = tags ? tags.length : 0;
  const totalCount = termCount > 0 ? termCount : standardCount;

  if (totalCount === 0 && !isEditMode) {
    return null;
  }

  return (
    <div className={styles.miniTagsContainer}>
      <button
        type="button"
        className={styles.miniTagsTile}
        onClick={(e) => {
          e.preventDefault();
          e.stopPropagation();
          setIsOpen(!isOpen);
        }}
        title={isOpen ? 'Collapse tags' : 'Open tags'}
        aria-expanded={isOpen}
        style={{
          color: isDarkBg ? '#FFFFFF' : tokens.colorNeutralForeground2,
          borderColor: isDarkBg ? 'rgba(255, 255, 255, 0.25)' : undefined,
          backgroundColor: isDarkBg ? 'rgba(255, 255, 255, 0.08)' : undefined
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
          <TagRegular style={{ fontSize: '13px', color: isDarkBg ? '#80D1F8' : tokens.colorBrandForeground1 }} />
          <span style={{ fontSize: '0.78rem', fontWeight: 600 }}>Tags</span>
          {totalCount > 0 && (
            <span
              style={{
                fontSize: '0.68rem',
                padding: '1px 6px',
                borderRadius: '10px',
                backgroundColor: isDarkBg ? 'rgba(255, 255, 255, 0.2)' : tokens.colorNeutralBackground3,
                color: isDarkBg ? '#FFFFFF' : tokens.colorNeutralForeground2,
                fontWeight: 600
              }}
            >
              {totalCount}
            </span>
          )}
        </div>
        {isOpen ? (
          <ChevronDownRegular style={{ fontSize: '12px', opacity: 0.8 }} />
        ) : (
          <ChevronRightRegular style={{ fontSize: '12px', opacity: 0.8 }} />
        )}
      </button>

      {isOpen && (
        <div className={styles.miniTagsPanel}>
          {termStoreTags && termStoreTags.length > 0 && (
            <TermStorePicker
              selectedTags={termStoreTags}
              onChange={(updated) => onUpdateTermStoreTags && onUpdateTermStoreTags(updated)}
              isEditMode={isEditMode}
            />
          )}

          {tags && tags.length > 0 && (!termStoreTags || termStoreTags.length === 0) && (
            <div className={styles.tagsWrapper}>
              {tags.map((tag) => (
                <Badge key={tag} appearance="outline" size="small" className={styles.badge}>
                  {tag}
                </Badge>
              ))}
            </div>
          )}

          {isEditMode && totalCount === 0 && (
            <span
              style={{
                fontSize: '0.72rem',
                color: isDarkBg ? 'rgba(255, 255, 255, 0.7)' : tokens.colorNeutralForeground3,
                fontStyle: 'italic'
              }}
            >
              No tags applied yet. Use card settings to add tags.
            </span>
          )}
        </div>
      )}
    </div>
  );
};

export const BlockRenderer: React.FC<IBlockRendererProps> = ({
  block,
  containerGridColumns,
  containerGridRows,
  containerCardHeightMode = 'auto',
  isEditMode = false,
  onEditProperties,
  onDelete,
  onUpdate
}) => {
  const styles = useStyles();
  const [isDialogOpen, setIsDialogOpen] = React.useState<boolean>(false);
  const [isDraggingBoundary, setIsDraggingBoundary] = React.useState<boolean>(false);
  const [liveColSpan, setLiveColSpan] = React.useState<number>(block.colSpan || 1);
  const [liveRowSpan, setLiveRowSpan] = React.useState<number>(block.rowSpan || 1);
  const cardWrapperRef = React.useRef<HTMLDivElement>(null);

  /**
   * Listen for a global "dashboard:card-edit-open" event. When another card's
   * dialog opens it broadcasts its block.id — all other instances close.
   */
  React.useEffect(() => {
    const handleOtherCardOpened = (e: Event): void => {
      const ce = e as CustomEvent<{ blockId: string }>;
      if (ce.detail?.blockId !== block.id) {
        setIsDialogOpen(false);
      }
    };
    window.addEventListener('dashboard:card-edit-open', handleOtherCardOpened);
    return () => {
      window.removeEventListener('dashboard:card-edit-open', handleOtherCardOpened);
    };
  }, [block.id]);

  /**
   * Opens this card's edit dialog and notifies all other BlockRenderer instances
   * to close their own dialogs via the global event.
   */
  const openCardDialog = (): void => {
    window.dispatchEvent(
      new CustomEvent('dashboard:card-edit-open', { detail: { blockId: block.id } })
    );
    setIsDialogOpen(true);
  };

  React.useEffect(() => {
    setLiveColSpan(block.colSpan || 1);
    setLiveRowSpan(block.rowSpan || 1);
  }, [block.colSpan, block.rowSpan]);

  const startBoundaryDrag = (e: React.MouseEvent, type: 'right' | 'bottom' | 'corner'): void => {
    e.preventDefault();
    e.stopPropagation();

    // Suppress SharePoint's native Canvas Move Web Part controller during card resize
    suppressSharePointWebPartDrag(true, cardWrapperRef.current || undefined);

    const stopSpCanvasDrag = (dragEv: Event): void => {
      dragEv.preventDefault();
      dragEv.stopPropagation();
      if ('stopImmediatePropagation' in dragEv) {
        (dragEv as DragEvent).stopImmediatePropagation();
      }
    };
    window.addEventListener('dragstart', stopSpCanvasDrag, { capture: true, passive: false });
    window.addEventListener('selectstart', stopSpCanvasDrag, { capture: true, passive: false });

    const startX = e.clientX;
    const startY = e.clientY;
    const startCols = block.colSpan || 1;
    const startRows = block.rowSpan || 1;
    const maxCols = containerGridColumns || 4;
    const maxRows = containerGridRows || 5;

    const cardRect = cardWrapperRef.current?.getBoundingClientRect();
    const cellWidth = cardRect ? (cardRect.width / startCols) : 320;
    const cellHeight = cardRect ? (cardRect.height / startRows) : 220;

    setIsDraggingBoundary(true);
    let currentCols = startCols;
    let currentRows = startRows;

    const onMouseMove = (moveEvent: MouseEvent): void => {
      moveEvent.preventDefault();
      moveEvent.stopPropagation();
      const deltaX = moveEvent.clientX - startX;
      const deltaY = moveEvent.clientY - startY;

      if (type === 'right' || type === 'corner') {
        const colDelta = Math.round(deltaX / Math.max(90, cellWidth * 0.65));
        currentCols = Math.max(1, Math.min(maxCols, startCols + colDelta));
        setLiveColSpan(currentCols);
      }

      if (type === 'bottom' || type === 'corner') {
        const rowDelta = Math.round(deltaY / Math.max(70, cellHeight * 0.65));
        currentRows = Math.max(1, Math.min(maxRows, startRows + rowDelta));
        setLiveRowSpan(currentRows);
      }
    };

    const onMouseUp = (): void => {
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', onMouseUp);
      window.removeEventListener('dragstart', stopSpCanvasDrag, { capture: true });
      window.removeEventListener('selectstart', stopSpCanvasDrag, { capture: true });
      suppressSharePointWebPartDrag(false, cardWrapperRef.current || undefined);
      setIsDraggingBoundary(false);

      if (onUpdate && (currentCols !== (block.colSpan || 1) || currentRows !== (block.rowSpan || 1))) {
        onUpdate({ colSpan: currentCols, rowSpan: currentRows });
      }
    };

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
  };

  const renderResizeHandles = (): React.ReactElement | null => {
    if (!isEditMode) return null;
    return (
      <>
        {/* Right Boundary Drag Handle */}
        <div
          className={styles.resizeHandleRight}
          draggable={false}
          onDragStart={(e) => { e.preventDefault(); e.stopPropagation(); }}
          onPointerDown={(e) => e.stopPropagation()}
          onMouseDown={(e) => startBoundaryDrag(e, 'right')}
          title="Drag boundary to adjust column span"
        >
          <div style={{ width: '3px', height: '24px', backgroundColor: tokens.colorBrandStroke1, borderRadius: '2px', opacity: 0.8 }} />
        </div>

        {/* Bottom Boundary Drag Handle */}
        <div
          className={styles.resizeHandleBottom}
          draggable={false}
          onDragStart={(e) => { e.preventDefault(); e.stopPropagation(); }}
          onPointerDown={(e) => e.stopPropagation()}
          onMouseDown={(e) => startBoundaryDrag(e, 'bottom')}
          title="Drag boundary to adjust row span"
        >
          <div style={{ width: '24px', height: '3px', backgroundColor: tokens.colorBrandStroke1, borderRadius: '2px', opacity: 0.8 }} />
        </div>

        {/* Corner Boundary Drag Handle */}
        <div
          className={styles.resizeHandleCorner}
          draggable={false}
          onDragStart={(e) => { e.preventDefault(); e.stopPropagation(); }}
          onPointerDown={(e) => e.stopPropagation()}
          onMouseDown={(e) => startBoundaryDrag(e, 'corner')}
          title="Drag corner to adjust column & row span"
        >
          <div style={{ width: '4px', height: '4px', backgroundColor: tokens.colorBrandBackground, borderRadius: '50%' }} />
        </div>

        {isDraggingBoundary && (
          <div className={styles.resizeLiveBadge}>
            <Badge appearance="filled" color="brand" size="small">
              📐 Span {liveColSpan} cols × {liveRowSpan} rows (Preview)
            </Badge>
          </div>
        )}
      </>
    );
  };

  const renderCardToolbar = (): React.ReactElement | null => {
    if (!isEditMode) return null;
    return (
      <div className={styles.cardToolbar}>
        <Button
          size="small"
          appearance="subtle"
          icon={<EditRegular />}
          title="Edit card properties"
          onClick={(e) => {
            e.stopPropagation();
            openCardDialog();
          }}
        />
        {onDelete && (
          <Button
            size="small"
            appearance="subtle"
            icon={<DeleteRegular />}
            title="Delete card"
            onClick={(e) => {
              e.stopPropagation();
              onDelete();
            }}
          />
        )}
      </div>
    );
  };

  // Helper: insert a sub-item into card
  const handleInsertCardItem = (index: number, itemType: ICardItemType): void => {
    if (!onUpdate) return;
    const currentItems = block.items ? [...block.items] : [];
    const newItem: ICardItem = {
      id: `item-${Date.now()}`,
      type: itemType,
      text: itemType === 'text' ? 'Text section content' : undefined,
      buttonLabel: itemType === 'button' ? 'Action Button' : undefined,
      buttonUrl: itemType === 'button' ? '#' : undefined,
      ctaHeading: itemType === 'cta' ? 'Call to Action' : undefined,
      ctaDescription: itemType === 'cta' ? 'Guidance details here.' : undefined,
      ctaButtonText: itemType === 'cta' ? 'Proceed' : undefined,
      imageUrl: itemType === 'image' ? 'https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?w=500&q=80' : undefined,
      videoUrl: itemType === 'video' ? 'https://www.w3schools.com/html/mov_bbb.mp4' : undefined,
      quickLinks: itemType === 'quickLinks' ? [
        { label: 'Policy Document', url: '#' },
        { label: 'Process Map', url: '#' }
      ] : undefined,
      galleryImages: itemType === 'gallery' ? [
        { url: 'https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?w=300&q=80' },
        { url: 'https://images.unsplash.com/photo-1497366216548-37526070297c?w=300&q=80' }
      ] : undefined,
      liveDataConfig: itemType === 'liveData' ? {
        apiUrl: 'demo-api/metric',
        jsonPath: 'data.metric',
        prefix: '£',
        refreshIntervalSeconds: 30
      } : undefined,
      termStoreTags: itemType === 'termStoreTags' ? [
        { id: 'sec-infra', label: 'Infrastructure', termSetName: 'Our Sectors' },
        { id: 'seg-comm', label: 'Commercial advisory', termSetName: 'Our Segments' }
      ] : undefined
    };

    currentItems.splice(index, 0, newItem);
    onUpdate({ items: currentItems });
  };

  const handleRemoveCardItem = (itemId: string): void => {
    if (!onUpdate || !block.items) return;
    onUpdate({ items: block.items.filter((i) => i.id !== itemId) });
  };

  const renderInnerItem = (item: ICardItem, idx: number): React.ReactElement => {
    return (
      <div key={item.id} style={{ position: 'relative', width: '100%' }}>
        {isEditMode && (
          <div style={{ position: 'absolute', right: 0, top: 0, zIndex: 5 }}>
            <Button
              size="small"
              appearance="subtle"
              icon={<DismissRegular />}
              onClick={() => handleRemoveCardItem(item.id)}
              title="Remove item"
            />
          </div>
        )}

        {item.type === 'text' && (
          <div>
            <RichTextEditable
              html={item.text || ''}
              isEditMode={isEditMode}
              placeholder="Text section content"
              onChange={(newHtml) => {
                if (block.items && onUpdate) {
                  const updatedItems = [...block.items];
                  updatedItems[idx].text = newHtml;
                  onUpdate({ items: updatedItems });
                }
              }}
              style={{
                color: tokens.colorNeutralForeground1,
                fontSize: '0.95rem'
              }}
            />
          </div>
        )}

        {item.type === 'button' && (
          <Button
            appearance={item.buttonVariant || 'primary'}
            as="a"
            href={item.buttonUrl || '#'}
            icon={<CursorClickRegular />}
          >
            {item.buttonLabel || 'Action Button'}
          </Button>
        )}

        {item.type === 'cta' && (
          <div className={styles.ctaBox}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
              <MegaphoneRegular style={{ color: tokens.colorBrandForeground1 }} />
              <Subtitle2>{item.ctaHeading || 'Call to Action'}</Subtitle2>
            </div>
            <Caption1 style={{ color: tokens.colorNeutralForeground2 }}>
              {item.ctaDescription || 'Guidance details here.'}
            </Caption1>
            {item.ctaButtonText && (
              <Button size="small" appearance="primary" as="a" href={item.ctaButtonUrl || '#'}>
                {item.ctaButtonText}
              </Button>
            )}
          </div>
        )}

        {item.type === 'hero' && (
          <div className={styles.heroBox}>
            <Title3>{block.title}</Title3>
            <Body1 style={{ color: tokens.colorNeutralForeground2 }}>{block.description}</Body1>
          </div>
        )}

        {item.type === 'divider' && <Divider style={{ margin: '8px 0' }} />}

        {item.type === 'image' && item.imageUrl && (
          <div style={{ width: '100%' }}>
            <img
              src={item.imageUrl}
              alt={item.imageAlt || 'Card image'}
              style={{ width: '100%', maxHeight: '180px', objectFit: 'cover', borderRadius: '6px' }}
            />
            {item.imageCaption && (
              <Caption1 style={{ color: tokens.colorNeutralForeground4, display: 'block', marginTop: '4px' }}>
                {item.imageCaption}
              </Caption1>
            )}
          </div>
        )}

        {item.type === 'gallery' && item.galleryImages && (
          <div className={styles.galleryGrid}>
            {item.galleryImages.map((img, i) => (
              <img key={i} src={img.url} alt="Gallery" className={styles.galleryImg} />
            ))}
          </div>
        )}

        {item.type === 'quickLinks' && item.quickLinks && (
          <div className={styles.quickLinksRow}>
            {item.quickLinks.map((ql, i) => (
              <Button key={i} size="small" appearance="outline" as="a" href={ql.url} icon={<OpenRegular />}>
                {ql.label}
              </Button>
            ))}
          </div>
        )}

        {item.type === 'video' && item.videoUrl && (
          <video
            src={item.videoUrl}
            controls
            style={{ width: '100%', maxHeight: '180px', borderRadius: '6px' }}
          />
        )}

        {item.type === 'liveData' && item.liveDataConfig && (
          <div>
            {isEditMode ? (
              <div
                style={{
                  padding: '10px',
                  borderRadius: tokens.borderRadiusMedium,
                  backgroundColor: tokens.colorNeutralBackground2,
                  border: `1px solid ${tokens.colorNeutralStroke2}`,
                  display: 'flex',
                  flexDirection: 'column',
                  gap: '8px',
                  marginBottom: '8px'
                }}
              >
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <Caption1 style={{ color: tokens.colorNeutralForeground2, fontWeight: 600 }}>
                    Real-time live data API (optional)
                  </Caption1>
                  <Badge appearance="tint" color="brand" size="small">
                    Live REST
                  </Badge>
                </div>
                <input
                  className={styles.inlineInput}
                  style={{
                    backgroundColor: tokens.colorNeutralBackground1,
                    border: `1px solid ${tokens.colorNeutralStroke1}`,
                    padding: '4px 8px',
                    borderRadius: '4px',
                    fontSize: '0.85rem'
                  }}
                  placeholder="API endpoint URL (e.g. https://... or demo-api/burnDown)"
                  value={item.liveDataConfig.apiUrl || ''}
                  onChange={(e) => {
                    if (block.items && onUpdate) {
                      const updatedItems = [...block.items];
                      updatedItems[idx].liveDataConfig = {
                        ...item.liveDataConfig,
                        apiUrl: e.target.value,
                        jsonPath: item.liveDataConfig?.jsonPath || 'value',
                        prefix: item.liveDataConfig?.prefix || '£',
                        refreshIntervalSeconds: item.liveDataConfig?.refreshIntervalSeconds || 30
                      };
                      onUpdate({ items: updatedItems });
                    }
                  }}
                />
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px' }}>
                  <input
                    className={styles.inlineInput}
                    style={{
                      backgroundColor: tokens.colorNeutralBackground1,
                      border: `1px solid ${tokens.colorNeutralStroke1}`,
                      padding: '4px 8px',
                      borderRadius: '4px',
                      fontSize: '0.85rem'
                    }}
                    placeholder="JSON path (e.g. data.metric)"
                    value={item.liveDataConfig.jsonPath || ''}
                    onChange={(e) => {
                      if (block.items && onUpdate) {
                        const updatedItems = [...block.items];
                        updatedItems[idx].liveDataConfig = {
                          ...item.liveDataConfig,
                          apiUrl: item.liveDataConfig?.apiUrl || '',
                          jsonPath: e.target.value,
                          prefix: item.liveDataConfig?.prefix || '£',
                          refreshIntervalSeconds: item.liveDataConfig?.refreshIntervalSeconds || 30
                        };
                        onUpdate({ items: updatedItems });
                      }
                    }}
                  />
                  <input
                    className={styles.inlineInput}
                    style={{
                      backgroundColor: tokens.colorNeutralBackground1,
                      border: `1px solid ${tokens.colorNeutralStroke1}`,
                      padding: '4px 8px',
                      borderRadius: '4px',
                      fontSize: '0.85rem'
                    }}
                    placeholder="Prefix (e.g. £)"
                    value={item.liveDataConfig.prefix || ''}
                    onChange={(e) => {
                      if (block.items && onUpdate) {
                        const updatedItems = [...block.items];
                        updatedItems[idx].liveDataConfig = {
                          ...item.liveDataConfig,
                          apiUrl: item.liveDataConfig?.apiUrl || '',
                          jsonPath: item.liveDataConfig?.jsonPath || 'value',
                          prefix: e.target.value,
                          refreshIntervalSeconds: item.liveDataConfig?.refreshIntervalSeconds || 30
                        };
                        onUpdate({ items: updatedItems });
                      }
                    }}
                  />
                </div>
                <LiveDataRenderer config={item.liveDataConfig} isEditMode={isEditMode} />
              </div>
            ) : (
              <LiveDataRenderer config={item.liveDataConfig} isEditMode={isEditMode} />
            )}
          </div>
        )}

        {item.type === 'termStoreTags' && (
          <TermStorePicker
            selectedTags={item.termStoreTags || []}
            onChange={(tags) => {
              if (block.items && onUpdate) {
                const updatedItems = [...block.items];
                updatedItems[idx].termStoreTags = tags;
                onUpdate({ items: updatedItems });
              }
            }}
            isEditMode={isEditMode}
          />
        )}
      </div>
    );
  };

  const effectiveHeightMode = (block.heightMode && block.heightMode !== 'default')
    ? block.heightMode
    : (containerCardHeightMode || 'auto');
  const isAutoHeight = effectiveHeightMode === 'auto';

  const currentEffectiveCol = isDraggingBoundary ? liveColSpan : (block.colSpan || 1);
  const currentEffectiveRow = isDraggingBoundary ? liveRowSpan : (block.rowSpan || 1);

  const wrapperGridStyle: React.CSSProperties = {
    gridColumn: currentEffectiveCol > 1 ? `span ${currentEffectiveCol}` : undefined,
    gridRow: currentEffectiveRow > 1 ? `span ${currentEffectiveRow}` : undefined,
    alignSelf: isAutoHeight ? 'start' : 'stretch',
    height: isAutoHeight ? 'auto' : '100%',
    position: 'relative',
    zIndex: isDraggingBoundary ? 40 : 1,
    transition: isDraggingBoundary ? 'none' : 'grid-column 0.15s ease, grid-row 0.15s ease'
  };

  const isDarkBg = isDarkColor(block.backgroundColor);
  const cardDynamicStyle: React.CSSProperties = {
    height: isAutoHeight ? 'auto' : '100%',
    backgroundColor: block.backgroundColor || undefined,
    color: block.textColor || (isDarkBg ? '#FFFFFF' : undefined),
    fontFamily: block.fontFamily || undefined,
    borderColor: isDraggingBoundary
      ? tokens.colorBrandStroke1
      : (isDarkBg ? 'rgba(255, 255, 255, 0.2)' : undefined),
    boxShadow: isDraggingBoundary
      ? `0 0 0 2px ${tokens.colorBrandStroke1}, 0 8px 24px rgba(0, 144, 220, 0.35)`
      : undefined
  };

  // Metric Block
  if (block.type === 'metric') {
    return (
      <div ref={cardWrapperRef} className={styles.cardWrapper} style={wrapperGridStyle}>
        {renderResizeHandles()}
        <div className={`${styles.metricCard} ${isEditMode ? styles.cardEditMode : ''}`} style={cardDynamicStyle}>
          {renderCardToolbar()}
          {/* Metric Title with Rich Text Editing */}
          <RichTextEditable
            tag="h3"
            html={block.title || ''}
            isEditMode={isEditMode}
            placeholder="Metric title"
            onChange={(newTitle) => onUpdate && onUpdate({ title: newTitle })}
            style={{
              fontWeight: 600,
              fontSize: block.titleFontSize || '1rem',
              color: block.textColor || (isDarkBg ? '#FFFFFF' : tokens.colorNeutralForeground1),
              marginBottom: '4px'
            }}
          />

          {block.liveDataConfig ? (
            <div style={{ marginTop: '6px', marginBottom: '6px' }}>
              <LiveDataRenderer config={block.liveDataConfig} isEditMode={isEditMode} />
            </div>
          ) : isEditMode && onUpdate ? (
            <input
              className={styles.inlineInput}
              style={{
                fontSize: '1.75rem',
                fontWeight: 700,
                color: tokens.colorBrandForeground1,
                fontFamily: tokens.fontFamilyBase,
                lineHeight: '2rem',
                margin: '4px 0'
              }}
              value={block.metricValue || ''}
              placeholder="£1,420,000"
              onChange={(e) => onUpdate({ metricValue: e.target.value })}
            />
          ) : (
            <div className={styles.metricValue}>{block.metricValue}</div>
          )}

          {block.metricTrend && (
            <div className={styles.metricTrendRow}>
              <Badge
                appearance="tint"
                color={block.metricTrendPositive ? 'success' : 'informative'}
                icon={<ArrowTrendingLinesRegular />}
                className={styles.badge}
              >
                {block.metricTrend}
              </Badge>
            </div>
          )}

          {/* Metric Description with Rich Text Editing */}
          <RichTextEditable
            tag="p"
            html={block.description || ''}
            isEditMode={isEditMode}
            placeholder="Metric summary"
            onChange={(newDesc) => onUpdate && onUpdate({ description: newDesc })}
            style={{
              fontSize: block.bodyFontSize || '0.85rem',
              color: block.textColor || (isDarkBg ? 'rgba(255,255,255,0.85)' : tokens.colorNeutralForeground3),
              marginTop: '4px'
            }}
          />

          {/* Mini Collapsible Tags */}
          <CardTagsCollapsible
            termStoreTags={block.termStoreTags}
            tags={block.tags}
            isEditMode={isEditMode}
            onUpdateTermStoreTags={(tags) => onUpdate && onUpdate({ termStoreTags: tags })}
            isDarkBg={isDarkBg}
          />
        </div>

        <CardEditDialog
          isOpen={isDialogOpen}
          block={block}
          maxColumns={containerGridColumns}
          maxRows={containerGridRows}
          onSave={(updated) => onUpdate && onUpdate(updated)}
          onDismiss={() => setIsDialogOpen(false)}
        />
      </div>
    );
  }

  // Embed Block
  if (block.type === 'embed') {
    return (
      <div ref={cardWrapperRef} className={styles.cardWrapper} style={wrapperGridStyle}>
        {renderResizeHandles()}
        <Card className={`${styles.card} ${isEditMode ? styles.cardEditMode : ''}`} style={cardDynamicStyle}>
          {renderCardToolbar()}
          <CardHeader
            image={
              <div
                className={styles.iconBox}
                style={{
                  backgroundColor: block.showIconBackground !== false
                    ? (block.iconBackgroundColor || tokens.colorBrandBackground2)
                    : 'transparent',
                  border: block.showIconBackground !== false
                    ? (block.iconBackgroundColor ? 'none' : `1px solid ${tokens.colorNeutralStroke2}`)
                    : 'none',
                  color: block.iconColor || tokens.colorBrandForeground2
                }}
              >
                {renderFluentIconPreview(block.iconName || 'AppIconDefault', block.iconColor)}
              </div>
            }
            header={
              <div style={{ width: '100%' }}>
                <RichTextEditable
                  tag="h3"
                  html={block.title || ''}
                  isEditMode={isEditMode}
                  placeholder="Tool title"
                  onChange={(newTitle) => onUpdate && onUpdate({ title: newTitle })}
                  style={{
                    fontWeight: 600,
                    fontSize: block.titleFontSize || '1.1rem',
                    color: block.textColor || (isDarkBg ? '#FFFFFF' : tokens.colorNeutralForeground1)
                  }}
                />
              </div>
            }
            description={
              <div style={{ width: '100%' }}>
                <RichTextEditable
                  tag="p"
                  html={block.description || ''}
                  isEditMode={isEditMode}
                  placeholder="Tool summary"
                  onChange={(newDesc) => onUpdate && onUpdate({ description: newDesc })}
                  style={{
                    fontSize: block.bodyFontSize || '0.85rem',
                    color: block.textColor || (isDarkBg ? 'rgba(255, 255, 255, 0.85)' : tokens.colorNeutralForeground3),
                    marginTop: '2px'
                  }}
                />
              </div>
            }
          />
          {block.embedUrl && (
            <iframe
              src={block.embedUrl}
              title={block.title}
              className={styles.embedContainer}
              sandbox="allow-scripts allow-same-origin"
            />
          )}
          {block.linkUrl && (
            <CardFooter>
              <Button
                appearance="subtle"
                icon={<OpenRegular />}
                iconPosition="after"
                as="a"
                href={block.linkUrl}
                target="_blank"
              >
                {block.linkText || 'Open link'}
              </Button>
            </CardFooter>
          )}
        </Card>

        <CardEditDialog
          isOpen={isDialogOpen}
          block={block}
          maxColumns={containerGridColumns}
          maxRows={containerGridRows}
          onSave={(updated) => onUpdate && onUpdate(updated)}
          onDismiss={() => setIsDialogOpen(false)}
        />
      </div>
    );
  }

  // Standard / Composable Card Block
  return (
    <div ref={cardWrapperRef} className={styles.cardWrapper} style={wrapperGridStyle}>
      {renderResizeHandles()}
      <Card className={`${styles.card} ${isEditMode ? styles.cardEditMode : ''}`} style={cardDynamicStyle}>
        {renderCardToolbar()}
        <CardHeader
          image={
            <div
              className={styles.iconBox}
              style={{
                backgroundColor: block.showIconBackground !== false
                  ? (block.iconBackgroundColor || tokens.colorBrandBackground2)
                  : 'transparent',
                border: block.showIconBackground !== false
                  ? (block.iconBackgroundColor ? 'none' : `1px solid ${tokens.colorNeutralStroke2}`)
                  : 'none',
                color: block.iconColor || tokens.colorBrandForeground2
              }}
            >
              {renderFluentIconPreview(block.iconName, block.iconColor)}
            </div>
          }
          header={
            <div style={{ width: '100%' }}>
              <RichTextEditable
                tag="h3"
                html={block.title || ''}
                isEditMode={isEditMode}
                placeholder="Card title"
                onChange={(newTitle) => onUpdate && onUpdate({ title: newTitle })}
                style={{
                  fontWeight: 600,
                  fontSize: block.titleFontSize || '1.15rem',
                  color: block.textColor || (isDarkBg ? '#FFFFFF' : tokens.colorNeutralForeground1)
                }}
              />
            </div>
          }
          action={
            block.badge ? (
              <Badge appearance="tint" color="brand" className={styles.badge}>
                {block.badge}
              </Badge>
            ) : undefined
          }
        />

        {/* Card Body Description with Rich Text Formatting */}
        <div style={{ padding: '0 16px', marginTop: '6px' }}>
          <RichTextEditable
            tag="p"
            html={block.description || ''}
            isEditMode={isEditMode}
            placeholder="Card summary"
            onChange={(newDesc) => onUpdate && onUpdate({ description: newDesc })}
            style={{
              color: block.textColor || (isDarkBg ? 'rgba(255, 255, 255, 0.85)' : tokens.colorNeutralForeground2),
              fontSize: block.bodyFontSize || '0.95rem',
              lineHeight: '1.4rem'
            }}
          />
        </div>

        {/* Nested Composable Items & Insertion Bars */}
        {block.items && block.items.length > 0 && (
          <div className={styles.innerItemsContainer}>
            {block.items.map((item, index) => (
              <React.Fragment key={item.id}>
                {isEditMode && (
                  <InsertionBar
                    onInsert={(type) => handleInsertCardItem(index, type)}
                    contextTitle={`Insert into ${block.title}`}
                  />
                )}
                {renderInnerItem(item, index)}
              </React.Fragment>
            ))}
            {isEditMode && (
              <InsertionBar
                onInsert={(type) => handleInsertCardItem(block.items ? block.items.length : 0, type)}
                contextTitle={`Insert into ${block.title}`}
              />
            )}
          </div>
        )}

        {/* Empty state insertion bar for Edit Mode when no items exist */}
        {isEditMode && (!block.items || block.items.length === 0) && (
          <div style={{ padding: '0 12px' }}>
            <InsertionBar
              alwaysVisible={true}
              onInsert={(type) => handleInsertCardItem(0, type)}
              contextTitle={`Add content to ${block.title}`}
            />
          </div>
        )}

        {/* Live Data API Field */}
        {block.liveDataConfig && (
          <div style={{ padding: '0 12px', marginTop: '6px' }}>
            <LiveDataRenderer config={block.liveDataConfig} isEditMode={isEditMode} />
          </div>
        )}

        {/* Mini Collapsible Tags */}
        <div style={{ padding: '0 12px' }}>
          <CardTagsCollapsible
            termStoreTags={block.termStoreTags}
            tags={block.tags}
            isEditMode={isEditMode}
            onUpdateTermStoreTags={(tags) => onUpdate && onUpdate({ termStoreTags: tags })}
            isDarkBg={isDarkBg}
          />
        </div>

        {block.linkUrl && (
          <CardFooter style={{ marginTop: tokens.spacingVerticalM }}>
            <Button
              appearance="subtle"
              icon={<OpenRegular />}
              iconPosition="after"
              as="a"
              href={block.linkUrl}
              target="_blank"
            >
              {block.linkText || 'Action link'}
            </Button>
          </CardFooter>
        )}
      </Card>

      {/* In-Place Card Edit Dialog */}
      <CardEditDialog
        isOpen={isDialogOpen}
        block={block}
        maxColumns={containerGridColumns}
        maxRows={containerGridRows}
        onSave={(updated) => onUpdate && onUpdate(updated)}
        onDismiss={() => setIsDialogOpen(false)}
      />
    </div>
  );
};
