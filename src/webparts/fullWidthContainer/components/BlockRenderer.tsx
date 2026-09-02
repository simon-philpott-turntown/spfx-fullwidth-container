/**
 * @file BlockRenderer.tsx
 * @description Renders individual child content blocks using Fluent UI 2 (Card, Badge, Button, Text, Icons, Composable Sub-elements).
 * Supports rich text formatting on titles, descriptions, and text elements with site theme and brand font inheritance.
 * Follows the Microsoft Fluent 2 Design System (https://fluent2.microsoft.design/).
 */

import * as React from 'react';
import { IContentBlock, ICardItem, ICardItemType } from '../models/IContainerModels';
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
  DismissRegular
} from '@fluentui/react-icons';
import { CardEditDialog, renderFluentIconPreview } from './CardEditDialog';
import { RichTextEditable } from './RichTextEditable';
import { InsertionBar } from './InsertionBar';
import { TermStorePicker } from './TermStorePicker';
import { LiveDataRenderer } from './LiveDataRenderer';

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
            setIsDialogOpen(true);
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

  const wrapperGridStyle: React.CSSProperties = {
    gridColumn: block.colSpan && block.colSpan > 1 ? `span ${block.colSpan}` : undefined,
    gridRow: block.rowSpan && block.rowSpan > 1 ? `span ${block.rowSpan}` : undefined,
    alignSelf: isAutoHeight ? 'start' : 'stretch',
    height: isAutoHeight ? 'auto' : '100%'
  };

  const cardDynamicStyle: React.CSSProperties = {
    height: isAutoHeight ? 'auto' : '100%'
  };

  // Metric Block
  if (block.type === 'metric') {
    return (
      <div className={styles.cardWrapper} style={wrapperGridStyle}>
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
              fontSize: '1rem',
              color: tokens.colorNeutralForeground1,
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
              fontSize: '0.85rem',
              color: tokens.colorNeutralForeground3,
              marginTop: '4px'
            }}
          />

          {/* Global Term Store Tags */}
          {block.termStoreTags && block.termStoreTags.length > 0 && (
            <div style={{ marginTop: '8px' }}>
              <TermStorePicker
                selectedTags={block.termStoreTags}
                onChange={(tags) => onUpdate && onUpdate({ termStoreTags: tags })}
                isEditMode={isEditMode}
              />
            </div>
          )}
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
      <div className={styles.cardWrapper} style={wrapperGridStyle}>
        <Card className={`${styles.card} ${isEditMode ? styles.cardEditMode : ''}`} style={cardDynamicStyle}>
          {renderCardToolbar()}
          <CardHeader
            image={
              <div className={styles.iconBox}>
                {renderFluentIconPreview(block.iconName || 'AppIconDefault')}
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
                    fontSize: '1.1rem',
                    color: tokens.colorNeutralForeground1
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
                    fontSize: '0.85rem',
                    color: tokens.colorNeutralForeground3,
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
    <div className={styles.cardWrapper} style={wrapperGridStyle}>
      <Card className={`${styles.card} ${isEditMode ? styles.cardEditMode : ''}`} style={cardDynamicStyle}>
        {renderCardToolbar()}
        <CardHeader
          image={
            <div className={styles.iconBox}>
              {renderFluentIconPreview(block.iconName)}
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
                  fontSize: '1.15rem',
                  color: tokens.colorNeutralForeground1
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
              color: tokens.colorNeutralForeground2,
              fontSize: '0.95rem',
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

        {/* Global Term Store Tags */}
        {block.termStoreTags && block.termStoreTags.length > 0 && (
          <div style={{ padding: '0 12px', marginTop: '6px' }}>
            <TermStorePicker
              selectedTags={block.termStoreTags}
              onChange={(tags) => onUpdate && onUpdate({ termStoreTags: tags })}
              isEditMode={isEditMode}
            />
          </div>
        )}

        {/* Standard tags */}
        {block.tags && block.tags.length > 0 && !block.termStoreTags && (
          <div className={styles.tagsWrapper} style={{ padding: '0 12px' }}>
            {block.tags.map((tag) => (
              <Badge key={tag} appearance="outline" size="small" className={styles.badge}>
                {tag}
              </Badge>
            ))}
          </div>
        )}

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
