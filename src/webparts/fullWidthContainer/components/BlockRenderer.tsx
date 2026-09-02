/**
 * @file BlockRenderer.tsx
 * @description Renders individual child content blocks using Fluent UI 2 (Card, Badge, Button, Text, Icons, Composable Sub-elements).
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
  DocumentRegular,
  MoneyRegular,
  ArrowTrendingLinesRegular,
  ShieldCheckmarkRegular,
  GlobeRegular,
  AppsRegular,
  LockClosedRegular,
  CheckmarkCircleRegular,
  WrenchRegular,
  EditRegular,
  DeleteRegular,
  CursorClickRegular,
  MegaphoneRegular,
  DismissRegular
} from '@fluentui/react-icons';
import { CardEditDialog } from './CardEditDialog';
import { FloatingTextToolbar } from './FloatingTextToolbar';
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
  isEditMode?: boolean;
  onEditProperties?: () => void;
  onDelete?: () => void;
  onUpdate?: (updatedFields: Partial<IContentBlock>) => void;
}

export const BlockRenderer: React.FC<IBlockRendererProps> = ({
  block,
  containerGridColumns,
  containerGridRows,
  isEditMode = false,
  onEditProperties,
  onDelete,
  onUpdate
}) => {
  const styles = useStyles();
  const [isDialogOpen, setIsDialogOpen] = React.useState<boolean>(false);
  const [isFocused, setIsFocused] = React.useState<boolean>(false);

  const renderFluent2Icon = (iconName?: string): React.ReactElement => {
    switch (iconName) {
      case 'Financial':
        return <MoneyRegular fontSize={20} />;
      case 'ComplianceAudit':
        return <ShieldCheckmarkRegular fontSize={20} />;
      case 'CheckList':
        return <CheckmarkCircleRegular fontSize={20} />;
      case 'TimelineProgress':
        return <ArrowTrendingLinesRegular fontSize={20} />;
      case 'Calculator':
        return <WrenchRegular fontSize={20} />;
      case 'AppIconDefault':
        return <AppsRegular fontSize={20} />;
      case 'Lock':
        return <LockClosedRegular fontSize={20} />;
      case 'Globe':
        return <GlobeRegular fontSize={20} />;
      case 'BookAnswers':
      default:
        return <DocumentRegular fontSize={20} />;
    }
  };

  const renderCardToolbar = (): React.ReactElement | null => {
    if (!isEditMode) return null;
    return (
      <div className={styles.cardToolbar}>
        <Button
          size="small"
          appearance="subtle"
          icon={<EditRegular />}
          title="Edit Card Properties"
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
            title="Delete Card"
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
        jsonPath: 'value',
        prefix: '£',
        refreshIntervalSeconds: 30
      } : undefined,
      termStoreTags: itemType === 'termStoreTags' ? [
        { id: 't-comm', label: 'Commercial' },
        { id: 't-p0', label: 'P0 Critical' }
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
          <Body1 style={{ color: tokens.colorNeutralForeground2 }}>
            {item.text || 'Text section content'}
          </Body1>
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
          <LiveDataRenderer config={item.liveDataConfig} isEditMode={isEditMode} />
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

  const wrapperGridStyle: React.CSSProperties = {
    gridColumn: block.colSpan && block.colSpan > 1 ? `span ${block.colSpan}` : undefined,
    gridRow: block.rowSpan && block.rowSpan > 1 ? `span ${block.rowSpan}` : undefined
  };

  // Metric Block
  if (block.type === 'metric') {
    return (
      <div className={styles.cardWrapper} style={wrapperGridStyle}>
        <div className={`${styles.metricCard} ${isEditMode ? styles.cardEditMode : ''}`}>
          {renderCardToolbar()}
          {isEditMode && onUpdate ? (
            <input
              className={styles.inlineInput}
              style={{ fontWeight: 600, fontSize: '1rem' }}
              value={block.title || ''}
              placeholder="Metric title"
              onChange={(e) => onUpdate({ title: e.target.value })}
            />
          ) : (
            <Subtitle2>{block.title}</Subtitle2>
          )}

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
                marginTop: '4px',
                marginBottom: '4px'
              }}
              value={block.metricValue || ''}
              placeholder="£100,000"
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

          {isEditMode && onUpdate ? (
            <input
              className={styles.inlineInput}
              style={{ fontSize: '0.85rem', color: tokens.colorNeutralForeground3 }}
              value={block.description || ''}
              placeholder="Metric description summary"
              onChange={(e) => onUpdate({ description: e.target.value })}
            />
          ) : (
            block.description && (
              <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                {block.description}
              </Caption1>
            )
          )}

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
        <Card className={`${styles.card} ${isEditMode ? styles.cardEditMode : ''}`}>
          {renderCardToolbar()}
          <CardHeader
            image={
              <div className={styles.iconBox}>
                <AppsRegular fontSize={20} />
              </div>
            }
            header={
              isEditMode && onUpdate ? (
                <input
                  className={styles.inlineInput}
                  style={{ fontWeight: 600, fontSize: '1.1rem' }}
                  value={block.title || ''}
                  placeholder="Tool title"
                  onChange={(e) => onUpdate({ title: e.target.value })}
                />
              ) : (
                <Title3>{block.title}</Title3>
              )
            }
            description={
              isEditMode && onUpdate ? (
                <input
                  className={styles.inlineInput}
                  style={{ fontSize: '0.85rem', color: tokens.colorNeutralForeground3 }}
                  value={block.description || ''}
                  placeholder="Tool description summary"
                  onChange={(e) => onUpdate({ description: e.target.value })}
                />
              ) : (
                <Caption1>{block.description}</Caption1>
              )
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
                {block.linkText || 'Open Tool'}
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
      <Card className={`${styles.card} ${isEditMode ? styles.cardEditMode : ''}`}>
        {renderCardToolbar()}
        <CardHeader
          image={
            <div className={styles.iconBox}>
              {renderFluent2Icon(block.iconName)}
            </div>
          }
          header={
            isEditMode && onUpdate ? (
              <input
                className={styles.inlineInput}
                style={{ fontWeight: 600, fontSize: '1.1rem' }}
                value={block.title || ''}
                placeholder="Card title"
                onChange={(e) => onUpdate({ title: e.target.value })}
              />
            ) : (
              <Title3>{block.title}</Title3>
            )
          }
          action={
            block.badge ? (
              <Badge appearance="tint" color="brand" className={styles.badge}>
                {block.badge}
              </Badge>
            ) : undefined
          }
        />

        {/* Card Body Description */}
        {isEditMode && onUpdate ? (
          <div>
            {isFocused && (
              <div style={{ marginBottom: '6px' }}>
                <FloatingTextToolbar />
              </div>
            )}
            <textarea
              className={styles.inlineInput}
              rows={2}
              style={{
                color: tokens.colorNeutralForeground2,
                marginTop: tokens.spacingVerticalS,
                resize: 'vertical'
              }}
              value={block.description || ''}
              placeholder="Card description summary"
              onFocus={() => setIsFocused(true)}
              onBlur={() => {
                setTimeout(() => setIsFocused(false), 250);
              }}
              onChange={(e) => onUpdate({ description: e.target.value })}
            />
          </div>
        ) : (
          block.description && (
            <Body1 style={{ color: tokens.colorNeutralForeground2, marginTop: tokens.spacingVerticalS }}>
              {block.description}
            </Body1>
          )
        )}

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
              {block.linkText || 'Learn more'}
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
