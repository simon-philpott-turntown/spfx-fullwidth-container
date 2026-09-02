/**
 * @file BlockRenderer.tsx
 * @description Renders individual child content blocks using Fluent UI 2 (Card, Badge, Button, Text, Icons).
 * Follows the Microsoft Fluent 2 Design System (https://fluent2.microsoft.design/).
 */

import * as React from 'react';
import { IContentBlock } from '../models/IContainerModels';
import {
  Card,
  CardHeader,
  CardFooter,
  Badge,
  Button,
  Title3,
  Title1,
  Subtitle2,
  Body1,
  Caption1,
  makeStyles,
  shorthands,
  tokens
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
  DeleteRegular
} from '@fluentui/react-icons';

const useStyles = makeStyles({
  cardWrapper: {
    position: 'relative',
    width: '100%',
    display: 'flex'
  },
  card: {
    width: '100%',
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
    backgroundColor: tokens.colorNeutralBackground3,
    marginTop: tokens.spacingVerticalS
  },
  iconBox: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: '36px',
    height: '36px',
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground2
  },
  badge: {
    whiteSpace: 'nowrap',
    flexShrink: 0,
    display: 'inline-flex',
    alignItems: 'center'
  },
  inlineInput: {
    backgroundColor: 'transparent',
    ...shorthands.border('1px', 'dashed', 'transparent'),
    ...shorthands.borderRadius(tokens.borderRadiusSmall),
    paddingTop: '2px',
    paddingBottom: '2px',
    paddingLeft: '4px',
    paddingRight: '4px',
    fontFamily: 'inherit',
    color: 'inherit',
    outline: 'none',
    width: '100%',
    boxSizing: 'border-box',
    ':hover': {
      ...shorthands.borderColor(tokens.colorBrandStroke1),
      backgroundColor: tokens.colorNeutralBackground1Hover
    },
    ':focus': {
      ...shorthands.borderColor(tokens.colorBrandBackground),
      backgroundColor: tokens.colorNeutralBackground1
    }
  }
});

export interface IBlockRendererProps {
  block: IContentBlock;
  isEditMode?: boolean;
  onUpdate?: (updatedFields: Partial<IContentBlock>) => void;
  onDelete?: () => void;
  onEditProperties?: () => void;
}

/**
 * Maps model icon string names to official Fluent 2 Regular Icons.
 */
function renderFluent2Icon(name?: string): React.ReactElement {
  switch (name) {
    case 'Financial':
    case 'Money':
    case 'Savings':
      return <MoneyRegular fontSize={20} />;
    case 'Shield':
    case 'ComplianceAudit':
    case 'VerifiedBrand':
      return <ShieldCheckmarkRegular fontSize={20} />;
    case 'BookAnswers':
    case 'DocumentManagement':
    case 'News':
      return <DocumentRegular fontSize={20} />;
    case 'CheckList':
      return <CheckmarkCircleRegular fontSize={20} />;
    case 'TimelineProgress':
      return <ArrowTrendingLinesRegular fontSize={20} />;
    case 'Calculator':
    case 'EngineeringGroup':
      return <WrenchRegular fontSize={20} />;
    case 'Lock':
      return <LockClosedRegular fontSize={20} />;
    case 'AppIconDefault':
      return <AppsRegular fontSize={20} />;
    default:
      return <GlobeRegular fontSize={20} />;
  }
}

import { CardEditDialog } from './CardEditDialog';
import { FloatingTextToolbar } from './FloatingTextToolbar';

export const BlockRenderer: React.FC<IBlockRendererProps> = ({
  block,
  isEditMode,
  onUpdate,
  onDelete,
  onEditProperties
}) => {
  const styles = useStyles();
  const [isDialogOpen, setIsDialogOpen] = React.useState<boolean>(false);
  const [isFocused, setIsFocused] = React.useState<boolean>(false);

  // Floating Edit Mode Quick Action Toolbar
  const renderCardToolbar = (): React.ReactElement | null => {
    if (!isEditMode) return null;
    return (
      <div className={styles.cardToolbar}>
        <Button
          size="small"
          appearance="subtle"
          icon={<EditRegular />}
          title="Edit Card Properties (In-Place Dialog)"
          onClick={(e) => {
            e.stopPropagation();
            setIsDialogOpen(true);
            if (onEditProperties) {
              onEditProperties();
            }
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

  // Metric Block (£)
  if (block.type === 'metric') {
    return (
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

        {isEditMode && onUpdate ? (
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
      </div>
    );
  }

  // Embed Block
  if (block.type === 'embed') {
    return (
      <div className={styles.cardWrapper}>
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
      </div>
    );
  }

  // Standard Card Block
  return (
    <div className={styles.cardWrapper}>
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

        {block.tags && block.tags.length > 0 && (
          <div className={styles.tagsWrapper}>
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
              {block.linkText || 'Learn More'}
            </Button>
          </CardFooter>
        )}
      </Card>

      {/* In-Place Card Edit Dialog */}
      <CardEditDialog
        isOpen={isDialogOpen}
        block={block}
        onSave={(updated) => {
          if (onUpdate) {
            onUpdate(updated);
          }
        }}
        onDismiss={() => setIsDialogOpen(false)}
      />
    </div>
  );
};
