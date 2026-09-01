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
  WrenchRegular
} from '@fluentui/react-icons';

const useStyles = makeStyles({
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
  metricCard: {
    width: '100%',
    boxSizing: 'border-box',
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
  }
});

export interface IBlockRendererProps {
  block: IContentBlock;
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

export const BlockRenderer: React.FC<IBlockRendererProps> = ({ block }) => {
  const styles = useStyles();

  // Metric Block (£)
  if (block.type === 'metric') {
    return (
      <div className={styles.metricCard}>
        <Subtitle2>{block.title}</Subtitle2>
        <div className={styles.metricValue}>{block.metricValue}</div>

        {block.metricTrend && (
          <div className={styles.metricTrendRow}>
            <Badge
              appearance="tint"
              color={block.metricTrendPositive ? 'success' : 'informative'}
              icon={<ArrowTrendingLinesRegular />}
            >
              {block.metricTrend}
            </Badge>
          </div>
        )}

        {block.description && (
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
            {block.description}
          </Caption1>
        )}
      </div>
    );
  }

  // Embed Block
  if (block.type === 'embed') {
    return (
      <Card className={styles.card}>
        <CardHeader
          image={
            <div className={styles.iconBox}>
              <AppsRegular fontSize={20} />
            </div>
          }
          header={<Title3>{block.title}</Title3>}
          description={<Caption1>{block.description}</Caption1>}
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
    );
  }

  // Standard Card Block
  return (
    <Card className={styles.card}>
      <CardHeader
        image={
          <div className={styles.iconBox}>
            {renderFluent2Icon(block.iconName)}
          </div>
        }
        header={<Title3>{block.title}</Title3>}
        action={
          block.badge ? (
            <Badge appearance="tint" color="brand">
              {block.badge}
            </Badge>
          ) : undefined
        }
      />

      {block.description && (
        <Body1 style={{ color: tokens.colorNeutralForeground2, marginTop: tokens.spacingVerticalS }}>
          {block.description}
        </Body1>
      )}

      {block.tags && block.tags.length > 0 && (
        <div className={styles.tagsWrapper}>
          {block.tags.map((tag) => (
            <Badge key={tag} appearance="outline" size="small">
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
          >
            {block.linkText || 'View Details'}
          </Button>
        </CardFooter>
      )}
    </Card>
  );
};
