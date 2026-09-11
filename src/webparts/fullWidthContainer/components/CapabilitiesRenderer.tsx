/**
 * @file CapabilitiesRenderer.tsx
 * @description Renders the 'capabilities' composable card content part.
 * Displays a live count header: "Capabilities applied here X - [secondary label]",
 * followed by a responsive grid/flex-wrap row of styled mini-cards.
 * Each mini-card displays title, subtitle, customizable tag badges, and an optional link.
 * Built 100% with Fluent UI 2 Griffel tokens and Turner & Townsend brand styling.
 */

import * as React from 'react';
import {
  makeStyles,
  tokens,
  shorthands,
  Badge,
  Button,
  Caption1,
  Body1,
  Subtitle2
} from '@fluentui/react-components';
import {
  OpenRegular,
  EditRegular,
  AddRegular,
  ArrowRightRegular
} from '@fluentui/react-icons';
import { ICapabilityItem, ICapabilityTag } from '../models/IContainerModels';

const useStyles = makeStyles({
  container: {
    width: '100%',
    display: 'flex',
    flexDirection: 'column',
    boxSizing: 'border-box',
    ...shorthands.gap('10px')
  },
  headerRow: {
    display: 'flex',
    alignItems: 'baseline',
    flexWrap: 'wrap',
    ...shorthands.gap('6px'),
    paddingBottom: '2px'
  },
  sectionTitle: {
    fontSize: '0.8rem',
    fontWeight: 700,
    letterSpacing: '0.6px',
    textTransform: 'uppercase',
    color: tokens.colorNeutralForeground3
  },
  countBadge: {
    fontSize: '0.85rem',
    fontWeight: 700,
    color: tokens.colorBrandForeground1,
    marginLeft: '2px',
    marginRight: '2px'
  },
  secondaryLabel: {
    fontSize: '0.78rem',
    fontWeight: 600,
    letterSpacing: '0.4px',
    textTransform: 'uppercase',
    color: tokens.colorNeutralForeground4
  },
  cardsGrid: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '12px',
    width: '100%',
    alignItems: 'stretch'
  },
  miniCard: {
    flex: '1 1 210px',
    maxWidth: '300px',
    minWidth: '200px',
    boxSizing: 'border-box',
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    ...shorthands.borderRadius('10px'),
    ...shorthands.padding('14px', '14px', '12px', '14px'),
    display: 'flex',
    flexDirection: 'column',
    justifyContent: 'space-between',
    boxShadow: tokens.shadow2,
    transitionProperty: 'all',
    transitionDuration: tokens.durationNormal,
    position: 'relative',
    '&:hover': {
      ...shorthands.borderColor(tokens.colorBrandStroke1),
      boxShadow: tokens.shadow4,
      transform: 'translateY(-1px)'
    }
  },
  cardTopContent: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('4px')
  },
  capabilityTitle: {
    fontSize: '0.95rem',
    fontWeight: 700,
    color: tokens.colorNeutralForeground1,
    lineHeight: '1.25'
  },
  capabilitySubtitle: {
    fontSize: '0.78rem',
    color: tokens.colorNeutralForeground3,
    lineHeight: '1.3'
  },
  tagContainer: {
    display: 'flex',
    flexWrap: 'wrap',
    ...shorthands.gap('5px'),
    marginTop: '10px',
    marginBottom: '8px'
  },
  actionLink: {
    fontSize: '0.82rem',
    fontWeight: 600,
    color: tokens.colorBrandForeground1,
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    textDecoration: 'none',
    marginTop: '6px',
    cursor: 'pointer',
    '&:hover': {
      textDecoration: 'underline'
    }
  },
  editActionBanner: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    marginTop: '4px'
  }
});

export interface ICapabilitiesRendererProps {
  sectionLabel?: string;
  secondaryLabel?: string;
  capabilities?: ICapabilityItem[];
  alignment?: 'left' | 'center' | 'right';
  isEditMode?: boolean;
  onEdit?: () => void;
}

/**
 * Maps common tag keyword patterns to harmonious Turner & Townsend / Fluent UI 2 tint colors.
 */
export const getTagStyle = (tag: ICapabilityTag): { bg: string; fg: string; border?: string } => {
  if (tag.color && tag.color.trim()) {
    return { bg: tag.color, fg: '#1E4479' };
  }
  const l = (tag.label || '').toLowerCase();
  if (l.includes('template') || l.includes('blue')) {
    return { bg: '#CCE9F8', fg: '#005A9E', border: '#99D3F1' }; // Cyan/blue tint
  }
  if (l.includes('mandatory') || l.includes('critical') || l.includes('red') || l.includes('orange')) {
    return { bg: '#FDE7E9', fg: '#A80000', border: '#F8B6BD' }; // Soft coral/red tint
  }
  if (l.includes('insight') || l.includes('green')) {
    return { bg: '#DFF6DD', fg: '#107C10', border: '#B8ECB5' }; // Soft mint/green tint
  }
  if (l.includes('learning') || l.includes('warn') || l.includes('amber') || l.includes('yellow')) {
    return { bg: '#FFF4CE', fg: '#795E00', border: '#FFE699' }; // Warm yellow/wheat tint
  }
  if (l.includes('infra') || l.includes('note') || l.includes('grey') || l.includes('gray')) {
    return { bg: '#F2EEE7', fg: '#505A60', border: '#E0DAD0' }; // TT Mushroom / grey tint
  }
  // Default neutral soft badge
  return { bg: '#F3F2F1', fg: '#323130', border: '#E1DFDD' };
};

export const CapabilitiesRenderer: React.FC<ICapabilitiesRendererProps> = ({
  sectionLabel = 'Capabilities applied here',
  secondaryLabel = '- WHAT EACH ONE GIVES YOU IN THE PROGRAMME SCENARIO',
  capabilities = [],
  alignment = 'left',
  isEditMode = false,
  onEdit
}) => {
  const styles = useStyles();

  const count = capabilities.length;
  const justifyVal = alignment === 'center' ? 'center' : alignment === 'right' ? 'flex-end' : 'flex-start';

  return (
    <div className={styles.container} style={{ alignItems: alignment === 'center' ? 'center' : alignment === 'right' ? 'flex-end' : 'stretch' }}>
      {/* Header with live dynamic count X */}
      <div className={styles.headerRow} style={{ justifyContent: justifyVal }}>
        <span className={styles.sectionTitle}>
          {sectionLabel}
        </span>
        <span className={styles.countBadge}>
          {count}
        </span>
        {secondaryLabel && (
          <span className={styles.secondaryLabel}>
            {secondaryLabel}
          </span>
        )}
      </div>

      {/* Mini-Cards Row */}
      {capabilities.length > 0 ? (
        <div className={styles.cardsGrid} style={{ justifyContent: justifyVal }}>
          {capabilities.map((cap) => {
            const tags = cap.tags || [];
            return (
              <div key={cap.id} className={styles.miniCard}>
                <div className={styles.cardTopContent}>
                  <div className={styles.capabilityTitle}>{cap.title}</div>
                  {cap.subtitle && (
                    <div className={styles.capabilitySubtitle}>{cap.subtitle}</div>
                  )}
                  {tags.length > 0 && (
                    <div className={styles.tagContainer}>
                      {tags.map((tag) => {
                        const styleConfig = getTagStyle(tag);
                        return (
                          <span
                            key={tag.id}
                            style={{
                              display: 'inline-flex',
                              alignItems: 'center',
                              fontSize: '0.72rem',
                              fontWeight: 600,
                              lineHeight: 1,
                              padding: '3px 7px',
                              borderRadius: '12px',
                              backgroundColor: styleConfig.bg,
                              color: styleConfig.fg,
                              border: styleConfig.border ? `1px solid ${styleConfig.border}` : 'none'
                            }}
                          >
                            {tag.label}
                          </span>
                        );
                      })}
                    </div>
                  )}
                </div>

                {/* Bottom link e.g. Open the standard → */}
                {cap.linkUrl ? (
                  <a
                    href={cap.linkUrl}
                    className={styles.actionLink}
                    target="_blank"
                    rel="noreferrer"
                    onClick={(e) => isEditMode && e.preventDefault()}
                  >
                    <span>{cap.linkLabel || 'Open the standard'}</span>
                    <ArrowRightRegular fontSize={14} />
                  </a>
                ) : (
                  <span className={styles.actionLink} onClick={() => isEditMode && onEdit && onEdit()}>
                    <span>{cap.linkLabel || 'Open the standard'}</span>
                    <ArrowRightRegular fontSize={14} />
                  </span>
                )}
              </div>
            );
          })}
        </div>
      ) : (
        <div style={{ padding: '8px 0' }}>
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
            No capabilities added yet. Click &apos;Edit capabilities&apos; below to configure.
          </Caption1>
        </div>
      )}

      {/* Edit Mode configuration trigger */}
      {isEditMode && onEdit && (
        <div className={styles.editActionBanner} style={{ justifyContent: justifyVal }}>
          <Button
            size="small"
            appearance="subtle"
            icon={<EditRegular />}
            onClick={onEdit}
            style={{ fontSize: '11px', color: tokens.colorBrandForeground1 }}
          >
            Configure capabilities ({count})
          </Button>
        </div>
      )}
    </div>
  );
};
