/**
 * @file BlockRenderer.tsx
 * @description Renders individual child content blocks (Cards, Metrics £, Embeds, Rich Content).
 */

import * as React from 'react';
import { IContentBlock } from '../models/IContainerModels';
import { FontIcon } from '@fluentui/react';
import styles from './FullWidthContainer.module.scss.css';

export interface IBlockRendererProps {
  block: IContentBlock;
  enableAnimation: boolean;
}

export const BlockRenderer: React.FC<IBlockRendererProps> = ({ block, enableAnimation }) => {
  const animClass = enableAnimation ? styles.animatedBlock : '';

  switch (block.type) {
    case 'metric':
      return (
        <div className={`${styles.metricCard} ${animClass}`}>
          <div className={styles.metricHeader}>
            <div className={styles.metricIconWrap}>
              <FontIcon iconName={block.iconName || 'Financial'} className={styles.metricIcon} />
            </div>
            {block.badge && <span className={styles.metricBadge}>{block.badge}</span>}
          </div>
          <div className={styles.metricValue}>{block.metricValue || '£0.00'}</div>
          <div className={styles.metricTitle}>{block.title}</div>
          {block.metricTrend && (
            <div
              className={`${styles.metricTrend} ${
                block.metricTrendPositive !== false ? styles.metricTrendPositive : styles.metricTrendNegative
              }`}
            >
              <FontIcon
                iconName={block.metricTrendPositive !== false ? 'StockUp' : 'StockDown'}
                className={styles.trendIcon}
              />
              <span>{block.metricTrend}</span>
            </div>
          )}
          {block.description && <p className={styles.metricDescription}>{block.description}</p>}
        </div>
      );

    case 'embed':
      return (
        <div className={`${styles.embedCard} ${animClass}`}>
          <div className={styles.embedHeader}>
            <div className={styles.embedTitleWrap}>
              <FontIcon iconName="ViewDashboard" className={styles.embedHeaderIcon} />
              <h4 className={styles.embedTitle}>{block.title}</h4>
            </div>
            {block.linkUrl && (
              <a
                href={block.linkUrl}
                target="_blank"
                rel="noopener noreferrer"
                className={styles.embedExternalLink}
              >
                <span>{block.linkText || 'Open'}</span>
                <FontIcon iconName="OpenInNewWindow" />
              </a>
            )}
          </div>
          {block.description && <p className={styles.embedDescription}>{block.description}</p>}
          <div className={styles.iframeWrapper}>
            {block.embedUrl ? (
              <iframe
                src={block.embedUrl}
                title={block.title}
                className={styles.embedIframe}
                sandbox="allow-scripts allow-same-origin allow-popups allow-forms"
                loading="lazy"
              />
            ) : (
              <div className={styles.embedPlaceholder}>
                <FontIcon iconName="Embed" className={styles.placeholderIcon} />
                <p>No embed URL specified. Configure in block properties.</p>
              </div>
            )}
          </div>
        </div>
      );

    case 'card':
    default:
      return (
        <div className={`${styles.contentCard} ${animClass}`}>
          <div className={styles.cardTopRow}>
            <div className={styles.cardIconBox}>
              <FontIcon iconName={block.iconName || 'Document'} className={styles.cardIcon} />
            </div>
            {block.badge && <span className={styles.cardBadge}>{block.badge}</span>}
          </div>
          <h4 className={styles.cardTitle}>{block.title}</h4>
          {block.description && <p className={styles.cardDescription}>{block.description}</p>}
          {block.tags && block.tags.length > 0 && (
            <div className={styles.tagList}>
              {block.tags.map((tag, idx) => (
                <span key={idx} className={styles.tagItem}>
                  #{tag}
                </span>
              ))}
            </div>
          )}
          {block.linkUrl && (
            <div className={styles.cardFooter}>
              <a
                href={block.linkUrl}
                target="_blank"
                rel="noopener noreferrer"
                className={styles.cardActionButton}
              >
                <span>{block.linkText || 'Explore'}</span>
                <FontIcon iconName="ChevronRightSmall" className={styles.cardActionIcon} />
              </a>
            </div>
          )}
        </div>
      );
  }
};
