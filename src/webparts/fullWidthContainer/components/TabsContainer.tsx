/**
 * @file TabsContainer.tsx
 * @description Tabbed navigation layout with pill buttons, active indicators, and child grid.
 */

import * as React from 'react';
import { IContainerSection } from '../models/IContainerModels';
import { FontIcon } from '@fluentui/react';
import { BlockRenderer } from './BlockRenderer';
import styles from './FullWidthContainer.module.scss.css';

export interface ITabsContainerProps {
  sections: IContainerSection[];
  enableAnimation: boolean;
  searchQuery: string;
}

export const TabsContainer: React.FC<ITabsContainerProps> = ({
  sections,
  enableAnimation,
  searchQuery
}) => {
  const [activeTabId, setActiveTabId] = React.useState<string>(
    sections.length > 0 ? sections[0].id : ''
  );

  const activeSection = sections.find((s) => s.id === activeTabId) || sections[0];

  const filteredBlocks = React.useMemo(() => {
    if (!activeSection) return [];
    if (!searchQuery.trim()) return activeSection.blocks;

    const q = searchQuery.toLowerCase();
    return activeSection.blocks.filter(
      (b) =>
        b.title.toLowerCase().includes(q) ||
        (b.description && b.description.toLowerCase().includes(q)) ||
        (b.badge && b.badge.toLowerCase().includes(q)) ||
        (b.tags && b.tags.some((t) => t.toLowerCase().includes(q)))
    );
  }, [activeSection, searchQuery]);

  if (!sections || sections.length === 0) {
    return (
      <div className={styles.emptyState}>
        <FontIcon iconName="Info" className={styles.emptyIcon} />
        <p>No container tabs configured.</p>
      </div>
    );
  }

  return (
    <div className={styles.tabsLayoutWrap}>
      {/* Navigation Pills */}
      <nav className={styles.tabNavList} role="tablist" aria-label="Container Sections">
        {sections.map((section) => {
          const isActive = section.id === (activeSection ? activeSection.id : '');
          return (
            <button
              key={section.id}
              role="tab"
              aria-selected={isActive}
              aria-controls={`tabpanel-${section.id}`}
              id={`tab-${section.id}`}
              className={`${styles.tabPillButton} ${isActive ? styles.tabPillButtonActive : ''}`}
              onClick={() => setActiveTabId(section.id)}
            >
              {section.iconName && (
                <FontIcon iconName={section.iconName} className={styles.tabPillIcon} />
              )}
              <span className={styles.tabPillLabel}>{section.title}</span>
              {section.badge && (
                <span
                  className={`${styles.tabPillBadge} ${
                    isActive ? styles.tabPillBadgeActive : ''
                  }`}
                >
                  {section.badge}
                </span>
              )}
            </button>
          );
        })}
      </nav>

      {/* Active Tab Panel */}
      {activeSection && (
        <div
          id={`tabpanel-${activeSection.id}`}
          role="tabpanel"
          aria-labelledby={`tab-${activeSection.id}`}
          className={styles.tabPanelContainer}
        >
          {activeSection.description && (
            <div className={styles.sectionHeaderBanner}>
              <p className={styles.sectionDescription}>{activeSection.description}</p>
            </div>
          )}

          {filteredBlocks.length > 0 ? (
            <div className={styles.cardsGrid}>
              {filteredBlocks.map((block) => (
                <BlockRenderer
                  key={block.id}
                  block={block}
                  enableAnimation={enableAnimation}
                />
              ))}
            </div>
          ) : (
            <div className={styles.emptyState}>
              <FontIcon iconName="SearchIssue" className={styles.emptyIcon} />
              <p>No matching content blocks found in &quot;{activeSection.title}&quot;.</p>
            </div>
          )}
        </div>
      )}
    </div>
  );
};
