/**
 * @file AccordionContainer.tsx
 * @description Accordion collapsible panels view for the Full-Width Container Web Part.
 */

import * as React from 'react';
import { IContainerSection } from '../models/IContainerModels';
import { FontIcon } from '@fluentui/react';
import { BlockRenderer } from './BlockRenderer';
import styles from './FullWidthContainer.module.scss.css';

export interface IAccordionContainerProps {
  sections: IContainerSection[];
  enableAnimation: boolean;
  searchQuery: string;
}

export const AccordionContainer: React.FC<IAccordionContainerProps> = ({
  sections,
  enableAnimation,
  searchQuery
}) => {
  // Store set of open section IDs
  const [openSectionIds, setOpenSectionIds] = React.useState<Record<string, boolean>>(() => {
    const initial: Record<string, boolean> = {};
    if (sections.length > 0) {
      initial[sections[0].id] = true; // open first section by default
    }
    return initial;
  });

  const toggleSection = (id: string): void => {
    setOpenSectionIds((prev) => ({
      ...prev,
      [id]: !prev[id]
    }));
  };

  const expandAll = (): void => {
    const allOpen: Record<string, boolean> = {};
    sections.forEach((s) => {
      allOpen[s.id] = true;
    });
    setOpenSectionIds(allOpen);
  };

  const collapseAll = (): void => {
    setOpenSectionIds({});
  };

  if (!sections || sections.length === 0) {
    return (
      <div className={styles.emptyState}>
        <FontIcon iconName="Info" className={styles.emptyIcon} />
        <p>No accordion sections configured.</p>
      </div>
    );
  }

  return (
    <div className={styles.accordionLayoutWrap}>
      {/* Global Expand / Collapse Toolbar */}
      <div className={styles.accordionToolbar}>
        <span className={styles.toolbarCount}>
          {sections.length} {sections.length === 1 ? 'Section' : 'Sections'}
        </span>
        <div className={styles.toolbarButtons}>
          <button type="button" className={styles.toolbarLink} onClick={expandAll}>
            <FontIcon iconName="DoubleChevronDown" />
            <span>Expand All</span>
          </button>
          <span className={styles.toolbarDivider}>•</span>
          <button type="button" className={styles.toolbarLink} onClick={collapseAll}>
            <FontIcon iconName="DoubleChevronUp" />
            <span>Collapse All</span>
          </button>
        </div>
      </div>

      {/* Accordion Panels List */}
      <div className={styles.accordionList}>
        {sections.map((section) => {
          const isOpen = !!openSectionIds[section.id];

          // Filter blocks by search query
          const q = searchQuery.toLowerCase().trim();
          const filteredBlocks = q
            ? section.blocks.filter(
                (b) =>
                  b.title.toLowerCase().includes(q) ||
                  (b.description && b.description.toLowerCase().includes(q)) ||
                  (b.badge && b.badge.toLowerCase().includes(q)) ||
                  (b.tags && b.tags.some((t) => t.toLowerCase().includes(q)))
              )
            : section.blocks;

          // If search is active and matches exist in this section, force open
          const showOpen = q && filteredBlocks.length > 0 ? true : isOpen;

          return (
            <div
              key={section.id}
              className={`${styles.accordionItem} ${showOpen ? styles.accordionItemExpanded : ''}`}
            >
              <button
                type="button"
                className={styles.accordionHeaderBtn}
                onClick={() => toggleSection(section.id)}
                aria-expanded={showOpen}
              >
                <div className={styles.accordionHeaderLeft}>
                  {section.iconName && (
                    <div className={styles.accordionIconCircle}>
                      <FontIcon iconName={section.iconName} className={styles.accordionSectionIcon} />
                    </div>
                  )}
                  <div className={styles.accordionHeaderText}>
                    <h3 className={styles.accordionTitle}>{section.title}</h3>
                    {section.description && (
                      <span className={styles.accordionSubtitle}>{section.description}</span>
                    )}
                  </div>
                </div>

                <div className={styles.accordionHeaderRight}>
                  {section.badge && (
                    <span className={styles.accordionBadge}>{section.badge}</span>
                  )}
                  <span className={styles.blockCountBadge}>
                    {section.blocks.length} {section.blocks.length === 1 ? 'item' : 'items'}
                  </span>
                  <div className={`${styles.chevronWrapper} ${showOpen ? styles.chevronRotated : ''}`}>
                    <FontIcon iconName="ChevronDown" />
                  </div>
                </div>
              </button>

              {showOpen && (
                <div className={styles.accordionContent}>
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
                      <p>No content items matching your search in this section.</p>
                    </div>
                  )}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
};
