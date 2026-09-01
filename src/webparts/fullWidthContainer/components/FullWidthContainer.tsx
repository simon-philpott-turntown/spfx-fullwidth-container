/**
 * @file FullWidthContainer.tsx
 * @description Root container component for the SPFx Full-Width Container Web Part.
 */

import * as React from 'react';
import type { IFullWidthContainerProps } from './IFullWidthContainerProps';
import { LayoutMode } from '../models/IContainerModels';
import { TabsContainer } from './TabsContainer';
import { AccordionContainer } from './AccordionContainer';
import { FontIcon, initializeIcons } from '@fluentui/react';
import styles from './FullWidthContainer.module.scss.css';

// Ensure Fluent UI icons are registered
initializeIcons();

export const FullWidthContainer: React.FC<IFullWidthContainerProps> = (props) => {
  const {
    title,
    subtitle,
    layoutMode: initialLayoutMode,
    containerStyle,
    compactPadding,
    enableAnimation,
    showSearch,
    sections,
    isDarkTheme
  } = props;

  // Local state for interactive layout mode override and search
  const [currentMode, setCurrentMode] = React.useState<LayoutMode>(initialLayoutMode || 'tabs');
  const [searchQuery, setSearchQuery] = React.useState<string>('');

  // Sync prop changes from property pane
  React.useEffect(() => {
    if (initialLayoutMode) {
      setCurrentMode(initialLayoutMode);
    }
  }, [initialLayoutMode]);

  const styleVariantClass =
    containerStyle === 'gradient'
      ? styles.gradient
      : containerStyle === 'minimal'
      ? styles.minimal
      : styles.glassmorphism;

  const darkThemeClass = isDarkTheme ? styles.darkTheme : '';
  const paddingClass = compactPadding ? styles.compactPadding : '';

  return (
    <section
      className={`${styles.fullWidthContainer} ${styleVariantClass} ${darkThemeClass} ${paddingClass}`}
      role="region"
      aria-label={title || 'Content Container'}
    >
      {/* Top Header & Interactive Controls */}
      <div className={styles.containerHeader}>
        <div className={styles.headerTitles}>
          <h2 className={styles.mainTitle}>{title || 'Interactive Content Hub'}</h2>
          {subtitle && <p className={styles.subtitle}>{subtitle}</p>}
        </div>

        <div className={styles.headerActions}>
          {/* Search Bar */}
          {showSearch !== false && (
            <div className={styles.searchBox}>
              <FontIcon iconName="Search" className={styles.searchIcon} />
              <input
                type="text"
                placeholder="Filter content..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className={styles.searchInput}
                aria-label="Filter content"
              />
              {searchQuery && (
                <button
                  type="button"
                  onClick={() => setSearchQuery('')}
                  style={{ border: 'none', background: 'transparent', cursor: 'pointer', padding: 0 }}
                  aria-label="Clear filter"
                >
                  <FontIcon iconName="Clear" style={{ fontSize: 12, color: '#94a3b8' }} />
                </button>
              )}
            </div>
          )}

          {/* Mode Switcher */}
          <div className={styles.modeSwitcher} role="group" aria-label="Layout view mode">
            <button
              type="button"
              className={`${styles.modeButton} ${
                currentMode === 'tabs' ? styles.modeButtonActive : ''
              }`}
              onClick={() => setCurrentMode('tabs')}
              title="Switch to Tabbed View"
            >
              <FontIcon iconName="Tab" />
              <span>Tabs</span>
            </button>
            <button
              type="button"
              className={`${styles.modeButton} ${
                currentMode === 'accordion' ? styles.modeButtonActive : ''
              }`}
              onClick={() => setCurrentMode('accordion')}
              title="Switch to Accordion View"
            >
              <FontIcon iconName="GroupList" />
              <span>Accordion</span>
            </button>
          </div>
        </div>
      </div>

      {/* Main Container View Rendering */}
      {currentMode === 'tabs' ? (
        <TabsContainer
          sections={sections || []}
          enableAnimation={enableAnimation !== false}
          searchQuery={searchQuery}
        />
      ) : (
        <AccordionContainer
          sections={sections || []}
          enableAnimation={enableAnimation !== false}
          searchQuery={searchQuery}
        />
      )}
    </section>
  );
};

export default FullWidthContainer;
