/**
 * @file FullWidthContainer.tsx
 * @description Root container component for the SPFx Full-Width Container Web Part.
 * Built entirely using Microsoft Fluent UI 2 (https://fluent2.microsoft.design/).
 */

import * as React from 'react';
import type { IFullWidthContainerProps } from './IFullWidthContainerProps';
import { LayoutMode } from '../models/IContainerModels';
import { TabsContainer } from './TabsContainer';
import { AccordionContainer } from './AccordionContainer';
import { getFluent2Theme } from '../utils/themeBridge';
import {
  FluentProvider,
  Input,
  TabList,
  Tab,
  Title1,
  Subtitle2,
  Badge,
  Button,
  Text,
  makeStyles,
  shorthands,
  tokens
} from '@fluentui/react-components';
import {
  SearchRegular,
  TabRegular,
  ListRegular
} from '@fluentui/react-icons';

const useStyles = makeStyles({
  root: {
    width: '100%',
    boxSizing: 'border-box',
    backgroundColor: tokens.colorNeutralBackground2,
    color: tokens.colorNeutralForeground1,
    minHeight: '120px'
  },
  inner: {
    width: '100%',
    boxSizing: 'border-box',
    ...shorthands.padding(tokens.spacingVerticalXL, tokens.spacingHorizontalXXL),
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap(tokens.spacingVerticalL)
  },
  innerCompact: {
    ...shorthands.padding(tokens.spacingVerticalM, tokens.spacingHorizontalL)
  },
  header: {
    display: 'flex',
    flexDirection: 'column',
    justifyContent: 'space-between',
    alignItems: 'flex-start',
    ...shorthands.gap(tokens.spacingVerticalM),
    '@media (min-width: 768px)': {
      flexDirection: 'row',
      alignItems: 'flex-end'
    }
  },
  headerTextCol: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap(tokens.spacingVerticalXXS)
  },
  headerControls: {
    display: 'flex',
    alignItems: 'center',
    flexWrap: 'wrap',
    ...shorthands.gap(tokens.spacingHorizontalM)
  },
  searchInput: {
    minWidth: '240px'
  },
  modeSwitcher: {
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorNeutralBackground3
  },
  editBanner: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    flexWrap: 'wrap',
    ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalM),
    backgroundColor: tokens.colorBrandBackground2,
    ...shorthands.border('1px', 'dashed', tokens.colorBrandStroke1),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    ...shorthands.gap(tokens.spacingHorizontalM)
  },
  editBannerLeft: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap(tokens.spacingHorizontalS)
  }
});

export const FullWidthContainer: React.FC<IFullWidthContainerProps> = (props) => {
  const {
    title,
    subtitle,
    layoutMode: initialLayoutMode,
    compactPadding,
    showSearch,
    sections,
    isDarkTheme,
    spfxTheme,
    isEditMode,
    onOpenPropertyPane
  } = props;

  const styles = useStyles();
  const [layoutMode, setLayoutMode] = React.useState<LayoutMode>(initialLayoutMode || 'tabs');
  const [searchQuery, setSearchQuery] = React.useState<string>('');

  // Sync state if property pane changes
  React.useEffect(() => {
    if (initialLayoutMode) {
      setLayoutMode(initialLayoutMode);
    }
  }, [initialLayoutMode]);

  // Compute dynamic Fluent 2 theme inheriting SharePoint Online palette
  const fluentTheme = React.useMemo(() => {
    return getFluent2Theme(spfxTheme, isDarkTheme);
  }, [spfxTheme, isDarkTheme]);

  return (
    <FluentProvider theme={fluentTheme} className={styles.root}>
      <div className={`${styles.inner} ${compactPadding ? styles.innerCompact : ''}`}>
        {/* Edit Mode Helper Notification Bar */}
        {isEditMode && (
          <div className={styles.editBanner}>
            <div className={styles.editBannerLeft}>
              <Badge appearance="filled" color="brand">SharePoint Edit Mode</Badge>
              <Caption1 style={{ color: tokens.colorNeutralForeground2 }}>
                Customise content sections, badges, British metrics (£), and layout in the properties sidebar.
              </Caption1>
            </div>
            {onOpenPropertyPane && (
              <Button
                size="small"
                appearance="outline"
                onClick={onOpenPropertyPane}
              >
                Open Properties Sidebar
              </Button>
            )}
          </div>
        )}

        {/* Header with Title & Controls */}
        <div className={styles.header}>
          <div className={styles.headerTextCol}>
            <Title1>{title || 'Interactive Content Hub'}</Title1>
            {subtitle && (
              <Subtitle2 style={{ color: tokens.colorNeutralForeground3 }}>
                {subtitle}
              </Subtitle2>
            )}
          </div>

          <div className={styles.headerControls}>
            {/* Search Filter Bar */}
            {showSearch !== false && (
              <Input
                className={styles.searchInput}
                contentBefore={<SearchRegular />}
                placeholder="Filter items, tags, GBP..."
                value={searchQuery}
                onChange={(e, data) => setSearchQuery(data.value)}
                size="medium"
              />
            )}

            {/* Layout Mode Switcher */}
            <TabList
              selectedValue={layoutMode}
              onTabSelect={(e, data) => setLayoutMode(data.value as LayoutMode)}
              size="small"
              appearance="subtle"
              className={styles.modeSwitcher}
            >
              <Tab value="tabs" icon={<TabRegular />}>
                Tabs
              </Tab>
              <Tab value="accordion" icon={<ListRegular />}>
                Accordion
              </Tab>
            </TabList>
          </div>
        </div>

        {/* Content Container Body */}
        {layoutMode === 'tabs' ? (
          <TabsContainer
            sections={sections}
            searchQuery={searchQuery}
          />
        ) : (
          <AccordionContainer
            sections={sections}
            searchQuery={searchQuery}
          />
        )}
      </div>
    </FluentProvider>
  );
};
