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
import { RichTextEditable } from './RichTextEditable';
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
  Caption1,
  makeStyles,
  shorthands,
  tokens,
  mergeClasses
} from '@fluentui/react-components';
import {
  SearchRegular,
  TabRegular,
  ListRegular
} from '@fluentui/react-icons';

const useStyles = makeStyles({
  rootBase: {
    width: '100%',
    boxSizing: 'border-box',
    color: tokens.colorNeutralForeground1,
    minHeight: '120px',
    transitionProperty: 'background-color, border-color, box-shadow',
    transitionDuration: '250ms'
  },
  rootStandard: {
    backgroundColor: tokens.colorNeutralBackground2,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2)
  },
  rootGlassmorphism: {
    backgroundColor: 'rgba(255, 255, 255, 0.65)',
    backdropFilter: 'blur(20px) saturate(180%)',
    ...shorthands.border('1px', 'solid', 'rgba(255, 255, 255, 0.5)'),
    boxShadow: tokens.shadow16
  },
  rootBranded: {
    backgroundColor: tokens.colorBrandBackground2,
    ...shorthands.border('1px', 'solid', tokens.colorBrandStroke2)
  },
  rootMinimal: {
    backgroundColor: 'transparent',
    ...shorthands.border('none')
  },
  rootGradient: {
    backgroundImage: 'linear-gradient(135deg, rgba(0, 120, 212, 0.08) 0%, rgba(255, 255, 255, 0.95) 100%)',
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2)
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
  },
  inlineTitleInput: {
    backgroundColor: 'transparent',
    ...shorthands.border('1px', 'dashed', 'transparent'),
    ...shorthands.borderRadius(tokens.borderRadiusSmall),
    paddingTop: '2px',
    paddingBottom: '2px',
    paddingLeft: '4px',
    paddingRight: '4px',
    fontFamily: 'inherit',
    fontSize: '1.75rem',
    fontWeight: 700,
    lineHeight: '2.25rem',
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
  },
  inlineSubtitleInput: {
    backgroundColor: 'transparent',
    ...shorthands.border('1px', 'dashed', 'transparent'),
    ...shorthands.borderRadius(tokens.borderRadiusSmall),
    paddingTop: '2px',
    paddingBottom: '2px',
    paddingLeft: '4px',
    paddingRight: '4px',
    fontFamily: 'inherit',
    fontSize: '0.95rem',
    color: tokens.colorNeutralForeground3,
    outline: 'none',
    width: '100%',
    boxSizing: 'border-box',
    marginTop: '2px',
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

import { FloatingTextToolbar } from './FloatingTextToolbar';

export const FullWidthContainer: React.FC<IFullWidthContainerProps> = (props) => {
  const {
    title,
    subtitle,
    layoutMode: initialLayoutMode,
    containerStyle = 'standard',
    compactPadding,
    showSearch,
    gridColumns,
    gridRows,
    cardHeightMode,
    webPartBackgroundColor,
    sections,
    isDarkTheme,
    spfxTheme,
    isEditMode,
    onOpenPropertyPane,
    onTitleChange,
    onSubtitleChange,
    onUpdateSection,
    onAddSection,
    onDeleteSection,
    onUpdateBlock,
    onAddBlock,
    onDeleteBlock,
    onEditBlockProperties
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

  const rootStyleClass = React.useMemo(() => {
    switch (containerStyle) {
      case 'glassmorphism':
        return styles.rootGlassmorphism;
      case 'branded':
        return styles.rootBranded;
      case 'minimal':
        return styles.rootMinimal;
      case 'gradient':
        return styles.rootGradient;
      case 'standard':
      default:
        return styles.rootStandard;
    }
  }, [containerStyle, styles]);

  const rootCustomStyle: React.CSSProperties = React.useMemo(() => {
    return webPartBackgroundColor ? { backgroundColor: webPartBackgroundColor } : {};
  }, [webPartBackgroundColor]);

  return (
    <FluentProvider theme={fluentTheme} className={mergeClasses(styles.rootBase, rootStyleClass)} style={rootCustomStyle}>
      <div className={mergeClasses(styles.inner, compactPadding ? styles.innerCompact : undefined)}>
        {/* Visual Edit Mode Indicator */}
        {isEditMode && (
          <div className={styles.editBanner}>
            <div className={styles.editBannerLeft}>
              <Badge appearance="filled" color="brand" size="small">SharePoint authoring mode</Badge>
              <Caption1 style={{ color: tokens.colorNeutralForeground2 }}>
                Full-width container active. Click on titles to edit or hover between items to add content.
              </Caption1>
            </div>
            {onOpenPropertyPane && (
              <Button
                appearance="subtle"
                size="small"
                onClick={onOpenPropertyPane}
                style={{ border: `1px solid ${tokens.colorBrandStroke1}` }}
              >
                Configure web part
              </Button>
            )}
          </div>
        )}

        {/* Header with Title & Controls */}
        <div className={styles.header}>
          <div className={styles.headerTextCol} style={{ flex: 1, minWidth: '280px' }}>
            <RichTextEditable
              tag="h1"
              html={title || ''}
              isEditMode={!!isEditMode}
              placeholder="Container title"
              onChange={(newTitle) => onTitleChange && onTitleChange(newTitle)}
              style={{
                fontSize: '1.75rem',
                fontWeight: 700,
                lineHeight: '2.25rem',
                color: 'inherit'
              }}
            />

            <RichTextEditable
              tag="p"
              html={subtitle || ''}
              isEditMode={!!isEditMode}
              placeholder="Container subtitle or description"
              onChange={(newSub) => onSubtitleChange && onSubtitleChange(newSub)}
              style={{
                fontSize: '0.95rem',
                color: tokens.colorNeutralForeground3,
                marginTop: '4px'
              }}
            />
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
            gridColumns={gridColumns}
            gridRows={gridRows}
            cardHeightMode={cardHeightMode}
            isEditMode={isEditMode}
            onUpdateBlock={onUpdateBlock}
            onDeleteBlock={onDeleteBlock}
            onAddBlock={onAddBlock}
            onEditBlockProperties={onEditBlockProperties}
            onAddSection={onAddSection}
            onUpdateSection={onUpdateSection}
            onDeleteSection={onDeleteSection}
          />
        ) : (
          <AccordionContainer
            sections={sections}
            searchQuery={searchQuery}
            gridColumns={gridColumns}
            gridRows={gridRows}
            cardHeightMode={cardHeightMode}
            isEditMode={isEditMode}
            onUpdateBlock={onUpdateBlock}
            onDeleteBlock={onDeleteBlock}
            onAddBlock={onAddBlock}
            onEditBlockProperties={onEditBlockProperties}
            onAddSection={onAddSection}
            onUpdateSection={onUpdateSection}
            onDeleteSection={onDeleteSection}
          />
        )}
      </div>
    </FluentProvider>
  );
};
