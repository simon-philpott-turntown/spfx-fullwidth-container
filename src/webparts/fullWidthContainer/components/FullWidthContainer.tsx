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
  Avatar,
  Caption1,
  Popover,
  PopoverSurface,
  PopoverTrigger,
  makeStyles,
  shorthands,
  tokens,
  mergeClasses
} from '@fluentui/react-components';
import {
  SearchRegular,
  TabRegular,
  ListRegular,
  EditRegular,
  PersonRegular,
  ShieldCheckmarkRegular
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
    ...shorthands.border('none')
  },
  rootGlassmorphism: {
    backgroundColor: 'rgba(255, 255, 255, 0.65)',
    backdropFilter: 'blur(20px) saturate(180%)',
    ...shorthands.border('1px', 'solid', 'rgba(255, 255, 255, 0.5)'),
    boxShadow: tokens.shadow16
  },
  rootBranded: {
    backgroundColor: tokens.colorBrandBackground2,
    ...shorthands.border('none')
  },
  rootMinimal: {
    backgroundColor: 'transparent',
    ...shorthands.border('none')
  },
  rootGradient: {
    backgroundImage: 'linear-gradient(135deg, rgba(0, 120, 212, 0.08) 0%, rgba(255, 255, 255, 0.95) 100%)',
    ...shorthands.border('none')
  },
  inner: {
    width: '100%',
    boxSizing: 'border-box',
    ...shorthands.padding(tokens.spacingVerticalXL, tokens.spacingHorizontalXXL),
    display: 'flex',
    flexDirection: 'column',
    // ...shorthands.gap(tokens.spacingVerticalL) // [USER_TEST: comment out vertical spacing in flex containers]
  },
  innerCompact: {
    ...shorthands.padding(tokens.spacingVerticalM, tokens.spacingHorizontalL)
  },
  header: {
    display: 'flex',
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'flex-start',
    width: '100%',
    gap: '16px'
  },
  headerTextCol: {
    display: 'flex',
    flexDirection: 'column',
    // ...shorthands.gap(tokens.spacingVerticalXXS), // [USER_TEST: comment out vertical spacing in flex containers]
    flex: 1
  },
  headerTopRight: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    flexShrink: 0
  },
  searchAndFiltersRow: {
    display: 'flex',
    alignItems: 'center',
    flexWrap: 'wrap',
    gap: tokens.spacingHorizontalM,
    width: '100%'
  },
  searchInput: {
    minWidth: '260px'
  },
  headerSearchInput: {
    minWidth: '220px',
    maxWidth: '320px',
    height: '32px',
    boxSizing: 'border-box',
    '& input': {
      paddingTop: '6px',
      paddingBottom: '6px',
      fontSize: '0.85rem'
    }
  },
  modeSwitcher: {
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorNeutralBackground3,
    height: '32px',
    display: 'flex',
    alignItems: 'center'
  },
  editBanner: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    flexWrap: 'wrap',
    ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalM),
    backgroundColor: tokens.colorNeutralBackground3,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
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
import { ComposableContentSection } from './ComposableContentSection';

export const FullWidthContainer: React.FC<IFullWidthContainerProps> = (props) => {
  const {
    title,
    subtitle,
    layoutMode: initialLayoutMode,
    containerStyle = 'standard',
    compactPadding,
    showSearch,
    searchAlignment = 'left',
    searchPlaceholder,
    onSearchPlaceholderChange,
    gridColumns,
    gridRows,
    cardHeightMode,
    webPartBackgroundColor,
    sections,
    headerContentItems = [],
    onUpdateHeaderContentItems,
    termFilters = [],
    onUpdateTermFilters,
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
    onEditBlockProperties,
    onSaveBackupToLibrary,
    onRestoreFromLibrary,
    lastBackupMessage,
    assetPickerService,
    userProfileDetails,
    userProfilePhotoUrl
  } = props;

  const styles = useStyles();
  const [layoutMode, setLayoutMode] = React.useState<LayoutMode>(initialLayoutMode || 'tabs');
  const [searchQuery, setSearchQuery] = React.useState<string>('');
  const [selectedFilterTerms, setSelectedFilterTerms] = React.useState<Record<string, string>>({});
  const [isSavingSnapshot, setIsSavingSnapshot] = React.useState<boolean>(false);
  const [searchPlaceholderLocal, setSearchPlaceholderLocal] = React.useState<string>(searchPlaceholder || 'Filter items, tags, GBP...');
  const [isEditingSearchPlaceholder, setIsEditingSearchPlaceholder] = React.useState<boolean>(false);


  // Sync state if property pane changes
  React.useEffect(() => {
    if (initialLayoutMode) {
      setLayoutMode(initialLayoutMode);
    }
  }, [initialLayoutMode]);

  // Sync search placeholder from property pane
  React.useEffect(() => {
    setSearchPlaceholderLocal(searchPlaceholder || 'Filter items, tags, GBP...');
  }, [searchPlaceholder]);

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

  const handleQuickSave = async (folderType: 'Backups' | 'Templates'): Promise<void> => {
    if (!onSaveBackupToLibrary) return;
    try {
      setIsSavingSnapshot(true);
      await onSaveBackupToLibrary(folderType);
    } finally {
      setIsSavingSnapshot(false);
    }
  };

  return (
    <FluentProvider theme={fluentTheme} className={mergeClasses(styles.rootBase, rootStyleClass)} style={rootCustomStyle}>
      <div className={mergeClasses(styles.inner, compactPadding ? styles.innerCompact : undefined)}>
        {/* Visual Edit Mode Indicator */}
        {isEditMode && (
          <div className={styles.editBanner}>
            <div className={styles.editBannerLeft}>
              <Badge appearance="filled" color="brand" size="small">Dashboard edit mode</Badge>
              <Caption1 style={{ color: tokens.colorNeutralForeground2 }}>
                Click on titles to edit or hover between items to add content to the cards.
              </Caption1>
              {lastBackupMessage && (
                <Caption1 style={{ fontWeight: 600, color: lastBackupMessage.startsWith('✓') ? tokens.colorPaletteGreenForeground1 : tokens.colorPaletteRedForeground1, marginLeft: '8px' }}>
                  {lastBackupMessage}
                </Caption1>
              )}
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              {userProfileDetails && (
                <Popover positioning="below-end">
                  <PopoverTrigger disableButtonEnhancement>
                    <button
                      type="button"
                      title={`Signed in as ${userProfileDetails.displayName || 'User'}${userProfileDetails.jobTitle ? ` (${userProfileDetails.jobTitle})` : ''}`}
                      style={{
                        background: 'transparent',
                        border: 'none',
                        padding: 0,
                        margin: 0,
                        cursor: 'pointer',
                        display: 'inline-flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        borderRadius: '50%',
                        outline: 'none'
                      }}
                    >
                      <Avatar
                        size={28}
                        name={userProfileDetails.displayName || 'User'}
                        image={userProfilePhotoUrl ? { src: userProfilePhotoUrl } : undefined}
                        badge={{
                          status: userProfileDetails.isSiteAdmin ? 'available' : 'do-not-disturb'
                        }}
                      />
                    </button>
                  </PopoverTrigger>
                  <PopoverSurface
                    style={{
                      zIndex: 1000000,
                      boxShadow: tokens.shadow28,
                      padding: '16px',
                      maxWidth: '440px',
                      minWidth: '320px',
                      display: 'flex',
                      flexDirection: 'column',
                      gap: '10px',
                      borderRadius: tokens.borderRadiusMedium
                    }}
                  >
                    <div style={{ display: 'flex', alignItems: 'center', gap: '10px', borderBottom: `1px solid ${tokens.colorNeutralStroke2}`, paddingBottom: '10px' }}>
                      <Avatar
                        size={40}
                        name={userProfileDetails.displayName || 'User'}
                        image={userProfilePhotoUrl ? { src: userProfilePhotoUrl } : undefined}
                      />
                      <div style={{ display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                          <Caption1 style={{ fontWeight: 700, fontSize: '1rem', lineHeight: '1.25rem', color: tokens.colorNeutralForeground1 }}>
                            {userProfileDetails.displayName || 'User'}
                          </Caption1>
                          {userProfileDetails.isSiteAdmin && (
                            <Badge
                              appearance="filled"
                              color="success"
                              size="medium"
                              icon={<ShieldCheckmarkRegular style={{ fontSize: '14px' }} />}
                              style={{
                                fontSize: '0.75rem',
                                fontWeight: 600,
                                padding: '2px 8px',
                                height: '22px',
                                display: 'inline-flex',
                                alignItems: 'center',
                                gap: '4px'
                              }}
                            >
                              Admin
                            </Badge>
                          )}
                        </div>
                        {userProfileDetails.jobTitle && (
                          <Caption1 style={{ color: tokens.colorNeutralForeground3, fontSize: '0.78rem' }}>
                            {userProfileDetails.jobTitle}
                          </Caption1>
                        )}
                        {userProfileDetails.department && (
                          <Caption1 style={{ color: tokens.colorNeutralForeground3, fontSize: '0.75rem', fontWeight: 600 }}>
                            {userProfileDetails.department}
                          </Caption1>
                        )}
                        {userProfileDetails.officeLocation && (
                          <Caption1 style={{ color: tokens.colorNeutralForeground4, fontSize: '0.72rem' }}>
                            {userProfileDetails.officeLocation}
                          </Caption1>
                        )}
                      </div>
                    </div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                      {(() => {
                        const orderedKeys: Array<{ key: string; label: string }> = [
                          { key: 'PreferredName', label: 'PreferredName' },
                          { key: 'FirstName', label: 'FirstName' },
                          { key: 'email', label: 'email' },
                          { key: 'isSiteAdmin', label: 'isSiteAdmin' },
                          { key: 'isAnonymousGuestUser', label: 'isAnonymousGuestUser' },
                          { key: 'isExternalGuestUser', label: 'isExternalGuestUser' },
                          { key: 'givenName', label: 'givenName' },
                          { key: 'jobTitle', label: 'jobTitle' },
                          { key: 'department', label: 'department' },
                          { key: 'companyName', label: 'companyName' },
                          { key: 'employeeId', label: 'employeeId' },
                          { key: 'manager', label: 'manager' },
                          { key: 'streetAddress', label: 'streetAddress' },
                          { key: 'city', label: 'city' },
                          { key: 'state', label: 'state' },
                          { key: 'country', label: 'country' },
                          { key: 'extensionAttribute10', label: 'extensionAttribute10 (Country)' },
                          { key: 'extensionAttribute12', label: 'extensionAttribute12 (Start Date)' }
                        ];

                        return orderedKeys.map(({ key, label }) => {
                          const val = userProfileDetails[key];
                          if (val === undefined || val === null || val === '') {
                            return null;
                          }

                          let renderedVal: string;
                          if (typeof val === 'boolean') {
                            renderedVal = val ? 'Yes' : 'No';
                          } else if (typeof val === 'object') {
                            renderedVal = JSON.stringify(val);
                          } else {
                            renderedVal = String(val);
                          }

                          return (
                            <div key={key} style={{ display: 'flex', justifyContent: 'space-between', gap: '12px', fontSize: '0.78rem' }}>
                              <span style={{ fontWeight: 600, color: tokens.colorNeutralForeground3, whiteSpace: 'nowrap' }}>{label}:</span>
                              <span style={{ color: tokens.colorNeutralForeground1, wordBreak: 'break-all', textAlign: 'right' }}>
                                {renderedVal}
                              </span>
                            </div>
                          );
                        });
                      })()}
                    </div>
                  </PopoverSurface>
                </Popover>
              )}
              {onSaveBackupToLibrary && (
                <Button
                  appearance="primary"
                  size="small"
                  disabled={isSavingSnapshot}
                  onClick={() => void handleQuickSave('Backups')}
                >
                  {isSavingSnapshot ? 'Saving snapshot...' : '💾 Save snapshot to Library'}
                </Button>
              )}
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
          </div>
        )}

        {/* Header: Title on Left, Mode Switcher Permanently Pinned to Top Right */}
        <div className={styles.header}>
          <div className={styles.headerTextCol}>
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

          {/* Search Input + Layout Mode Switcher — permanently pinned top-right */}
          <div className={styles.headerTopRight}>
            {/* Search input is always visible in top-right */}
            {showSearch !== false && (
              <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                <Input
                  className={styles.headerSearchInput}
                  contentBefore={<SearchRegular style={{ fontSize: '16px' }} />}
                  placeholder={searchPlaceholderLocal}
                  value={searchQuery}
                  onChange={(e, data) => setSearchQuery(data.value)}
                  size="medium"
                />
                {isEditMode && (
                  <Popover
                    open={isEditingSearchPlaceholder}
                    positioning="below-end"
                    inline
                    onOpenChange={(_, data) => {
                      if (!data.open && onSearchPlaceholderChange) {
                        onSearchPlaceholderChange(searchPlaceholderLocal);
                      }
                      setIsEditingSearchPlaceholder(data.open);
                    }}
                  >
                    <PopoverTrigger disableButtonEnhancement>
                      <Button
                        size="small"
                        appearance="subtle"
                        icon={<EditRegular />}
                        title="Edit search prompt text"
                        onMouseDown={(e) => e.preventDefault()}
                      />
                    </PopoverTrigger>
                    <PopoverSurface
                      style={{
                        zIndex: 1000,
                        boxShadow: tokens.shadow16,
                        padding: '12px',
                        display: 'flex',
                        flexDirection: 'column',
                        gap: '8px',
                        minWidth: '280px',
                        borderRadius: tokens.borderRadiusMedium
                      }}
                    >
                      <Caption1 style={{ fontWeight: 600 }}>Search prompt text</Caption1>
                      <Input
                        size="medium"
                        value={searchPlaceholderLocal}
                        placeholder="Filter items, tags, GBP..."
                        onChange={(e, data) => setSearchPlaceholderLocal(data.value)}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') {
                            if (onSearchPlaceholderChange) {
                              onSearchPlaceholderChange(searchPlaceholderLocal);
                            }
                            setIsEditingSearchPlaceholder(false);
                          }
                          if (e.key === 'Escape') {
                            setSearchPlaceholderLocal(searchPlaceholder || 'Filter items, tags, GBP...');
                            setIsEditingSearchPlaceholder(false);
                          }
                        }}
                        // eslint-disable-next-line jsx-a11y/no-autofocus
                        autoFocus
                      />
                      <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                        Press Enter to apply · Escape to cancel
                      </Caption1>
                    </PopoverSurface>
                  </Popover>
                )}
              </div>
            )}

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

        {/* Web Part Top Content Section (Buttons, Process Model, Content blocks before sections) */}
        <ComposableContentSection
          items={headerContentItems}
          isEditMode={isEditMode}
          contextTitle="Add content to dashboard header"
          assetPickerService={assetPickerService}
          onUpdateItems={(newItems) => {
            if (onUpdateHeaderContentItems) {
              onUpdateHeaderContentItems(newItems);
            }
          }}
        />


        {/* Content Container Body */}
        {layoutMode === 'tabs' ? (
          <TabsContainer
            sections={sections}
            searchQuery={searchQuery}
            selectedFilterTerms={selectedFilterTerms}
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
            assetPickerService={assetPickerService}
          />
        ) : (
          <AccordionContainer
            sections={sections}
            searchQuery={searchQuery}
            selectedFilterTerms={selectedFilterTerms}
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
            assetPickerService={assetPickerService}
          />
        )}

      </div>
    </FluentProvider>
  );
};
