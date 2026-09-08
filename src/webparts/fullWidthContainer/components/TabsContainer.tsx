/**
 * @file TabsContainer.tsx
 * @description Tabbed navigation layout using Fluent UI 2 TabList and Tab components.
 * Follows the Microsoft Fluent 2 Design System (https://fluent2.microsoft.design/).
 */

import * as React from 'react';
import { IContainerSection } from '../models/IContainerModels';
import { BlockRenderer } from './BlockRenderer';
import {
  TabList,
  Tab,
  Badge,
  Button,
  SelectTabData,
  SelectTabEvent,
  TabValue,
  makeStyles,
  shorthands,
  tokens,
  Body1
} from '@fluentui/react-components';
import {
  InfoRegular,
  AddRegular,
  EditRegular,
  DeleteRegular
} from '@fluentui/react-icons';
import { SectionEditDialog } from './SectionEditDialog';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap(tokens.spacingVerticalL)
  },
  tabHeaderRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    ...shorthands.borderBottom('1px', 'solid', tokens.colorNeutralStroke2),
    paddingBottom: tokens.spacingVerticalXS,
    gap: '12px'
  },
  tabList: {
    overflowX: 'auto',
    scrollbarWidth: 'thin',
    flex: 1
  },
  tabActionsRow: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap(tokens.spacingHorizontalXXS),
    flexShrink: 0
  },
  grid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fill, minmax(320px, 1fr))',
    ...shorthands.gap(tokens.spacingHorizontalL),
    width: '100%'
  },
  addCardButton: {
    minHeight: '140px',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    ...shorthands.border('2px', 'dashed', tokens.colorBrandStroke2),
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorBrandForeground1,
    cursor: 'pointer',
    ...shorthands.gap(tokens.spacingVerticalXS),
    transitionProperty: 'border-color, background-color',
    transitionDuration: '150ms',
    ':hover': {
      ...shorthands.borderColor(tokens.colorBrandBackground),
      backgroundColor: tokens.colorBrandBackground2
    }
  },
  emptyState: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    ...shorthands.padding(tokens.spacingVerticalXXL),
    color: tokens.colorNeutralForeground3,
    ...shorthands.gap(tokens.spacingVerticalS)
  },
  badge: {
    whiteSpace: 'nowrap',
    flexShrink: 0,
    display: 'inline-flex',
    alignItems: 'center',
    marginLeft: tokens.spacingHorizontalXS
  }
});

export interface ITabsContainerProps {
  sections: IContainerSection[];
  searchQuery: string;
  selectedFilterTerms?: Record<string, string>;
  gridColumns?: number;
  gridRows?: number;
  cardHeightMode?: 'auto' | 'equal';
  isEditMode?: boolean;
  onUpdateBlock?: (sectionId: string, blockId: string, updatedFields: Partial<import('../models/IContainerModels').IContentBlock>) => void;
  onDeleteBlock?: (sectionId: string, blockId: string) => void;
  onAddBlock?: (sectionId: string) => void;
  onEditBlockProperties?: (sectionIndex: number, blockIndex: number) => void;
  onAddSection?: () => void;
  onUpdateSection?: (sectionId: string, updatedFields: Partial<IContainerSection>) => void;
  onDeleteSection?: (sectionId: string) => void;
  assetPickerService?: import('../services/IAssetPickerService').IAssetPickerService;
}


import { renderUnifiedIcon } from './CustomSvgIconRegistry';

function renderTabIcon(section?: IContainerSection): React.ReactElement {
  const iconKey = section?.iconName || 'BookAnswers';
  const iconColor = section?.iconColor;
  const showBg = section?.showIconBackground;
  const bg = section?.iconBackgroundColor || tokens.colorBrandBackground2;

  const iconElement = renderUnifiedIcon(iconKey, iconColor, '18px');

  if (showBg) {
    return (
      <span
        style={{
          display: 'inline-flex',
          alignItems: 'center',
          justifyContent: 'center',
          width: '26px',
          height: '26px',
          borderRadius: '6px',
          backgroundColor: bg,
          color: iconColor || tokens.colorBrandForeground2,
          marginRight: '2px'
        }}
      >
        {iconElement}
      </span>
    );
  }

  return iconElement;
}

export const TabsContainer: React.FC<ITabsContainerProps> = ({
  sections,
  searchQuery,
  selectedFilterTerms,
  gridColumns,
  gridRows,
  cardHeightMode = 'auto',
  isEditMode,
  onUpdateBlock,
  onDeleteBlock,
  onAddBlock,
  onEditBlockProperties,
  onAddSection,
  onUpdateSection,
  onDeleteSection,
  assetPickerService
}) => {
  const styles = useStyles();
  const [selectedTab, setSelectedTab] = React.useState<TabValue>(() =>
    sections.length > 0 ? sections[0].id : ''
  );
  const [isSectionDialogOpen, setIsSectionDialogOpen] = React.useState<boolean>(false);
  const [editingSection, setEditingSection] = React.useState<IContainerSection | undefined>(undefined);

  const handleTabSelect = (event: SelectTabEvent, data: SelectTabData): void => {
    setSelectedTab(data.value);
  };

  const activeSectionIndex = Math.max(
    0,
    sections.findIndex((s) => s.id === selectedTab)
  );
  const activeSection = sections[activeSectionIndex] || (sections.length > 0 ? sections[0] : undefined);

  const [activeItemFilter, setActiveItemFilter] = React.useState<string>('');

  React.useEffect(() => {
    const handleCardFilterApply = (e: Event): void => {
      const ce = e as CustomEvent<{ sourceItemId: string; filterValue: string }>;
      setActiveItemFilter(ce.detail?.filterValue || '');
    };
    window.addEventListener('dashboard:card-filter-apply', handleCardFilterApply);
    return () => {
      window.removeEventListener('dashboard:card-filter-apply', handleCardFilterApply);
    };
  }, []);

  if (!sections || sections.length === 0) {
    return (
      <div className={styles.emptyState}>
        <InfoRegular fontSize={24} />
        <Body1>No sections configured.</Body1>
      </div>
    );
  }

  // Active term filters from dropdowns (filterId -> term label)
  const activeDropdownFilters: string[] = selectedFilterTerms
    ? Object.keys(selectedFilterTerms)
        .map((key) => selectedFilterTerms[key])
        .filter((val): val is string => !!val && val.trim().length > 0)
    : [];

  // Filter blocks by search query AND active term store dropdown filters AND interactive item filter (buttons/process)
  const q = searchQuery ? searchQuery.toLowerCase().trim() : '';
  const itemFilt = activeItemFilter ? activeItemFilter.toLowerCase().trim() : '';
  const blocks = (activeSection && Array.isArray(activeSection.blocks)) ? activeSection.blocks : [];

  const filteredBlocks = blocks.filter((b) => {
    // If the card itself is the host of the filter buttons or process model, keep it visible so author can interact
    const isFilterHost = b.items && b.items.some((it) => it.type === 'filterButtons' || it.type === 'processModel');

    // 1. Interactive Button / Process Stage Filter Matching
    if (itemFilt && !isFilterHost) {
      const titleMatch = b.title ? b.title.toLowerCase().includes(itemFilt) : false;
      const descMatch = b.description ? b.description.toLowerCase().includes(itemFilt) : false;
      const badgeMatch = b.badge ? b.badge.toLowerCase().includes(itemFilt) : false;
      const tagMatch = b.tags && Array.isArray(b.tags) && b.tags.some((t) => t.toLowerCase().includes(itemFilt));
      const termStoreMatch = b.termStoreTags && Array.isArray(b.termStoreTags) && b.termStoreTags.some((t) => {
        const labelMatch = t.label ? t.label.toLowerCase().includes(itemFilt) : false;
        const setMatch = t.termSetName ? t.termSetName.toLowerCase().includes(itemFilt) : false;
        return labelMatch || setMatch;
      });
      const innerItemsMatch = b.items && Array.isArray(b.items) && b.items.some((item) => {
        const textMatch = item.text ? item.text.toLowerCase().includes(itemFilt) : false;
        const ctaMatch = item.ctaHeading ? item.ctaHeading.toLowerCase().includes(itemFilt) : false;
        const btnMatch = item.buttonLabel ? item.buttonLabel.toLowerCase().includes(itemFilt) : false;
        return textMatch || ctaMatch || btnMatch;
      });

      const matchesItemFilter = titleMatch || descMatch || badgeMatch || tagMatch || termStoreMatch || innerItemsMatch;
      if (!matchesItemFilter) return false;
    }

    // 2. Text Search Query Matching
    if (q) {
      const titleMatch = b.title ? b.title.toLowerCase().includes(q) : false;
      const descMatch = b.description ? b.description.toLowerCase().includes(q) : false;
      const badgeMatch = b.badge ? b.badge.toLowerCase().includes(q) : false;
      const metricMatch = b.metricValue ? b.metricValue.toLowerCase().includes(q) : false;
      const trendMatch = b.metricTrend ? b.metricTrend.toLowerCase().includes(q) : false;
      const tagMatch = b.tags && Array.isArray(b.tags) && b.tags.some((t) => t.toLowerCase().includes(q));
      const termStoreMatch = b.termStoreTags && Array.isArray(b.termStoreTags) && b.termStoreTags.some((t) => {
        const labelMatch = t.label ? t.label.toLowerCase().includes(q) : false;
        const setMatch = t.termSetName ? t.termSetName.toLowerCase().includes(q) : false;
        const pathMatch = t.path ? t.path.toLowerCase().includes(q) : false;
        return labelMatch || setMatch || pathMatch;
      });
      const innerItemsMatch = b.items && Array.isArray(b.items) && b.items.some((item) => {
        const textMatch = item.text ? item.text.toLowerCase().includes(q) : false;
        const ctaMatch = item.ctaHeading ? item.ctaHeading.toLowerCase().includes(q) : false;
        const btnMatch = item.buttonLabel ? item.buttonLabel.toLowerCase().includes(q) : false;
        const innerTermMatch = item.termStoreTags && Array.isArray(item.termStoreTags) && item.termStoreTags.some((t) => {
          return (t.label && t.label.toLowerCase().includes(q)) || (t.termSetName && t.termSetName.toLowerCase().includes(q));
        });
        return textMatch || ctaMatch || btnMatch || innerTermMatch;
      });

      const matchesSearch = titleMatch || descMatch || badgeMatch || metricMatch || trendMatch || tagMatch || termStoreMatch || innerItemsMatch;
      if (!matchesSearch) return false;
    }

    // 3. Global Term Store Dropdown Filters Matching (conjunction / AND match)
    if (activeDropdownFilters.length > 0) {
      // Collect all tags on this block and its items
      const blockTags = b.termStoreTags || [];
      const itemTags: import('../models/IContainerModels').ITermStoreTag[] = (b.items || []).reduce<import('../models/IContainerModels').ITermStoreTag[]>(
        (acc, it) => acc.concat(it.termStoreTags || []),
        []
      );
      const allCardTags = blockTags.concat(itemTags);

      // Each active dropdown term must be matched on the card
      const matchesAllDropdowns = activeDropdownFilters.every((reqTerm: string) => {
        const target = reqTerm.toLowerCase().trim();
        return allCardTags.some((tag) => {
          return (
            (tag.label && tag.label.toLowerCase() === target) ||
            (tag.path && tag.path.toLowerCase().includes(target))
          );
        });
      });

      if (!matchesAllDropdowns) return false;
    }

    return true;
  });


  const gridStyle: React.CSSProperties = {
    gridTemplateColumns:
      gridColumns && gridColumns > 0
        ? `repeat(${gridColumns}, 1fr)`
        : 'repeat(auto-fill, minmax(320px, 1fr))',
    ...(activeSection?.backgroundColor
      ? {
          backgroundColor: activeSection.backgroundColor,
          padding: '16px',
          borderRadius: '8px'
        }
      : {}),
    ...(activeSection?.backgroundImage
      ? {
          backgroundImage: `url("${activeSection.backgroundImage}")`,
          backgroundSize: 'cover',
          backgroundPosition: 'center',
          backgroundRepeat: 'no-repeat',
          padding: '16px',
          borderRadius: '8px'
        }
      : {})
  };

  return (
    <div className={styles.container}>
      {/* Header Row: Fluent 2 TabList with Edit/Delete Section Icons */}
      <div className={styles.tabHeaderRow}>
        <TabList
          selectedValue={selectedTab}
          onTabSelect={handleTabSelect}
          size="medium"
          appearance="subtle"
          className={styles.tabList}
        >
          {sections.map((section) => (
            <Tab
              key={section.id}
              value={section.id}
              icon={renderTabIcon(section)}
            >
              {section.title}
              {section.badge && (
                <Badge appearance="filled" size="small" className={styles.badge}>
                  {section.badge}
                </Badge>
              )}
            </Tab>
          ))}
        </TabList>

        {/* Section Edit / Delete action icons in authoring mode */}
        {isEditMode && activeSection && (
          <div className={styles.tabActionsRow}>
            <Button
              size="small"
              appearance="subtle"
              icon={<EditRegular />}
              title={`Edit active section properties (${activeSection.title})`}
              onClick={(e) => {
                e.stopPropagation();
                setEditingSection(activeSection);
                setIsSectionDialogOpen(true);
              }}
            >
              Edit section
            </Button>

            {sections.length > 1 && onDeleteSection && (
              <Button
                size="small"
                appearance="subtle"
                icon={<DeleteRegular />}
                title={`Delete active section (${activeSection.title})`}
                onClick={(e) => {
                  e.stopPropagation();
                  if (window.confirm(`Are you sure you want to delete section "${activeSection.title}"?`)) {
                    onDeleteSection(activeSection.id);
                  }
                }}
              />
            )}

            {onAddSection && (
              <Button
                size="small"
                appearance="subtle"
                icon={<AddRegular />}
                title="Add new section"
                onClick={(e) => {
                  e.stopPropagation();
                  onAddSection();
                }}
              />
            )}
          </div>
        )}
      </div>

      {/* Section Edit Side Dialog */}
      <SectionEditDialog
        isOpen={isSectionDialogOpen}
        section={editingSection}
        canDelete={sections.length > 1}
        assetPickerService={assetPickerService}
        onSave={(updated) => {
          if (editingSection && onUpdateSection) {
            onUpdateSection(editingSection.id, updated);
          }
        }}
        onDelete={() => {
          if (editingSection && onDeleteSection) {
            onDeleteSection(editingSection.id);
          }
        }}
        onDismiss={() => {
          setIsSectionDialogOpen(false);
          setEditingSection(undefined);
        }}
      />

      {/* Cards Grid */}
      <div className={styles.grid} style={gridStyle}>
        {filteredBlocks.map((block, blkIdx) => (
          <BlockRenderer
            key={block.id}
            block={block}
            containerGridColumns={gridColumns}
            containerGridRows={gridRows}
            containerCardHeightMode={cardHeightMode}
            isEditMode={isEditMode}
            onUpdate={(fields) => {
              if (activeSection && onUpdateBlock) {
                onUpdateBlock(activeSection.id, block.id, fields);
              }
            }}
            onDelete={() => {
              if (activeSection && onDeleteBlock) {
                onDeleteBlock(activeSection.id, block.id);
              }
            }}
            onEditProperties={() => {
              if (onEditBlockProperties) {
                onEditBlockProperties(activeSectionIndex, blkIdx);
              }
            }}
            assetPickerService={assetPickerService}
          />
        ))}

        {/* Quick Add Card Action in Edit Mode */}
        {isEditMode && activeSection && onAddBlock && (
          <div
            className={styles.addCardButton}
            onClick={() => onAddBlock(activeSection.id)}
            title="Add a new card to this section"
          >
            <AddRegular fontSize={24} />
            <Body1 style={{ fontWeight: 600 }}>Add new card</Body1>
          </div>
        )}
      </div>

      {filteredBlocks.length === 0 && !isEditMode && (
        <div className={styles.emptyState}>
          <InfoRegular fontSize={24} />
          <Body1>No content items found matching &quot;{searchQuery}&quot; in this section.</Body1>
        </div>
      )}
    </div>
  );
};
