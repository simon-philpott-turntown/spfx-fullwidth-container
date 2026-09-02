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
  SelectTabData,
  SelectTabEvent,
  TabValue,
  makeStyles,
  shorthands,
  tokens,
  Body1
} from '@fluentui/react-components';
import {
  BookRegular,
  MoneyRegular,
  AppsRegular,
  ShieldCheckmarkRegular,
  FolderRegular,
  InfoRegular,
  AddRegular,
  CheckmarkCircleRegular,
  ArrowTrendingLinesRegular,
  WrenchRegular,
  LockClosedRegular,
  GlobeRegular,
  DocumentRegular
} from '@fluentui/react-icons';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap(tokens.spacingVerticalL)
  },
  tabList: {
    ...shorthands.borderBottom('1px', 'solid', tokens.colorNeutralStroke2),
    paddingBottom: tokens.spacingVerticalXS,
    overflowX: 'auto',
    scrollbarWidth: 'thin'
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
  gridColumns?: number;
  gridRows?: number;
  isEditMode?: boolean;
  onUpdateBlock?: (sectionId: string, blockId: string, updatedFields: Partial<import('../models/IContainerModels').IContentBlock>) => void;
  onDeleteBlock?: (sectionId: string, blockId: string) => void;
  onAddBlock?: (sectionId: string) => void;
  onEditBlockProperties?: (sectionIndex: number, blockIndex: number) => void;
  onAddSection?: () => void;
  onUpdateSection?: (sectionId: string, updatedFields: Partial<IContainerSection>) => void;
}

function renderTabIcon(name?: string): React.ReactElement {
  switch (name) {
    case 'BookAnswers':
    case 'Book':
      return <BookRegular fontSize={18} />;
    case 'Financial':
    case 'Money':
      return <MoneyRegular fontSize={18} />;
    case 'AppIconDefault':
    case 'App':
    case 'Apps':
      return <AppsRegular fontSize={18} />;
    case 'ComplianceAudit':
    case 'Shield':
      return <ShieldCheckmarkRegular fontSize={18} />;
    case 'CheckList':
    case 'Checkmark':
      return <CheckmarkCircleRegular fontSize={18} />;
    case 'TimelineProgress':
    case 'Timeline':
      return <ArrowTrendingLinesRegular fontSize={18} />;
    case 'Calculator':
    case 'Wrench':
      return <WrenchRegular fontSize={18} />;
    case 'Lock':
      return <LockClosedRegular fontSize={18} />;
    case 'Globe':
      return <GlobeRegular fontSize={18} />;
    case 'DocumentManagement':
    case 'Document':
      return <DocumentRegular fontSize={18} />;
    default:
      return <FolderRegular fontSize={18} />;
  }
}

export const TabsContainer: React.FC<ITabsContainerProps> = ({
  sections,
  searchQuery,
  gridColumns,
  gridRows,
  isEditMode,
  onUpdateBlock,
  onDeleteBlock,
  onAddBlock,
  onEditBlockProperties
}) => {
  const styles = useStyles();
  const [selectedTab, setSelectedTab] = React.useState<TabValue>(() =>
    sections.length > 0 ? sections[0].id : ''
  );

  const handleTabSelect = (event: SelectTabEvent, data: SelectTabData): void => {
    setSelectedTab(data.value);
  };

  const activeSectionIndex = Math.max(
    0,
    sections.findIndex((s) => s.id === selectedTab)
  );
  const activeSection = sections[activeSectionIndex] || (sections.length > 0 ? sections[0] : undefined);

  if (!sections || sections.length === 0) {
    return (
      <div className={styles.emptyState}>
        <InfoRegular fontSize={24} />
        <Body1>No sections configured.</Body1>
      </div>
    );
  }

  // Filter blocks by search query
  const q = searchQuery ? searchQuery.toLowerCase().trim() : '';
  const blocks = (activeSection && Array.isArray(activeSection.blocks)) ? activeSection.blocks : [];
  const filteredBlocks = q
    ? blocks.filter(
        (b) =>
          (b.title && b.title.toLowerCase().includes(q)) ||
          (b.description && b.description.toLowerCase().includes(q)) ||
          (b.badge && b.badge.toLowerCase().includes(q)) ||
          (b.tags && Array.isArray(b.tags) && b.tags.some((t) => t.toLowerCase().includes(q)))
      )
    : blocks;

  const gridStyle: React.CSSProperties = {
    gridTemplateColumns:
      gridColumns && gridColumns > 0
        ? `repeat(${gridColumns}, 1fr)`
        : 'repeat(auto-fill, minmax(320px, 1fr))'
  };

  return (
    <div className={styles.container}>
      {/* Fluent 2 TabList */}
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
            icon={renderTabIcon(section.iconName)}
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

      {/* Cards Grid */}
      <div className={styles.grid} style={gridStyle}>
        {filteredBlocks.map((block, blkIdx) => (
          <BlockRenderer
            key={block.id}
            block={block}
            containerGridColumns={gridColumns}
            containerGridRows={gridRows}
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
