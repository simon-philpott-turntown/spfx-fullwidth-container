/**
 * @file AccordionContainer.tsx
 * @description Collapsible accordion layout using Fluent UI 2 Accordion components.
 * Follows the Microsoft Fluent 2 Design System (https://fluent2.microsoft.design/).
 */

import * as React from 'react';
import { IContainerSection } from '../models/IContainerModels';
import { BlockRenderer } from './BlockRenderer';
import {
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
  AccordionToggleEventHandler,
  AccordionToggleData,
  AccordionToggleEvent,
  Badge,
  Button,
  makeStyles,
  shorthands,
  tokens,
  Body1,
  Caption1,
  Text
} from '@fluentui/react-components';
import {
  ChevronDoubleDownRegular,
  ChevronDoubleUpRegular,
  InfoRegular,
  FolderRegular,
  BookRegular,
  MoneyRegular,
  AppsRegular,
  ShieldCheckmarkRegular,
  AddRegular
} from '@fluentui/react-icons';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap(tokens.spacingVerticalM)
  },
  toolbar: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    paddingBottom: tokens.spacingVerticalS,
    ...shorthands.borderBottom('1px', 'solid', tokens.colorNeutralStroke2)
  },
  toolbarActions: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap(tokens.spacingHorizontalS)
  },
  accordionItem: {
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke1),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorNeutralBackground1,
    marginBottom: tokens.spacingVerticalS,
    overflow: 'hidden'
  },
  headerContent: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap(tokens.spacingHorizontalM),
    width: '100%'
  },
  headerTextCol: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'flex-start'
  },
  headerBadges: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap(tokens.spacingHorizontalXS),
    marginLeft: 'auto',
    marginRight: tokens.spacingHorizontalM,
    flexShrink: 0
  },
  badge: {
    whiteSpace: 'nowrap',
    flexShrink: 0,
    display: 'inline-flex',
    alignItems: 'center'
  },
  grid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fill, minmax(320px, 1fr))',
    ...shorthands.gap(tokens.spacingHorizontalL),
    ...shorthands.padding(tokens.spacingVerticalM, '0')
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
  }
});

export interface IAccordionContainerProps {
  sections: IContainerSection[];
  searchQuery: string;
  isEditMode?: boolean;
  onUpdateBlock?: (sectionId: string, blockId: string, updatedFields: Partial<import('../models/IContainerModels').IContentBlock>) => void;
  onDeleteBlock?: (sectionId: string, blockId: string) => void;
  onAddBlock?: (sectionId: string) => void;
  onEditBlockProperties?: (sectionIndex: number, blockIndex: number) => void;
  onAddSection?: () => void;
  onUpdateSection?: (sectionId: string, updatedFields: Partial<IContainerSection>) => void;
}

function renderSectionIcon(name?: string): React.ReactElement {
  switch (name) {
    case 'BookAnswers':
      return <BookRegular fontSize={20} />;
    case 'Financial':
    case 'Money':
      return <MoneyRegular fontSize={20} />;
    case 'AppIconDefault':
      return <AppsRegular fontSize={20} />;
    case 'ComplianceAudit':
    case 'Shield':
      return <ShieldCheckmarkRegular fontSize={20} />;
    default:
      return <FolderRegular fontSize={20} />;
  }
}

export const AccordionContainer: React.FC<IAccordionContainerProps> = ({
  sections,
  searchQuery,
  isEditMode,
  onUpdateBlock,
  onDeleteBlock,
  onAddBlock,
  onEditBlockProperties
}) => {
  const styles = useStyles();
  const [openItems, setOpenItems] = React.useState<string[]>(() =>
    sections.length > 0 ? [sections[0].id] : []
  );

  const handleToggle: AccordionToggleEventHandler<string> = (
    event: AccordionToggleEvent,
    data: AccordionToggleData<string>
  ) => {
    setOpenItems(data.openItems);
  };

  const expandAll = (): void => {
    setOpenItems(sections.map((s) => s.id));
  };

  const collapseAll = (): void => {
    setOpenItems([]);
  };

  if (!sections || sections.length === 0) {
    return (
      <div className={styles.emptyState}>
        <InfoRegular fontSize={24} />
        <Body1>No accordion sections configured.</Body1>
      </div>
    );
  }

  const q = searchQuery ? searchQuery.toLowerCase().trim() : '';

  return (
    <div className={styles.container}>
      {/* Global Toolbar */}
      <div className={styles.toolbar}>
        <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
          {sections.length} {sections.length === 1 ? 'Section' : 'Sections'}
        </Caption1>
        <div className={styles.toolbarActions}>
          <Button
            appearance="subtle"
            size="small"
            icon={<ChevronDoubleDownRegular />}
            onClick={expandAll}
          >
            Expand All
          </Button>
          <Button
            appearance="subtle"
            size="small"
            icon={<ChevronDoubleUpRegular />}
            onClick={collapseAll}
          >
            Collapse All
          </Button>
        </div>
      </div>

      {/* Fluent 2 Accordion */}
      <Accordion
        collapsible
        multiple
        openItems={openItems}
        onToggle={handleToggle}
      >
        {sections.map((section, secIdx) => {
          const blocks = (section && Array.isArray(section.blocks)) ? section.blocks : [];
          const filteredBlocks = q
            ? blocks.filter(
                (b) =>
                  (b.title && b.title.toLowerCase().includes(q)) ||
                  (b.description && b.description.toLowerCase().includes(q)) ||
                  (b.badge && b.badge.toLowerCase().includes(q)) ||
                  (b.tags && Array.isArray(b.tags) && b.tags.some((t) => t.toLowerCase().includes(q)))
              )
            : blocks;

          return (
            <AccordionItem
              key={section.id}
              value={section.id}
              className={styles.accordionItem}
            >
              <AccordionHeader>
                <div className={styles.headerContent}>
                  {renderSectionIcon(section.iconName)}
                  <div className={styles.headerTextCol}>
                    <Body1 style={{ fontWeight: 600 }}>{section.title}</Body1>
                    {section.description && (
                      <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                        {section.description}
                      </Caption1>
                    )}
                  </div>
                  <div className={styles.headerBadges}>
                    {section.badge && (
                      <Badge appearance="tint" color="brand" size="small" className={styles.badge}>
                        {section.badge}
                      </Badge>
                    )}
                    <Badge appearance="outline" size="small" className={styles.badge}>
                      {(section.blocks ? section.blocks.length : 0)} {(section.blocks && section.blocks.length === 1) ? 'item' : 'items'}
                    </Badge>
                  </div>
                </div>
              </AccordionHeader>

              <AccordionPanel>
                <div className={styles.grid}>
                  {filteredBlocks.map((block, blkIdx) => (
                    <BlockRenderer
                      key={block.id}
                      block={block}
                      isEditMode={isEditMode}
                      onUpdate={(fields) => {
                        if (onUpdateBlock) {
                          onUpdateBlock(section.id, block.id, fields);
                        }
                      }}
                      onDelete={() => {
                        if (onDeleteBlock) {
                          onDeleteBlock(section.id, block.id);
                        }
                      }}
                      onEditProperties={() => {
                        if (onEditBlockProperties) {
                          onEditBlockProperties(secIdx, blkIdx);
                        }
                      }}
                    />
                  ))}

                  {/* Add Card Button in Accordion */}
                  {isEditMode && onAddBlock && (
                    <div
                      className={styles.addCardButton}
                      onClick={() => onAddBlock(section.id)}
                      title="Add a new card to this section"
                    >
                      <AddRegular fontSize={24} />
                      <Body1 style={{ fontWeight: 600 }}>Add New Card</Body1>
                    </div>
                  )}
                </div>

                {filteredBlocks.length === 0 && !isEditMode && (
                  <div className={styles.emptyState}>
                    <Caption1>No items matching &quot;{searchQuery}&quot; in this section.</Caption1>
                  </div>
                )}
              </AccordionPanel>
            </AccordionItem>
          );
        })}
      </Accordion>
    </div>
  );
};
