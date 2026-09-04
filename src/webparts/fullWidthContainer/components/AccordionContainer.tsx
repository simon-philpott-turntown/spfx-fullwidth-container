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
  Caption1
} from '@fluentui/react-components';
import {
  ChevronDoubleDownRegular,
  ChevronDoubleUpRegular,
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
}

import { renderUnifiedIcon } from './CustomSvgIconRegistry';

function renderSectionIcon(section?: IContainerSection): React.ReactElement {
  const iconKey = section?.iconName || 'BookAnswers';
  const iconColor = section?.iconColor;
  const showBg = section?.showIconBackground;
  const bg = section?.iconBackgroundColor || tokens.colorBrandBackground2;

  const iconElement = renderUnifiedIcon(iconKey, iconColor, '20px');

  if (showBg) {
    return (
      <span
        style={{
          display: 'inline-flex',
          alignItems: 'center',
          justifyContent: 'center',
          width: '32px',
          height: '32px',
          borderRadius: '6px',
          backgroundColor: bg,
          color: iconColor || tokens.colorBrandForeground2,
          marginRight: '6px',
          flexShrink: 0
        }}
      >
        {iconElement}
      </span>
    );
  }

  return iconElement;
}

export const AccordionContainer: React.FC<IAccordionContainerProps> = ({
  sections,
  searchQuery,
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
  onDeleteSection
}) => {
  const styles = useStyles();
  const [openItems, setOpenItems] = React.useState<string[]>(() =>
    sections.length > 0 ? [sections[0].id] : []
  );
  const [isSectionDialogOpen, setIsSectionDialogOpen] = React.useState<boolean>(false);
  const [editingSection, setEditingSection] = React.useState<IContainerSection | undefined>(undefined);

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

  const gridStyle: React.CSSProperties = {
    gridTemplateColumns:
      gridColumns && gridColumns > 0
        ? `repeat(${gridColumns}, 1fr)`
        : 'repeat(auto-fill, minmax(320px, 1fr))'
  };

  return (
    <div className={styles.container}>
      {/* Global Toolbar */}
      <div className={styles.toolbar}>
        <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
          {sections.length} {sections.length === 1 ? 'section' : 'sections'}
        </Caption1>
        <div className={styles.toolbarActions}>
          <Button
            appearance="subtle"
            size="small"
            icon={<ChevronDoubleDownRegular />}
            onClick={expandAll}
          >
            Expand all
          </Button>
          <Button
            appearance="subtle"
            size="small"
            icon={<ChevronDoubleUpRegular />}
            onClick={collapseAll}
          >
            Collapse all
          </Button>
          {isEditMode && onAddSection && (
            <Button
              appearance="subtle"
              size="small"
              icon={<AddRegular />}
              onClick={onAddSection}
              title="Add new section"
            >
              Add section
            </Button>
          )}
        </div>
      </div>

      {/* Section Edit Side Dialog */}
      <SectionEditDialog
        isOpen={isSectionDialogOpen}
        section={editingSection}
        canDelete={sections.length > 1}
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
            ? blocks.filter((b) => {
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

                return titleMatch || descMatch || badgeMatch || metricMatch || trendMatch || tagMatch || termStoreMatch || innerItemsMatch;
              })
            : blocks;

          return (
            <AccordionItem
              key={section.id}
              value={section.id}
              className={styles.accordionItem}
            >
              <AccordionHeader>
                <div className={styles.headerContent}>
                  {renderSectionIcon(section)}
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
                    {isEditMode && (
                      <div
                        style={{ display: 'flex', alignItems: 'center', gap: '2px', marginLeft: '6px' }}
                        onClick={(e) => e.stopPropagation()}
                      >
                        <Button
                          size="small"
                          appearance="subtle"
                          icon={<EditRegular />}
                          title={`Edit section properties (${section.title})`}
                          onClick={(e) => {
                            e.stopPropagation();
                            setEditingSection(section);
                            setIsSectionDialogOpen(true);
                          }}
                        />
                        {sections.length > 1 && onDeleteSection && (
                          <Button
                            size="small"
                            appearance="subtle"
                            icon={<DeleteRegular />}
                            title={`Delete section (${section.title})`}
                            onClick={(e) => {
                              e.stopPropagation();
                              if (window.confirm(`Are you sure you want to delete section "${section.title}"?`)) {
                                onDeleteSection(section.id);
                              }
                            }}
                          />
                        )}
                      </div>
                    )}
                  </div>
                </div>
              </AccordionHeader>

              <AccordionPanel style={section.backgroundColor ? { backgroundColor: section.backgroundColor, padding: '12px', borderRadius: '8px' } : undefined}>
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
                      <Body1 style={{ fontWeight: 600 }}>Add new card</Body1>
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
