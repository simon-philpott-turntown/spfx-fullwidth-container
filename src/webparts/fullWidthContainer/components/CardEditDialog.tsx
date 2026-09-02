/**
 * @file CardEditDialog.tsx
 * @description In-place Fluent UI 2 Modal Dialog for editing all card properties directly on canvas.
 * Integrates Fluent UI 2 Visual Icon Picker with search and Global Term Store taxonomy tags.
 * Follows the Microsoft Fluent 2 Design System.
 */

import * as React from 'react';
import {
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogContent,
  DialogActions,
  Button,
  Input,
  Textarea,
  Label,
  Dropdown,
  Option,
  makeStyles,
  shorthands,
  tokens
} from '@fluentui/react-components';
import {
  DismissRegular,
  SaveRegular,
  MoneyRegular,
  ShieldCheckmarkRegular,
  DocumentRegular,
  AppsRegular,
  WrenchRegular,
  GlobeRegular,
  ArrowTrendingLinesRegular,
  LockClosedRegular,
  CheckmarkCircleRegular,
  BookOpenRegular,
  FolderRegular,
  SparkleRegular,
  ChartMultipleRegular,
  ReceiptMoneyRegular,
  PeopleRegular,
  BuildingRegular,
  MegaphoneRegular,
  StarRegular
} from '@fluentui/react-icons';
import { IContentBlock, BlockType } from '../models/IContainerModels';
import { TermStorePicker } from './TermStorePicker';
import { FluentIconPicker } from './FluentIconPicker';

const useStyles = makeStyles({
  dialogSurface: {
    maxWidth: '580px',
    width: '92vw',
    maxHeight: '88vh',
    overflowY: 'auto',
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: tokens.shadow28
  },
  contentGrid: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap(tokens.spacingVerticalM)
  },
  fieldRow: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap(tokens.spacingVerticalXXS)
  },
  twoColRow: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    ...shorthands.gap(tokens.spacingHorizontalM)
  },
  iconPreviewBox: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    ...shorthands.padding(tokens.spacingVerticalXS, tokens.spacingHorizontalS),
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke1),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorNeutralBackground2
  },
  iconDisplay: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap(tokens.spacingHorizontalS),
    fontSize: '20px',
    color: tokens.colorBrandForeground1
  }
});

export interface ICardEditDialogProps {
  isOpen: boolean;
  block: IContentBlock | undefined;
  maxColumns?: number;
  maxRows?: number;
  onSave: (updatedBlock: IContentBlock) => void;
  onDismiss: () => void;
}

export const renderFluentIconPreview = (iconKey?: string): React.ReactElement => {
  switch (iconKey) {
    case 'Financial':
      return <MoneyRegular />;
    case 'ReceiptMoney':
      return <ReceiptMoneyRegular />;
    case 'TimelineProgress':
      return <ArrowTrendingLinesRegular />;
    case 'ChartMultiple':
      return <ChartMultipleRegular />;
    case 'ComplianceAudit':
      return <ShieldCheckmarkRegular />;
    case 'CheckList':
      return <CheckmarkCircleRegular />;
    case 'Calculator':
      return <WrenchRegular />;
    case 'AppIconDefault':
      return <AppsRegular />;
    case 'Lock':
      return <LockClosedRegular />;
    case 'Globe':
      return <GlobeRegular />;
    case 'Document':
      return <DocumentRegular />;
    case 'Folder':
      return <FolderRegular />;
    case 'Sparkle':
      return <SparkleRegular />;
    case 'People':
      return <PeopleRegular />;
    case 'Building':
      return <BuildingRegular />;
    case 'Megaphone':
      return <MegaphoneRegular />;
    case 'Star':
      return <StarRegular />;
    case 'BookAnswers':
    default:
      return <BookOpenRegular />;
  }
};

export const CardEditDialog: React.FC<ICardEditDialogProps> = ({
  isOpen,
  block,
  maxColumns = 4,
  maxRows = 5,
  onSave,
  onDismiss
}) => {
  const styles = useStyles();

  const [formData, setFormData] = React.useState<Partial<IContentBlock>>({});
  const [isIconPickerOpen, setIsIconPickerOpen] = React.useState<boolean>(false);

  const effectiveMaxCols = Math.max(1, maxColumns > 0 ? maxColumns : 4);
  const effectiveMaxRows = Math.max(1, maxRows > 0 ? maxRows : 5);

  React.useEffect(() => {
    if (block) {
      setFormData({
        ...block,
        colSpan: Math.min(block.colSpan || 1, effectiveMaxCols),
        rowSpan: Math.min(block.rowSpan || 1, effectiveMaxRows)
      });
    }
  }, [block, effectiveMaxCols, effectiveMaxRows]);

  if (!isOpen || !block) {
    return null;
  }

  const handleSave = (): void => {
    if (block && formData) {
      onSave({
        ...block,
        ...formData,
        colSpan: Math.min(formData.colSpan || 1, effectiveMaxCols),
        rowSpan: Math.min(formData.rowSpan || 1, effectiveMaxRows)
      } as IContentBlock);
    }
    onDismiss();
  };

  const colOptions = Array.from({ length: effectiveMaxCols }, (_, i) => i + 1);
  const rowOptions = Array.from({ length: effectiveMaxRows }, (_, i) => i + 1);

  return (
    <>
      <Dialog open={isOpen} onOpenChange={(e, data) => { if (!data.open) onDismiss(); }}>
        <DialogSurface className={styles.dialogSurface}>
          <DialogTitle
            action={
              <Button
                appearance="subtle"
                aria-label="close"
                icon={<DismissRegular />}
                onClick={onDismiss}
              />
            }
          >
            Edit Card & Metric Properties
          </DialogTitle>

          <DialogBody>
            <DialogContent className={styles.contentGrid}>
              {/* Block Type & Fluent 2 Icon Visual Selection */}
              <div className={styles.twoColRow}>
                <div className={styles.fieldRow}>
                  <Label required weight="semibold">Card type</Label>
                  <Dropdown
                    value={
                      formData.type === 'metric'
                        ? 'British metric stat (£)'
                        : formData.type === 'embed'
                        ? 'Embed or tool'
                        : 'Standard content card'
                    }
                    onOptionSelect={(e, data) => {
                      setFormData({ ...formData, type: data.optionValue as BlockType });
                    }}
                  >
                    <Option value="card">Standard content card</Option>
                    <Option value="metric">British metric stat (£)</Option>
                    <Option value="embed">Embed or tool</Option>
                  </Dropdown>
                </div>

                {/* Fluent UI 2 Icon Picker */}
                <div className={styles.fieldRow}>
                  <Label weight="semibold">Fluent UI 2 Icon</Label>
                  <div className={styles.iconPreviewBox}>
                    <div className={styles.iconDisplay}>
                      {renderFluentIconPreview(formData.iconName)}
                      <span style={{ fontSize: '0.85rem', color: tokens.colorNeutralForeground1, fontWeight: 500 }}>
                        {formData.iconName || 'BookAnswers'}
                      </span>
                    </div>
                    <Button
                      size="small"
                      appearance="outline"
                      onClick={() => setIsIconPickerOpen(true)}
                    >
                      Browse Icons...
                    </Button>
                  </div>
                </div>
              </div>

              {/* Title */}
              <div className={styles.fieldRow}>
                <Label required weight="semibold">Title</Label>
                <Input
                  value={formData.title || ''}
                  placeholder="Card title"
                  onChange={(e, data) => setFormData({ ...formData, title: data.value })}
                />
              </div>

              {/* Description */}
              <div className={styles.fieldRow}>
                <Label weight="semibold">Description or body text</Label>
                <Textarea
                  rows={3}
                  value={formData.description || ''}
                  placeholder="Card summary"
                  onChange={(e, data) => setFormData({ ...formData, description: data.value })}
                />
              </div>

              {/* Metric Fields (if type === 'metric') */}
              {formData.type === 'metric' && (
                <div className={styles.twoColRow}>
                  <div className={styles.fieldRow}>
                    <Label weight="semibold">Metric value (GBP £)</Label>
                    <Input
                      value={formData.metricValue || ''}
                      placeholder="£1,420,000"
                      onChange={(e, data) => setFormData({ ...formData, metricValue: data.value })}
                    />
                  </div>
                  <div className={styles.fieldRow}>
                    <Label weight="semibold">Trend badge</Label>
                    <Input
                      value={formData.metricTrend || ''}
                      placeholder="+10%"
                      onChange={(e, data) => setFormData({ ...formData, metricTrend: data.value })}
                    />
                  </div>
                </div>
              )}

              {/* Embed URL (if type === 'embed') */}
              {formData.type === 'embed' && (
                <div className={styles.fieldRow}>
                  <Label weight="semibold">iFrame embed URL</Label>
                  <Input
                    value={formData.embedUrl || ''}
                    placeholder="https://..."
                    onChange={(e, data) => setFormData({ ...formData, embedUrl: data.value })}
                  />
                </div>
              )}

              {/* Card Layout & Grid Sizing (Dynamic based on container columns/rows) */}
              <div className={styles.twoColRow}>
                <div className={styles.fieldRow}>
                  <Label weight="semibold">Column span (Max {effectiveMaxCols})</Label>
                  <Dropdown
                    value={
                      formData.colSpan === 1
                        ? '1 column (standard)'
                        : `Span ${formData.colSpan || 1} columns`
                    }
                    onOptionSelect={(e, data) => {
                      setFormData({ ...formData, colSpan: Number(data.optionValue) || 1 });
                    }}
                  >
                    {colOptions.map((c) => (
                      <Option key={c} value={String(c)}>
                        {c === 1 ? '1 column (standard)' : c === effectiveMaxCols ? `Span ${c} columns (full width)` : `Span ${c} columns`}
                      </Option>
                    ))}
                  </Dropdown>
                </div>

                <div className={styles.fieldRow}>
                  <Label weight="semibold">Row span (Max {effectiveMaxRows})</Label>
                  <Dropdown
                    value={
                      formData.rowSpan === 1
                        ? '1 row (standard)'
                        : `Span ${formData.rowSpan || 1} rows`
                    }
                    onOptionSelect={(e, data) => {
                      setFormData({ ...formData, rowSpan: Number(data.optionValue) || 1 });
                    }}
                  >
                    {rowOptions.map((r) => (
                      <Option key={r} value={String(r)}>
                        {r === 1 ? '1 row (standard)' : `Span ${r} rows`}
                      </Option>
                    ))}
                  </Dropdown>
                </div>
              </div>

              {/* Height Mode Behavior */}
              <div className={styles.fieldRow}>
                <Label weight="semibold">Card height behavior</Label>
                <Dropdown
                  value={
                    formData.heightMode === 'equal'
                      ? 'Equal row height (match tallest)'
                      : formData.heightMode === 'auto'
                      ? 'Fit content height (independent)'
                      : 'Inherit container setting'
                  }
                  onOptionSelect={(e, data) => {
                    setFormData({ ...formData, heightMode: data.optionValue as 'default' | 'auto' | 'equal' });
                  }}
                >
                  <Option value="default">Inherit container setting</Option>
                  <Option value="auto">Fit content height (independent)</Option>
                  <Option value="equal">Equal row height (match tallest)</Option>
                </Dropdown>
              </div>

              {/* Global Term Store Taxonomy Tags */}
              <div className={styles.fieldRow}>
                <Label weight="semibold">Global term store • Intranet taxonomy tags</Label>
                <TermStorePicker
                  selectedTags={formData.termStoreTags || []}
                  onChange={(tags) => setFormData({ ...formData, termStoreTags: tags })}
                  isEditMode={true}
                />
              </div>

              {/* Live Data API Configuration */}
              <div className={styles.fieldRow}>
                <Label weight="semibold">Real-time live data API (optional)</Label>
                <div
                  style={{
                    padding: '10px',
                    borderRadius: tokens.borderRadiusMedium,
                    backgroundColor: tokens.colorNeutralBackground2,
                    border: `1px solid ${tokens.colorNeutralStroke2}`,
                    display: 'flex',
                    flexDirection: 'column',
                    gap: '8px'
                  }}
                >
                  <Input
                    size="small"
                    placeholder="API endpoint URL (e.g. https://... or demo-api/burnDown)"
                    value={formData.liveDataConfig?.apiUrl || ''}
                    onChange={(e, data) => {
                      setFormData({
                        ...formData,
                        liveDataConfig: {
                          apiUrl: data.value,
                          jsonPath: formData.liveDataConfig?.jsonPath || 'value',
                          prefix: formData.liveDataConfig?.prefix || '£',
                          refreshIntervalSeconds: formData.liveDataConfig?.refreshIntervalSeconds || 30
                        }
                      });
                    }}
                  />
                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px' }}>
                    <Input
                      size="small"
                      placeholder="JSON path (e.g. data.metric)"
                      value={formData.liveDataConfig?.jsonPath || ''}
                      onChange={(e, data) => {
                        if (formData.liveDataConfig) {
                          setFormData({
                            ...formData,
                            liveDataConfig: { ...formData.liveDataConfig, jsonPath: data.value }
                          });
                        }
                      }}
                    />
                    <Input
                      size="small"
                      placeholder="Prefix (e.g. £)"
                      value={formData.liveDataConfig?.prefix || ''}
                      onChange={(e, data) => {
                        if (formData.liveDataConfig) {
                          setFormData({
                            ...formData,
                            liveDataConfig: { ...formData.liveDataConfig, prefix: data.value }
                          });
                        }
                      }}
                    />
                  </div>
                </div>
              </div>

              {/* Badge & Tags */}
              <div className={styles.twoColRow}>
                <div className={styles.fieldRow}>
                  <Label weight="semibold">Status or category badge</Label>
                  <Input
                    value={formData.badge || ''}
                    placeholder="Badge"
                    onChange={(e, data) => setFormData({ ...formData, badge: data.value })}
                  />
                </div>
                <div className={styles.fieldRow}>
                  <Label weight="semibold">Additional tags (comma-separated)</Label>
                  <Input
                    value={formData.tags ? formData.tags.join(', ') : ''}
                    placeholder="Tag 1, Tag 2"
                    onChange={(e, data) => {
                      const splitTags = data.value.split(',').map((t) => t.trim()).filter((t) => !!t);
                      setFormData({ ...formData, tags: splitTags });
                    }}
                  />
                </div>
              </div>

              {/* Link & Action */}
              <div className={styles.twoColRow}>
                <div className={styles.fieldRow}>
                  <Label weight="semibold">Button or link text</Label>
                  <Input
                    value={formData.linkText || ''}
                    placeholder="Action link"
                    onChange={(e, data) => setFormData({ ...formData, linkText: data.value })}
                  />
                </div>
                <div className={styles.fieldRow}>
                  <Label weight="semibold">Destination URL</Label>
                  <Input
                    value={formData.linkUrl || ''}
                    placeholder="https://..."
                    onChange={(e, data) => setFormData({ ...formData, linkUrl: data.value })}
                  />
                </div>
              </div>
            </DialogContent>

            <DialogActions>
              <Button appearance="secondary" onClick={onDismiss}>
                Cancel
              </Button>
              <Button appearance="primary" icon={<SaveRegular />} onClick={handleSave}>
                Save Card
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      {/* Fluent UI 2 Visual Icon Picker Modal */}
      <FluentIconPicker
        isOpen={isIconPickerOpen}
        selectedIconKey={formData.iconName || 'BookAnswers'}
        onSelectIcon={(iconKey) => setFormData({ ...formData, iconName: iconKey })}
        onDismiss={() => setIsIconPickerOpen(false)}
      />
    </>
  );
};
