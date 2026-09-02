/**
 * @file CardEditDialog.tsx
 * @description In-place Fluent UI 2 Modal Dialog for editing all card properties directly on canvas.
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
  CheckmarkCircleRegular
} from '@fluentui/react-icons';
import { IContentBlock, BlockType } from '../models/IContainerModels';
import { TermStorePicker } from './TermStorePicker';

const useStyles = makeStyles({
  dialogSurface: {
    maxWidth: '560px',
    width: '90vw',
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
    ...shorthands.gap(tokens.spacingHorizontalS),
    ...shorthands.padding(tokens.spacingVerticalXS, tokens.spacingHorizontalS),
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorNeutralBackground2
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

const AVAILABLE_ICONS = [
  { key: 'BookAnswers', label: 'Book or documentation', icon: <DocumentRegular /> },
  { key: 'Financial', label: 'Financial or pound (£)', icon: <MoneyRegular /> },
  { key: 'ComplianceAudit', label: 'Shield or governance', icon: <ShieldCheckmarkRegular /> },
  { key: 'CheckList', label: 'Checklist or verified', icon: <CheckmarkCircleRegular /> },
  { key: 'TimelineProgress', label: 'Trending lines or metrics', icon: <ArrowTrendingLinesRegular /> },
  { key: 'Calculator', label: 'Wrench or engineering', icon: <WrenchRegular /> },
  { key: 'AppIconDefault', label: 'Apps or tool grid', icon: <AppsRegular /> },
  { key: 'Lock', label: 'Lock or security', icon: <LockClosedRegular /> },
  { key: 'Globe', label: 'Globe or portal', icon: <GlobeRegular /> }
];

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
          Edit card properties
        </DialogTitle>

        <DialogBody>
          <DialogContent className={styles.contentGrid}>
            {/* Block Type */}
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

              {/* Icon Selection */}
              <div className={styles.fieldRow}>
                <Label weight="semibold">Icon</Label>
                <Dropdown
                  value={
                    AVAILABLE_ICONS.find((i) => i.key === formData.iconName)?.label || 'Globe or portal'
                  }
                  onOptionSelect={(e, data) => {
                    setFormData({ ...formData, iconName: data.optionValue });
                  }}
                >
                  {AVAILABLE_ICONS.map((item) => (
                    <Option key={item.key} value={item.key}>
                      {item.label}
                    </Option>
                  ))}
                </Dropdown>
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

            {/* Global Term Store Tags */}
            <div className={styles.fieldRow}>
              <Label weight="semibold">Global term store tags</Label>
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
              Save card
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};
