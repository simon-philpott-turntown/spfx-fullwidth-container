/**
 * @file CardEditDialog.tsx
 * @description In-place Fluent UI 2 Right-Hand Side Property Panel for editing all card properties.
 * Slides in from the right matching the native SharePoint Property Pane architecture (Screenshot 2).
 * Integrates Fluent UI 2 Visual Icon Picker with search and Global Term Store taxonomy tags.
 */

import * as React from 'react';
import {
  Button,
  Input,
  Textarea,
  Label,
  Dropdown,
  Option,
  makeStyles,
  shorthands,
  tokens,
  Subtitle2,
  Caption1,
  Checkbox,
  Divider
} from '@fluentui/react-components';
import { renderUnifiedIcon } from './CustomSvgIconRegistry';
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
import { BrandColorPickerPopover } from './BrandColorPickerPopover';

const useStyles = makeStyles({
  backdrop: {
    position: 'fixed',
    top: 0,
    left: 0,
    width: '100vw',
    height: '100vh',
    backgroundColor: 'transparent',
    zIndex: 999,
    display: 'flex',
    justifyContent: 'flex-end',
    pointerEvents: 'auto'
  },
  sidePanel: {
    position: 'relative',
    maxWidth: '92vw',
    height: '100vh',
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.borderLeft('1px', 'solid', tokens.colorNeutralStroke1),
    boxShadow: tokens.shadow28,
    display: 'flex',
    flexDirection: 'column',
    boxSizing: 'border-box',
    pointerEvents: 'auto'
  },
  leftResizeHandle: {
    position: 'absolute',
    left: '-4px',
    top: 0,
    bottom: 0,
    width: '8px',
    cursor: 'ew-resize',
    zIndex: 100,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    ':hover': {
      backgroundColor: tokens.colorBrandStroke1
    }
  },
  panelHeader: {
    padding: '16px 18px 12px 18px',
    ...shorthands.borderBottom('1px', 'solid', tokens.colorNeutralStroke2),
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'flex-start'
  },
  panelBody: {
    flex: 1,
    padding: '16px 18px',
    overflowY: 'auto',
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap(tokens.spacingVerticalM)
  },
  panelFooter: {
    padding: '12px 18px',
    ...shorthands.borderTop('1px', 'solid', tokens.colorNeutralStroke2),
    backgroundColor: tokens.colorNeutralBackground2,
    display: 'flex',
    justifyContent: 'flex-end',
    ...shorthands.gap('8px')
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
  },
  colorGrid: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '6px',
    alignItems: 'center',
    marginTop: '4px'
  },
  colorSwatch: {
    width: '24px',
    height: '24px',
    borderRadius: '4px',
    border: '1px solid #d1d1d1',
    cursor: 'pointer',
    transition: 'transform 0.1s ease',
    ':hover': {
      transform: 'scale(1.15)',
      boxShadow: tokens.shadow4
    }
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

export const renderFluentIconPreview = (iconKey?: string, iconColor?: string): JSX.Element => {
  return renderUnifiedIcon(iconKey, iconColor);
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

  // Local form state
  const [formData, setFormData] = React.useState<Partial<IContentBlock>>({});
  const [tagsInput, setTagsInput] = React.useState<string>('');
  const [isIconPickerOpen, setIsIconPickerOpen] = React.useState<boolean>(false);
  const [panelWidth, setPanelWidth] = React.useState<number>(400);

  const startResizeDrag = (e: React.MouseEvent): void => {
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX;
    const startWidth = panelWidth;

    const onMouseMove = (moveEvent: MouseEvent): void => {
      const deltaX = startX - moveEvent.clientX;
      const newWidth = Math.min(850, Math.max(340, startWidth + deltaX));
      setPanelWidth(newWidth);
    };

    const onMouseUp = (): void => {
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', onMouseUp);
    };

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
  };

  // Sync state when block changes
  React.useEffect(() => {
    if (block) {
      setFormData({
        ...block,
        colSpan: block.colSpan || 1,
        rowSpan: block.rowSpan || 1,
        heightMode: block.heightMode || 'default',
        type: block.type || 'card'
      });
      setTagsInput(block.tags ? block.tags.join(', ') : '');
    }
  }, [block, isOpen]);

  if (!isOpen || !block) {
    return null;
  }

  const effectiveMaxCols = Math.max(1, maxColumns);
  const effectiveMaxRows = Math.max(1, maxRows);

  const colOptions = [];
  for (let c = 1; c <= effectiveMaxCols; c++) {
    colOptions.push(c);
  }

  const rowOptions = [];
  for (let r = 1; r <= effectiveMaxRows; r++) {
    rowOptions.push(r);
  }

  const handleSave = (): void => {
    const parsedTags = tagsInput
      .split(',')
      .map((t) => t.trim())
      .filter((t) => t.length > 0);

    const updated: IContentBlock = {
      ...block,
      ...formData,
      tags: parsedTags.length > 0 ? parsedTags : undefined
    } as IContentBlock;

    onSave(updated);
    onDismiss();
  };

  return (
    <>
      <div className={styles.backdrop} onClick={onDismiss}>
        <div
          className={styles.sidePanel}
          style={{ width: `${panelWidth}px` }}
          onClick={(e) => e.stopPropagation()}
        >
          {/* Left-edge draggable resize bar */}
          <div
            className={styles.leftResizeHandle}
            onMouseDown={startResizeDrag}
            title="Drag to adjust sidebar width"
          />
          {/* Header */}
          <div className={styles.panelHeader}>
            <div>
              <Subtitle2 style={{ fontWeight: 700, color: tokens.colorNeutralForeground1, display: 'block' }}>
                Edit Card & Metric Properties
              </Subtitle2>
              <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                Configure card layout, icons, metrics, and tags on the right pane.
              </Caption1>
            </div>
            <Button
              appearance="subtle"
              size="small"
              icon={<DismissRegular />}
              onClick={onDismiss}
              aria-label="Close panel"
            />
          </div>

          {/* Body */}
          <div className={styles.panelBody}>
            {/* Block Type */}
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

            {/* Icon Picker */}
            <div className={styles.fieldRow}>
              <Label weight="semibold">Card Icon</Label>
              <div className={styles.iconPreviewBox}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                  <div
                    style={{
                      width: '40px',
                      height: '40px',
                      borderRadius: tokens.borderRadiusMedium,
                      backgroundColor: formData.showIconBackground !== false
                        ? (formData.iconBackgroundColor || tokens.colorBrandBackground2)
                        : 'transparent',
                      border: formData.showIconBackground !== false ? `1px solid ${tokens.colorNeutralStroke2}` : 'none',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      fontSize: '20px',
                      color: formData.iconColor || tokens.colorBrandForeground1,
                      flexShrink: 0
                    }}
                  >
                    {renderFluentIconPreview(formData.iconName, formData.iconColor)}
                  </div>
                  <div>
                    <div style={{ fontSize: '0.85rem', color: tokens.colorNeutralForeground1, fontWeight: 600 }}>
                      {formData.iconName || 'BookAnswers'}
                    </div>
                    <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                      {formData.iconName?.startsWith('svg-') ? 'SVG Icon' : 'Standard Icon'}
                    </Caption1>
                  </div>
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

            {/* Icon Styling: Color & Background Container with Visual Pickers */}
            <div className={styles.twoColRow}>
              <div className={styles.fieldRow}>
                <Label weight="semibold">Icon glyph colour</Label>
                <Caption1 style={{ color: tokens.colorNeutralForeground3, marginBottom: '4px' }}>
                  Brand foreground color
                </Caption1>
                <BrandColorPickerPopover
                  selectedColor={formData.iconColor}
                  onChange={(hex) => setFormData({ ...formData, iconColor: hex })}
                  defaultLabel="Default (brand cyan foreground)"
                  defaultColorHex="#0090DC"
                />
              </div>

              <div className={styles.fieldRow}>
                <Label weight="semibold">Icon background box</Label>
                <Caption1 style={{ color: tokens.colorNeutralForeground3, marginBottom: '4px' }}>
                  Rounded container background
                </Caption1>
                <BrandColorPickerPopover
                  selectedColor={formData.iconBackgroundColor}
                  onChange={(hex) => setFormData({ ...formData, iconBackgroundColor: hex, showIconBackground: true })}
                  defaultLabel="Default (soft tint box)"
                  defaultColorHex="#CCE9F8"
                />
              </div>
            </div>

            <div style={{ margin: '-4px 0 8px 0' }}>
              <Checkbox
                label="Include icon background container"
                checked={formData.showIconBackground !== false}
                onChange={(e, data) => setFormData({ ...formData, showIconBackground: !!data.checked })}
              />
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

            {/* Grid Spanning */}
            <div className={styles.twoColRow}>
              <div className={styles.fieldRow}>
                <Label weight="semibold">Column span (Max {effectiveMaxCols})</Label>
                <Dropdown
                  value={
                    formData.colSpan && formData.colSpan > 1
                      ? `Span ${formData.colSpan} cols`
                      : '1 column (standard)'
                  }
                  onOptionSelect={(e, data) => {
                    setFormData({ ...formData, colSpan: Number(data.optionValue) || 1 });
                  }}
                >
                  {colOptions.map((c) => (
                    <Option key={c} value={c.toString()}>
                      {c === 1 ? '1 column (standard)' : `Span ${c} columns`}
                    </Option>
                  ))}
                </Dropdown>
              </div>

              <div className={styles.fieldRow}>
                <Label weight="semibold">Row span (Max {effectiveMaxRows})</Label>
                <Dropdown
                  value={
                    formData.rowSpan && formData.rowSpan > 1
                      ? `Span ${formData.rowSpan} rows`
                      : '1 row (standard)'
                  }
                  onOptionSelect={(e, data) => {
                    setFormData({ ...formData, rowSpan: Number(data.optionValue) || 1 });
                  }}
                >
                  {rowOptions.map((r) => (
                    <Option key={r} value={r.toString()}>
                      {r === 1 ? '1 row (standard)' : `Span ${r} rows`}
                    </Option>
                  ))}
                </Dropdown>
              </div>
            </div>

            {/* Card Height Behavior */}
            <div className={styles.fieldRow}>
              <Label weight="semibold">Card height behavior</Label>
              <Dropdown
                value={
                  formData.heightMode === 'auto'
                    ? 'Fit content height (independent)'
                    : formData.heightMode === 'equal'
                    ? 'Equal row height (match tallest)'
                    : 'Inherit container setting'
                }
                onOptionSelect={(e, data) => {
                  setFormData({
                    ...formData,
                    heightMode: data.optionValue as 'default' | 'auto' | 'equal'
                  });
                }}
              >
                <Option value="default">Inherit container setting</Option>
                <Option value="auto">Fit content height (independent)</Option>
                <Option value="equal">Equal row height (match tallest)</Option>
              </Dropdown>
            </div>

            {/* Card Background Colour (Independent Visual Popup) */}
            <div className={styles.fieldRow}>
              <Label weight="semibold">Card background colour (Turner &amp; Townsend brand)</Label>
              <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                Choose a brand background or tint for this card via visual popover.
              </Caption1>
              <div style={{ marginTop: '4px' }}>
                <BrandColorPickerPopover
                  selectedColor={formData.backgroundColor}
                  onChange={(hex) => setFormData({ ...formData, backgroundColor: hex })}
                  defaultLabel="Default (white / inherit section background)"
                  defaultColorHex="#FFFFFF"
                />
              </div>
            </div>

            {/* ── Typography ─────────────────────────────────────── */}
            <Divider style={{ margin: '4px 0 2px' }}>
              <Caption1 style={{ color: tokens.colorNeutralForeground3, fontWeight: 600 }}>Typography overrides</Caption1>
            </Divider>

            {/* Text colour */}
            <div className={styles.fieldRow}>
              <Label weight="semibold">Card text colour</Label>
              <Caption1 style={{ color: tokens.colorNeutralForeground3, marginBottom: '4px' }}>
                Overrides title, description, and body text colour on this card.
              </Caption1>
              <BrandColorPickerPopover
                selectedColor={formData.textColor}
                onChange={(hex) => setFormData({ ...formData, textColor: hex })}
                defaultLabel="Default (inherit theme foreground)"
                defaultColorHex="#292929"
              />
            </div>

            {/* Font family + Title font size */}
            <div className={styles.twoColRow}>
              <div className={styles.fieldRow}>
                <Label weight="semibold">Font family</Label>
                <Dropdown
                  value={formData.fontFamily || 'Default (site theme)'}
                  onOptionSelect={(e, data) => {
                    setFormData({ ...formData, fontFamily: data.optionValue === '__default__' ? undefined : (data.optionValue as string) });
                  }}
                >
                  <Option value="__default__">Default (site theme)</Option>
                  <Option value="Inter, sans-serif">Inter</Option>
                  <Option value="'Segoe UI', sans-serif">Segoe UI</Option>
                  <Option value="Georgia, serif">Georgia (serif)</Option>
                  <Option value="'Courier New', monospace">Courier New (mono)</Option>
                  <Option value="Arial, sans-serif">Arial</Option>
                  <Option value="Verdana, sans-serif">Verdana</Option>
                  <Option value="Tahoma, sans-serif">Tahoma</Option>
                  <Option value="'Times New Roman', serif">Times New Roman</Option>
                </Dropdown>
              </div>

              <div className={styles.fieldRow}>
                <Label weight="semibold">Title font size</Label>
                <Dropdown
                  value={formData.titleFontSize || 'Default'}
                  onOptionSelect={(e, data) => {
                    setFormData({ ...formData, titleFontSize: data.optionValue === '__default__' ? undefined : (data.optionValue as string) });
                  }}
                >
                  <Option value="__default__">Default</Option>
                  <Option value="0.85rem">Small (0.85rem)</Option>
                  <Option value="1rem">Normal (1rem)</Option>
                  <Option value="1.1rem">Medium (1.1rem)</Option>
                  <Option value="1.25rem">Large (1.25rem)</Option>
                  <Option value="1.5rem">X-Large (1.5rem)</Option>
                  <Option value="1.75rem">2X-Large (1.75rem)</Option>
                  <Option value="2rem">3X-Large (2rem)</Option>
                </Dropdown>
              </div>
            </div>

            {/* Body font size */}
            <div className={styles.fieldRow}>
              <Label weight="semibold">Body / description font size</Label>
              <Dropdown
                value={formData.bodyFontSize || 'Default'}
                onOptionSelect={(e, data) => {
                  setFormData({ ...formData, bodyFontSize: data.optionValue === '__default__' ? undefined : (data.optionValue as string) });
                }}
              >
                <Option value="__default__">Default</Option>
                <Option value="0.75rem">X-Small (0.75rem)</Option>
                <Option value="0.825rem">Small (0.825rem)</Option>
                <Option value="0.9rem">Slightly small (0.9rem)</Option>
                <Option value="0.95rem">Near-default (0.95rem)</Option>
                <Option value="1rem">Normal (1rem)</Option>
                <Option value="1.1rem">Slightly large (1.1rem)</Option>
                <Option value="1.25rem">Large (1.25rem)</Option>
              </Dropdown>
            </div>

            {/* Global Term Store Taxonomy Integration */}
            <div className={styles.fieldRow}>
              <Label weight="semibold">Global term store • Intranet taxonomy tags</Label>
              <TermStorePicker
                selectedTags={formData.termStoreTags || []}
                onChange={(tags) => setFormData({ ...formData, termStoreTags: tags })}
                isEditMode={true}
              />
            </div>

            {/* Real-Time Live Data API */}
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
                    placeholder="JSON path (e.g. data.metric)"
                    value={formData.liveDataConfig?.jsonPath || ''}
                    onChange={(e, data) => {
                      if (formData.liveDataConfig) {
                        setFormData({
                          ...formData,
                          liveDataConfig: {
                            ...formData.liveDataConfig,
                            jsonPath: data.value
                          }
                        });
                      }
                    }}
                  />
                  <Input
                    placeholder="Prefix (e.g. £)"
                    value={formData.liveDataConfig?.prefix || ''}
                    onChange={(e, data) => {
                      if (formData.liveDataConfig) {
                        setFormData({
                          ...formData,
                          liveDataConfig: {
                            ...formData.liveDataConfig,
                            prefix: data.value
                          }
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
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <Label weight="semibold">Status or category badge</Label>
                  {formData.badge && (
                    <Button
                      size="small"
                      appearance="subtle"
                      onClick={() => setFormData({ ...formData, badge: '' })}
                      style={{ fontSize: '0.72rem', height: '20px', padding: '0 4px', color: tokens.colorNeutralForeground3 }}
                      title="Remove badge (do not display)"
                    >
                      Clear badge
                    </Button>
                  )}
                </div>
                <Input
                  value={formData.badge || ''}
                  placeholder="Leave empty for no badge (e.g. Core, Live, Finance)"
                  onChange={(e, data) => setFormData({ ...formData, badge: data.value })}
                />
              </div>
              <div className={styles.fieldRow}>
                <Label weight="semibold">Additional tags (comma-separated)</Label>
                <Input
                  value={tagsInput}
                  placeholder="Tag 1, Tag 2"
                  onChange={(e, data) => setTagsInput(data.value)}
                />
              </div>
            </div>

            {/* Actions / Links */}
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
                  placeholder="#"
                  onChange={(e, data) => setFormData({ ...formData, linkUrl: data.value })}
                />
              </div>
            </div>
          </div>

          {/* Footer */}
          <div className={styles.panelFooter}>
            <Button appearance="secondary" onClick={onDismiss}>
              Cancel
            </Button>
            <Button
              appearance="primary"
              icon={<SaveRegular />}
              onClick={handleSave}
            >
              Save Card
            </Button>
          </div>
        </div>
      </div>

      {/* Visual Fluent UI 2 Icon Picker */}
      <FluentIconPicker
        isOpen={isIconPickerOpen}
        selectedIconKey={formData.iconName || 'BookAnswers'}
        onSelectIcon={(iconKey) => {
          setFormData({ ...formData, iconName: iconKey });
        }}
        onDismiss={() => setIsIconPickerOpen(false)}
      />
    </>
  );
};
