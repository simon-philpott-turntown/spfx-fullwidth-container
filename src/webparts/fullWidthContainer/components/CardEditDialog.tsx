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
  Divider,
  Portal
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
  StarRegular,
  DeleteRegular,
  FolderOpenRegular,
  ArrowUpRegular,
  ArrowDownRegular
} from '@fluentui/react-icons';
import { IContentBlock, BlockType, ICardItem } from '../models/IContainerModels';
import { TermStorePicker } from './TermStorePicker';
import { FluentIconPicker } from './FluentIconPicker';
import { BrandColorPickerPopover } from './BrandColorPickerPopover';
import { FluentAssetExplorerDialog } from './FluentAssetExplorerDialog';

const useStyles = makeStyles({
  backdrop: {
    position: 'fixed',
    top: 0,
    left: 0,
    width: '100vw',
    height: '100vh',
    backgroundColor: 'rgba(0, 0, 0, 0.15)',
    zIndex: 1000000,
    display: 'flex',
    justifyContent: 'flex-end',
    pointerEvents: 'auto'
  },
  sidePanel: {
    position: 'relative',
    maxWidth: 'calc(100vw - 64px)',
    height: '100vh',
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.borderLeft('1px', 'solid', tokens.colorNeutralStroke1),
    boxShadow: tokens.shadow28,
    display: 'flex',
    flexDirection: 'column',
    boxSizing: 'border-box',
    pointerEvents: 'auto',
    marginRight: '48px'
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
    ...shorthands.gap('8px'),
    position: 'sticky',
    bottom: 0,
    zIndex: 10
  },
  fieldRow: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap(tokens.spacingVerticalXXS)
  },
  twoColRow: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))',
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
  assetPickerService?: import('../services/IAssetPickerService').IAssetPickerService;
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
  assetPickerService,
  onSave,
  onDismiss
}) => {
  const styles = useStyles();

  // Local form state
  const [formData, setFormData] = React.useState<Partial<IContentBlock>>({});
  const [tagsInput, setTagsInput] = React.useState<string>('');
  const [isIconPickerOpen, setIsIconPickerOpen] = React.useState<boolean>(false);
  const [isAssetExplorerOpen, setIsAssetExplorerOpen] = React.useState<boolean>(false);
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
      <Portal>
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

            {/* Transparent Card Mode (Layout Only) */}
            <div className={styles.fieldRow} style={{ marginTop: '2px', marginBottom: '2px' }}>
              <Checkbox
                label="Transparent card (layout container only)"
                checked={formData.transparentCard === true}
                onChange={(e, data) => {
                  setFormData({
                    ...formData,
                    transparentCard: data.checked === true
                  });
                }}
              />
              <Caption1 style={{ color: tokens.colorNeutralForeground3, marginLeft: '28px' }}>
                Hides the card box, border, shadow, and background so only its inner content appears directly on the section canvas.
              </Caption1>
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

            {/* Card Background Image */}
            <div className={styles.fieldRow}>
              <Label weight="semibold">Card background image (optional)</Label>
              <Input
                value={formData.backgroundImage || ''}
                placeholder="https://... or /sites/.../banner.png"
                onChange={(e, data) => setFormData({ ...formData, backgroundImage: data.value })}
              />
              <div style={{ display: 'flex', gap: '8px', marginTop: '6px' }}>
                <Button
                  size="small"
                  appearance="primary"
                  icon={<FolderOpenRegular />}
                  onClick={async () => {
                    if (assetPickerService) {
                      try {
                        const results = await assetPickerService.openNativeFilePicker({
                          title: 'Select Card Background Image',
                          itemType: 'image',
                          acceptedExtensions: ['.png', '.jpg', '.jpeg', '.gif', '.svg', '.webp'],
                          allowMultiple: false
                        });
                        if (results && results.length > 0) {
                          setFormData({
                            ...formData,
                            backgroundImage: results[0].fileAbsoluteUrl || results[0].serverRelativeUrl
                          });
                          return;
                        }
                      } catch (err) {
                        console.warn('[CardEditDialog] Native picker fallback:', err);
                      }
                    }
                    setIsAssetExplorerOpen(true);
                  }}
                >
                  Pick from SharePoint / OneDrive
                </Button>
                {assetPickerService && (
                  <Button
                    size="small"
                    appearance="subtle"
                    onClick={() => setIsAssetExplorerOpen(true)}
                  >
                    Browse Site Assets
                  </Button>
                )}
                {formData.backgroundImage && (
                  <Button
                    size="small"
                    appearance="subtle"
                    onClick={() => setFormData({ ...formData, backgroundImage: '' })}
                  >
                    Clear Image
                  </Button>
                )}
              </div>
              <Caption1 style={{ color: tokens.colorNeutralForeground3, marginTop: '2px' }}>
                Display a background image watermark or pattern behind the card contents.
              </Caption1>
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

            {/* Composable Inner Items Configuration */}
            {formData.items && formData.items.length > 0 && (
              <>
                <Divider style={{ margin: '8px 0 4px' }}>
                  <Caption1 style={{ color: tokens.colorNeutralForeground3, fontWeight: 600 }}>
                    Card Inner Items ({formData.items.length})
                  </Caption1>
                </Divider>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                  {formData.items.map((item, idx) => (
                    <div
                      key={item.id || idx}
                      style={{
                        padding: '8px',
                        borderRadius: tokens.borderRadiusMedium,
                        border: `1px solid ${tokens.colorNeutralStroke2}`,
                        backgroundColor: tokens.colorNeutralBackground2,
                        display: 'flex',
                        flexDirection: 'column',
                        gap: '6px'
                      }}
                    >
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                        <Caption1 style={{ fontWeight: 600, textTransform: 'uppercase', color: tokens.colorBrandForeground1 }}>
                          #{idx + 1} {item.type}
                        </Caption1>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '2px' }}>
                          <Button
                            size="small"
                            appearance="subtle"
                            icon={<ArrowUpRegular />}
                            disabled={idx === 0}
                            title="Move item up"
                            onClick={() => {
                              if (idx === 0) return;
                              const newItems = [...(formData.items || [])];
                              const temp = newItems[idx - 1];
                              newItems[idx - 1] = newItems[idx];
                              newItems[idx] = temp;
                              setFormData({ ...formData, items: newItems });
                            }}
                          />
                          <Button
                            size="small"
                            appearance="subtle"
                            icon={<ArrowDownRegular />}
                            disabled={!formData.items || idx === formData.items.length - 1}
                            title="Move item down"
                            onClick={() => {
                              if (!formData.items || idx === formData.items.length - 1) return;
                              const newItems = [...(formData.items || [])];
                              const temp = newItems[idx + 1];
                              newItems[idx + 1] = newItems[idx];
                              newItems[idx] = temp;
                              setFormData({ ...formData, items: newItems });
                            }}
                          />
                          <Button
                            size="small"
                            appearance="subtle"
                            icon={<DeleteRegular />}
                            title="Remove item"
                            onClick={() => {
                              const newItems = [...(formData.items || [])];
                              newItems.splice(idx, 1);
                              setFormData({ ...formData, items: newItems });
                            }}
                          />
                        </div>
                      </div>

                      {item.type === 'button' && (
                        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '6px' }}>
                          <div>
                            <Label size="small">Button Label</Label>
                            <Input
                              size="small"
                              value={item.buttonLabel || ''}
                              placeholder="Action Button"
                              onChange={(e, data) => {
                                const newItems = [...(formData.items || [])];
                                newItems[idx] = { ...newItems[idx], buttonLabel: data.value };
                                setFormData({ ...formData, items: newItems });
                              }}
                            />
                          </div>
                          <div>
                            <Label size="small">Button URL</Label>
                            <Input
                              size="small"
                              value={item.buttonUrl || ''}
                              placeholder="https://..."
                              onChange={(e, data) => {
                                const newItems = [...(formData.items || [])];
                                newItems[idx] = { ...newItems[idx], buttonUrl: data.value };
                                setFormData({ ...formData, items: newItems });
                              }}
                            />
                          </div>
                        </div>
                      )}

                      {item.type === 'image' && (
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                          <Label size="small">Image URL</Label>
                          <Input
                            size="small"
                            value={item.imageUrl || ''}
                            placeholder="https://images.unsplash.com/..."
                            onChange={(e, data) => {
                              const newItems = [...(formData.items || [])];
                              newItems[idx] = { ...newItems[idx], imageUrl: data.value };
                              setFormData({ ...formData, items: newItems });
                            }}
                          />
                          <Label size="small">Caption / Alt text</Label>
                          <Input
                            size="small"
                            value={item.imageCaption || ''}
                            placeholder="Image caption"
                            onChange={(e, data) => {
                              const newItems = [...(formData.items || [])];
                              newItems[idx] = { ...newItems[idx], imageCaption: data.value, imageAlt: data.value };
                              setFormData({ ...formData, items: newItems });
                            }}
                          />
                        </div>
                      )}

                      {item.type === 'video' && (
                        <div>
                          <Label size="small">Video URL (MP4 / WebM)</Label>
                          <Input
                            size="small"
                            value={item.videoUrl || ''}
                            placeholder="https://..."
                            onChange={(e, data) => {
                              const newItems = [...(formData.items || [])];
                              newItems[idx] = { ...newItems[idx], videoUrl: data.value };
                              setFormData({ ...formData, items: newItems });
                            }}
                          />
                        </div>
                      )}

                      {item.type === 'cta' && (
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '6px' }}>
                            <div>
                              <Label size="small">Heading</Label>
                              <Input
                                size="small"
                                value={item.ctaHeading || ''}
                                placeholder="Call to Action"
                                onChange={(e, data) => {
                                  const newItems = [...(formData.items || [])];
                                  newItems[idx] = { ...newItems[idx], ctaHeading: data.value };
                                  setFormData({ ...formData, items: newItems });
                                }}
                              />
                            </div>
                            <div>
                              <Label size="small">Button Text</Label>
                              <Input
                                size="small"
                                value={item.ctaButtonText || ''}
                                placeholder="Proceed"
                                onChange={(e, data) => {
                                  const newItems = [...(formData.items || [])];
                                  newItems[idx] = { ...newItems[idx], ctaButtonText: data.value };
                                  setFormData({ ...formData, items: newItems });
                                }}
                              />
                            </div>
                          </div>
                          <div>
                            <Label size="small">Button URL</Label>
                            <Input
                              size="small"
                              value={item.ctaButtonUrl || ''}
                              placeholder="https://..."
                              onChange={(e, data) => {
                                const newItems = [...(formData.items || [])];
                                newItems[idx] = { ...newItems[idx], ctaButtonUrl: data.value };
                                setFormData({ ...formData, items: newItems });
                              }}
                            />
                          </div>
                          <div>
                            <Label size="small">Guidance details</Label>
                            <Input
                              size="small"
                              value={item.ctaDescription || ''}
                              placeholder="Guidance text"
                              onChange={(e, data) => {
                                const newItems = [...(formData.items || [])];
                                newItems[idx] = { ...newItems[idx], ctaDescription: data.value };
                                setFormData({ ...formData, items: newItems });
                              }}
                            />
                          </div>
                        </div>
                      )}
                      {item.type === 'editorial' && (
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                          <Label size="small">Kicker / Category</Label>
                          <Input
                            size="small"
                            value={item.editorialKicker || ''}
                            placeholder="e.g. INSIGHT • STRATEGY"
                            onChange={(e, data) => {
                              const newItems = [...(formData.items || [])];
                              newItems[idx] = { ...newItems[idx], editorialKicker: data.value };
                              setFormData({ ...formData, items: newItems });
                            }}
                          />
                          <Label size="small">Editorial Title</Label>
                          <Input
                            size="small"
                            value={item.editorialTitle || ''}
                            placeholder="Article title"
                            onChange={(e, data) => {
                              const newItems = [...(formData.items || [])];
                              newItems[idx] = { ...newItems[idx], editorialTitle: data.value };
                              setFormData({ ...formData, items: newItems });
                            }}
                          />
                          <Label size="small">Summary Body</Label>
                          <Textarea
                            value={item.editorialBody || ''}
                            placeholder="Editorial commentary..."
                            onChange={(e, data) => {
                              const newItems = [...(formData.items || [])];
                              newItems[idx] = { ...newItems[idx], editorialBody: data.value };
                              setFormData({ ...formData, items: newItems });
                            }}
                          />
                          <Label size="small">Read More URL</Label>
                          <Input
                            size="small"
                            value={item.editorialUrl || ''}
                            placeholder="https://..."
                            onChange={(e, data) => {
                              const newItems = [...(formData.items || [])];
                              newItems[idx] = { ...newItems[idx], editorialUrl: data.value };
                              setFormData({ ...formData, items: newItems });
                            }}
                          />
                        </div>
                      )}

                      {item.type === 'hero' && (
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                          <Label size="small">Hero Title</Label>
                          <Input
                            size="small"
                            value={item.heroTitle || ''}
                            placeholder="Hero banner heading"
                            onChange={(e, data) => {
                              const newItems = [...(formData.items || [])];
                              newItems[idx] = { ...newItems[idx], heroTitle: data.value };
                              setFormData({ ...formData, items: newItems });
                            }}
                          />
                          <Label size="small">Hero Subtitle</Label>
                          <Textarea
                            value={item.heroSubtitle || ''}
                            placeholder="Hero description message..."
                            onChange={(e, data) => {
                              const newItems = [...(formData.items || [])];
                              newItems[idx] = { ...newItems[idx], heroSubtitle: data.value };
                              setFormData({ ...formData, items: newItems });
                            }}
                          />
                          <Label size="small">Background Image URL (optional)</Label>
                          <Input
                            size="small"
                            value={item.heroBgUrl || ''}
                            placeholder="https://images.unsplash.com/..."
                            onChange={(e, data) => {
                              const newItems = [...(formData.items || [])];
                              newItems[idx] = { ...newItems[idx], heroBgUrl: data.value };
                              setFormData({ ...formData, items: newItems });
                            }}
                          />
                        </div>
                      )}

                      {item.type === 'link' && (
                        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '6px' }}>
                          <div>
                            <Label size="small">Link Text</Label>
                            <Input
                              size="small"
                              value={item.linkText || item.buttonLabel || ''}
                              placeholder="Access Resource"
                              onChange={(e, data) => {
                                const newItems = [...(formData.items || [])];
                                newItems[idx] = { ...newItems[idx], linkText: data.value, buttonLabel: data.value };
                                setFormData({ ...formData, items: newItems });
                              }}
                            />
                          </div>
                          <div>
                            <Label size="small">Destination URL</Label>
                            <Input
                              size="small"
                              value={item.linkUrl || item.buttonUrl || ''}
                              placeholder="https://..."
                              onChange={(e, data) => {
                                const newItems = [...(formData.items || [])];
                                newItems[idx] = { ...newItems[idx], linkUrl: data.value, buttonUrl: data.value };
                                setFormData({ ...formData, items: newItems });
                              }}
                            />
                          </div>
                        </div>
                      )}

                      {item.type === 'divider' && (
                        <div>
                          <Label size="small">Line Style</Label>
                          <Dropdown
                            value={item.dividerStyle || 'solid'}
                            onOptionSelect={(e, data) => {
                              const newItems = [...(formData.items || [])];
                              newItems[idx] = { ...newItems[idx], dividerStyle: data.optionValue as 'solid' | 'dashed' };
                              setFormData({ ...formData, items: newItems });
                            }}
                          >
                            <Option value="solid">Solid separator</Option>
                            <Option value="dashed">Dashed divider</Option>
                          </Dropdown>
                        </div>
                      )}

                      {item.type === 'text' && (
                        <div>
                          <Label size="small">Text Content</Label>
                          <Textarea
                            value={item.text || ''}
                            placeholder="Section text content..."
                            rows={3}
                            onChange={(e, data) => {
                              const newItems = [...(formData.items || [])];
                              newItems[idx] = { ...newItems[idx], text: data.value };
                              setFormData({ ...formData, items: newItems });
                            }}
                          />
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              </>
            )}
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
      </Portal>

      {/* Visual Fluent UI 2 Icon Picker */}
      <FluentIconPicker
        isOpen={isIconPickerOpen}
        selectedIconKey={formData.iconName || 'BookAnswers'}
        onSelectIcon={(iconKey) => {
          setFormData({ ...formData, iconName: iconKey });
        }}
        onDismiss={() => setIsIconPickerOpen(false)}
      />

      {/* Pure Fluent UI 2 Site Asset Explorer for Card Background */}
      {assetPickerService && (
        <FluentAssetExplorerDialog
          isOpen={isAssetExplorerOpen}
          assetService={assetPickerService}
          title="Select Card Background Image"
          itemType="image"
          allowMultiple={false}
          acceptedExtensions={['.png', '.jpg', '.jpeg', '.gif', '.svg', '.webp']}
          onDismiss={() => setIsAssetExplorerOpen(false)}
          onSelect={(results) => {
            if (results && results.length > 0) {
              setFormData({
                ...formData,
                backgroundImage: results[0].fileAbsoluteUrl || results[0].serverRelativeUrl
              });
            }
            setIsAssetExplorerOpen(false);
          }}
        />
      )}
    </>
  );
};
