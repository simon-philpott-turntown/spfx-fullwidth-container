/**
 * @file CardItemPropertyEditor.tsx
 * @description Contextual Item Property Editor dialog for editing any of the 13 card inner item types.
 * Supports native SharePoint asset picking (ODSP FilePicker) and Fluent UI 2 site asset explorer.
 */

import * as React from 'react';
import {
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogContent,
  DialogActions,
  Badge,
  Button,
  Input,
  Label,
  Textarea,
  Dropdown,
  Option,
  Divider,
  Caption1,
  Portal,
  makeStyles,
  tokens,
  shorthands
} from '@fluentui/react-components';
import {
  DismissRegular,
  SaveRegular,
  AddRegular,
  DeleteRegular,
  CursorClickRegular,
  ImageRegular,
  VideoRegular,
  MegaphoneRegular,
  LinkRegular,
  LineHorizontal1Regular,
  NewsRegular,
  SparkleRegular,
  LinkSquareRegular,
  DataTrendingRegular,
  TagRegular,
  TextDescriptionRegular,
  FolderOpenRegular,
  FilterRegular,
  ArrowRoutingRegular,
  ChevronDownRegular,
  ChevronRightRegular,
  OpenRegular,
  EyeRegular,
  TextAlignLeftRegular,
  TextAlignCenterRegular,
  TextAlignRightRegular,
  ArrowUpRegular,
  ArrowDownRegular
} from '@fluentui/react-icons';
import { ICardItem, ICardItemType, IFilterButtonItem, IProcessStepItem } from '../models/IContainerModels';
import { TermStorePicker } from './TermStorePicker';
import { IAssetPickerService, IFilePickerResult } from '../services/IAssetPickerService';
import { FluentAssetExplorerDialog } from './FluentAssetExplorerDialog';

const useStyles = makeStyles({
  surface: {
    maxWidth: '740px',
    width: '92vw',
    maxHeight: '88vh',
    zIndex: 1000000,
    boxShadow: tokens.shadow28,
    ...shorthands.borderRadius(tokens.borderRadiusXLarge),
    ...shorthands.padding('24px'),
    display: 'flex',
    flexDirection: 'column'
  },
  headerRow: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    paddingBottom: '14px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    marginBottom: '16px'
  },
  fieldRow: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap(tokens.spacingVerticalXXS),
    marginBottom: '14px'
  },
  twoColRow: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    ...shorthands.gap(tokens.spacingHorizontalM),
    marginBottom: '12px'
  },
  threeColRow: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr 1fr',
    ...shorthands.gap(tokens.spacingHorizontalM),
    marginBottom: '12px'
  },
  listContainer: {
    display: 'flex',
    flexDirection: 'column',
    gap: '10px',
    marginTop: '6px'
  },
  stepCard: {
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    boxShadow: tokens.shadow4,
    overflow: 'hidden',
    transitionProperty: 'box-shadow, border-color',
    transitionDuration: '0.15s',
    transitionTimingFunction: 'ease'
  },
  stepCardExpanded: {
    ...shorthands.border('1px', 'solid', tokens.colorBrandStroke1),
    boxShadow: tokens.shadow8
  },
  stepHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    paddingTop: '10px',
    paddingBottom: '10px',
    paddingLeft: '14px',
    paddingRight: '10px',
    backgroundColor: tokens.colorNeutralBackground2,
    cursor: 'pointer',
    userSelect: 'none'
  },
  stepHeaderLeft: {
    display: 'flex',
    alignItems: 'center',
    gap: '10px',
    flexGrow: 1,
    minWidth: 0
  },
  stepHeaderRight: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    flexShrink: 0
  },
  stepBody: {
    paddingTop: '14px',
    paddingBottom: '14px',
    paddingLeft: '16px',
    paddingRight: '16px',
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
    backgroundColor: tokens.colorNeutralBackground1,
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`
  },
  stepSectionDivider: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    marginTop: '2px',
    marginBottom: '2px',
    color: tokens.colorNeutralForeground3,
    fontSize: '0.75rem',
    fontWeight: 600,
    textTransform: 'uppercase',
    letterSpacing: '0.04em'
  },
  miniPreviewContainer: {
    paddingTop: '10px',
    paddingBottom: '10px',
    paddingLeft: '14px',
    paddingRight: '14px',
    marginBottom: '16px',
    backgroundColor: tokens.colorNeutralBackground2,
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    ...shorthands.border('1px', 'dashed', tokens.colorNeutralStroke1),
    display: 'flex',
    flexDirection: 'column',
    gap: '6px'
  },
  miniPreviewBar: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    overflowX: 'auto',
    paddingTop: '4px',
    paddingBottom: '4px'
  },
  miniChevron: {
    display: 'flex',
    flexDirection: 'column',
    paddingTop: '4px',
    paddingBottom: '4px',
    paddingLeft: '10px',
    paddingRight: '10px',
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusSmall,
    fontSize: '0.72rem',
    whiteSpace: 'nowrap',
    boxShadow: tokens.shadow2
  },
  listItemRow: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr auto',
    gap: '8px',
    alignItems: 'center',
    padding: '8px 10px',
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`
  },
  pickerActionsRow: {
    display: 'flex',
    gap: '8px',
    marginTop: '6px'
  }
});

export interface ICardItemPropertyEditorProps {
  isOpen: boolean;
  item: ICardItem | null;
  assetPickerService?: IAssetPickerService;
  onDismiss: () => void;
  onSave: (updatedItem: ICardItem) => void;
}

/**
 * Returns an icon representing the card item type.
 */
const getItemIcon = (type?: ICardItemType): React.ReactElement => {
  switch (type) {
    case 'button': return <CursorClickRegular />;
    case 'filterButtons': return <FilterRegular />;
    case 'processModel': return <ArrowRoutingRegular />;
    case 'image': return <ImageRegular />;
    case 'video': return <VideoRegular />;
    case 'cta': return <MegaphoneRegular />;
    case 'link': return <LinkRegular />;
    case 'divider': return <LineHorizontal1Regular />;
    case 'editorial': return <NewsRegular />;
    case 'hero': return <SparkleRegular />;
    case 'gallery': return <ImageRegular />;
    case 'quickLinks': return <LinkSquareRegular />;
    case 'liveData': return <DataTrendingRegular />;
    case 'termStoreTags': return <TagRegular />;
    default: return <TextDescriptionRegular />;
  }
};

export const CardItemPropertyEditor: React.FC<ICardItemPropertyEditorProps> = ({
  isOpen,
  item,
  assetPickerService,
  onDismiss,
  onSave
}) => {
  const styles = useStyles();
  const [formData, setFormData] = React.useState<ICardItem | null>(null);
  const [expandedStepId, setExpandedStepId] = React.useState<string | null>(null);
  const [isFluentExplorerOpen, setIsFluentExplorerOpen] = React.useState<boolean>(false);
  const [explorerMode, setExplorerMode] = React.useState<{
    type: 'image' | 'video' | 'gallery' | 'hero' | 'cta' | 'editorial';
    allowMultiple: boolean;
  }>({ type: 'image', allowMultiple: false });

  React.useEffect(() => {
    if (item) {
      setFormData(JSON.parse(JSON.stringify(item)));
      if (item.type === 'processModel' && item.processSteps && item.processSteps.length > 0) {
        setExpandedStepId(item.processSteps[0].id || 'step-0');
      }
    } else {
      setFormData(null);
      setExpandedStepId(null);
    }
  }, [item]);

  if (!isOpen || !formData) return null;

  const handleFieldChange = (field: keyof ICardItem, value: any): void => {
    setFormData((prev) => (prev ? { ...prev, [field]: value } : null));
  };

  /**
   * Primary: Trigger native SharePoint hosted FilePicker modal dialog.
   * If cancelled or unavailable, fall back to pure Fluent UI 2 site asset explorer.
   */
  const handleLaunchNativePicker = async (
    pickerType: 'image' | 'video' | 'gallery' | 'hero' | 'cta' | 'editorial',
    allowMultiple: boolean = false
  ): Promise<void> => {
    if (assetPickerService) {
      const accepted = pickerType === 'video'
        ? ['.mp4', '.webm', '.mov']
        : ['.png', '.jpg', '.jpeg', '.gif', '.svg', '.webp'];

      try {
        const results = await assetPickerService.openNativeFilePicker({
          title: `Select ${pickerType === 'video' ? 'Video' : 'Image'}`,
          itemType: pickerType === 'video' ? 'video' : 'image',
          acceptedExtensions: accepted,
          allowMultiple
        });

        if (results && results.length > 0) {
          applySelectedAssets(pickerType, results);
          return;
        }
      } catch (err) {
        console.warn('[CardItemPropertyEditor] Native picker fallback:', err);
      }
    }

    // Fallback: Open Pure Fluent UI 2 Explorer Modal
    setExplorerMode({ type: pickerType, allowMultiple });
    setIsFluentExplorerOpen(true);
  };

  const applySelectedAssets = (
    pickerType: 'image' | 'video' | 'gallery' | 'hero' | 'cta' | 'editorial',
    results: IFilePickerResult[]
  ): void => {
    if (!results || results.length === 0) return;

    if (pickerType === 'gallery') {
      const existing = formData.galleryImages || [];
      const newImages = results.map(r => ({
        url: r.fileAbsoluteUrl || r.serverRelativeUrl,
        caption: r.fileName
      }));
      handleFieldChange('galleryImages', [...existing, ...newImages]);
    } else if (pickerType === 'video') {
      handleFieldChange('videoUrl', results[0].fileAbsoluteUrl || results[0].serverRelativeUrl);
    } else if (pickerType === 'hero') {
      handleFieldChange('heroBgUrl', results[0].fileAbsoluteUrl || results[0].serverRelativeUrl);
    } else if (pickerType === 'cta') {
      handleFieldChange('ctaBgUrl', results[0].fileAbsoluteUrl || results[0].serverRelativeUrl);
    } else if (pickerType === 'editorial') {
      handleFieldChange('editorialImageUrl', results[0].fileAbsoluteUrl || results[0].serverRelativeUrl);
    } else {
      // Standard image
      handleFieldChange('imageUrl', results[0].fileAbsoluteUrl || results[0].serverRelativeUrl);
      if (!formData.imageCaption) {
        handleFieldChange('imageCaption', results[0].fileName);
      }
      if (!formData.imageAlt) {
        handleFieldChange('imageAlt', results[0].fileName);
      }
    }
  };

  const renderTypeSpecificFields = (): React.ReactElement => {
    switch (formData.type) {
      case 'button':
        return (
          <>
            <div className={styles.twoColRow}>
              <div className={styles.fieldRow}>
                <Label weight="semibold">Button label</Label>
                <Input
                  value={formData.buttonLabel || ''}
                  placeholder="e.g. View documentation"
                  onChange={(e, data) => handleFieldChange('buttonLabel', data.value)}
                />
              </div>
              <div className={styles.fieldRow}>
                <Label weight="semibold">Appearance variant</Label>
                <Dropdown
                  value={formData.buttonVariant || 'primary'}
                  onOptionSelect={(e, data) => handleFieldChange('buttonVariant', data.optionValue)}
                >
                  <Option value="primary">Primary (solid brand)</Option>
                  <Option value="outline">Outline (bordered)</Option>
                  <Option value="subtle">Subtle (text button)</Option>
                </Dropdown>
              </div>
            </div>
            <div className={styles.fieldRow}>
              <Label weight="semibold">Destination link URL</Label>
              <Input
                value={formData.buttonUrl || ''}
                placeholder="https://..."
                onChange={(e, data) => handleFieldChange('buttonUrl', data.value)}
              />
            </div>
          </>
        );

      case 'image':
        return (
          <>
            <div className={styles.fieldRow}>
              <Label weight="semibold">Image source URL</Label>
              <Input
                value={formData.imageUrl || ''}
                placeholder="https://images.unsplash.com/... or /sites/.../image.png"
                onChange={(e, data) => handleFieldChange('imageUrl', data.value)}
              />
              <div className={styles.pickerActionsRow}>
                <Button
                  size="small"
                  appearance="primary"
                  icon={<FolderOpenRegular />}
                  onClick={() => void handleLaunchNativePicker('image', false)}
                >
                  Pick from SharePoint / OneDrive
                </Button>
                {assetPickerService && (
                  <Button
                    size="small"
                    appearance="subtle"
                    onClick={() => {
                      setExplorerMode({ type: 'image', allowMultiple: false });
                      setIsFluentExplorerOpen(true);
                    }}
                  >
                    Browse Site Library Explorer
                  </Button>
                )}
              </div>
              <Caption1 style={{ color: tokens.colorNeutralForeground3, marginTop: '4px' }}>
                Select an image via native SharePoint File Picker (Site Assets, OneDrive, Stock Images) or paste a URL.
              </Caption1>
            </div>
            <div className={styles.twoColRow}>
              <div className={styles.fieldRow}>
                <Label weight="semibold">Caption</Label>
                <Input
                  value={formData.imageCaption || ''}
                  placeholder="e.g. Project overview schematic"
                  onChange={(e, data) => handleFieldChange('imageCaption', data.value)}
                />
              </div>
              <div className={styles.fieldRow}>
                <Label weight="semibold">Alt text (Accessibility)</Label>
                <Input
                  value={formData.imageAlt || ''}
                  placeholder="Descriptive image summary"
                  onChange={(e, data) => handleFieldChange('imageAlt', data.value)}
                />
              </div>
            </div>
            {formData.imageUrl && (
              <div style={{ marginTop: '8px', border: `1px solid ${tokens.colorNeutralStroke2}`, borderRadius: '6px', overflow: 'hidden' }}>
                <img
                  src={formData.imageUrl}
                  alt={formData.imageAlt || 'Preview'}
                  style={{ width: '100%', maxHeight: '160px', objectFit: 'cover', display: 'block' }}
                />
              </div>
            )}
          </>
        );

      case 'video':
        return (
          <>
            <div className={styles.fieldRow}>
              <Label weight="semibold">Video media URL (MP4, WebM)</Label>
              <Input
                value={formData.videoUrl || ''}
                placeholder="https://.../video.mp4"
                onChange={(e, data) => handleFieldChange('videoUrl', data.value)}
              />
              <div className={styles.pickerActionsRow}>
                <Button
                  size="small"
                  appearance="primary"
                  icon={<FolderOpenRegular />}
                  onClick={() => void handleLaunchNativePicker('video', false)}
                >
                  Pick from SharePoint / OneDrive
                </Button>
                {assetPickerService && (
                  <Button
                    size="small"
                    appearance="subtle"
                    onClick={() => {
                      setExplorerMode({ type: 'video', allowMultiple: false });
                      setIsFluentExplorerOpen(true);
                    }}
                  >
                    Browse Library Explorer
                  </Button>
                )}
              </div>
              <Caption1 style={{ color: tokens.colorNeutralForeground3, marginTop: '4px' }}>
                Select a video file from SharePoint/Stream or paste a direct video link.
              </Caption1>
            </div>
            {formData.videoUrl && (
              <div style={{ marginTop: '8px', borderRadius: '6px', overflow: 'hidden' }}>
                <video
                  src={formData.videoUrl}
                  controls
                  style={{ width: '100%', maxHeight: '160px', display: 'block', backgroundColor: '#000' }}
                />
              </div>
            )}
          </>
        );

      case 'cta':
        return (
          <>
            <div className={styles.fieldRow}>
              <Label weight="semibold">Call to action heading</Label>
              <Input
                value={formData.ctaHeading || ''}
                placeholder="e.g. Ready to transform your infrastructure?"
                onChange={(e, data) => handleFieldChange('ctaHeading', data.value)}
              />
            </div>
            <div className={styles.fieldRow}>
              <Label weight="semibold">Description / Guidance</Label>
              <Textarea
                value={formData.ctaDescription || ''}
                placeholder="Describe next steps or value proposition..."
                onChange={(e, data) => handleFieldChange('ctaDescription', data.value)}
              />
            </div>
            <div className={styles.twoColRow}>
              <div className={styles.fieldRow}>
                <Label weight="semibold">Action button label</Label>
                <Input
                  value={formData.ctaButtonText || ''}
                  placeholder="e.g. Get Started"
                  onChange={(e, data) => handleFieldChange('ctaButtonText', data.value)}
                />
              </div>
              <div className={styles.fieldRow}>
                <Label weight="semibold">Action button URL</Label>
                <Input
                  value={formData.ctaButtonUrl || ''}
                  placeholder="https://..."
                  onChange={(e, data) => handleFieldChange('ctaButtonUrl', data.value)}
                />
              </div>
            </div>
            <div className={styles.fieldRow}>
              <Label weight="semibold">CTA background image (optional)</Label>
              <Input
                value={formData.ctaBgUrl || ''}
                placeholder="https://... or /sites/.../cta-banner.png"
                onChange={(e, data) => handleFieldChange('ctaBgUrl', data.value)}
              />
              <div className={styles.pickerActionsRow}>
                <Button
                  size="small"
                  appearance="subtle"
                  icon={<FolderOpenRegular />}
                  onClick={() => void handleLaunchNativePicker('cta', false)}
                >
                  Pick CTA Background from SharePoint
                </Button>
              </div>
            </div>
          </>
        );

      case 'editorial':
        return (
          <>
            <div className={styles.fieldRow}>
              <Label weight="semibold">Kicker / Category tag</Label>
              <Input
                value={formData.editorialKicker || ''}
                placeholder="e.g. INSIGHT • STRATEGY"
                onChange={(e, data) => handleFieldChange('editorialKicker', data.value)}
              />
            </div>
            <div className={styles.fieldRow}>
              <Label weight="semibold">Editorial title</Label>
              <Input
                value={formData.editorialTitle || ''}
                placeholder="e.g. Net zero delivery frameworks"
                onChange={(e, data) => handleFieldChange('editorialTitle', data.value)}
              />
            </div>
            <div className={styles.fieldRow}>
              <Label weight="semibold">Lead thumbnail image (optional)</Label>
              <Input
                value={formData.editorialImageUrl || ''}
                placeholder="https://... or /sites/.../news.png"
                onChange={(e, data) => handleFieldChange('editorialImageUrl', data.value)}
              />
              <div className={styles.pickerActionsRow}>
                <Button
                  size="small"
                  appearance="subtle"
                  icon={<FolderOpenRegular />}
                  onClick={() => void handleLaunchNativePicker('editorial', false)}
                >
                  Pick Lead Image from SharePoint
                </Button>
              </div>
            </div>
            <div className={styles.fieldRow}>
              <Label weight="semibold">Summary paragraph</Label>
              <Textarea
                value={formData.editorialBody || ''}
                placeholder="Brief editorial commentary or executive summary..."
                onChange={(e, data) => handleFieldChange('editorialBody', data.value)}
              />
            </div>
            <div className={styles.fieldRow}>
              <Label weight="semibold">Read more URL (optional)</Label>
              <Input
                value={formData.editorialUrl || ''}
                placeholder="https://..."
                onChange={(e, data) => handleFieldChange('editorialUrl', data.value)}
              />
            </div>
          </>
        );

      case 'hero':
        return (
          <>
            <div className={styles.fieldRow}>
              <Label weight="semibold">Hero title</Label>
              <Input
                value={formData.heroTitle || ''}
                placeholder="e.g. Turner & Townsend Programme Advisory"
                onChange={(e, data) => handleFieldChange('heroTitle', data.value)}
              />
            </div>
            <div className={styles.fieldRow}>
              <Label weight="semibold">Hero subtitle / message</Label>
              <Textarea
                value={formData.heroSubtitle || ''}
                placeholder="Hero supporting text or highlight description..."
                onChange={(e, data) => handleFieldChange('heroSubtitle', data.value)}
              />
            </div>
            <div className={styles.fieldRow}>
              <Label weight="semibold">Background image URL (optional)</Label>
              <Input
                value={formData.heroBgUrl || ''}
                placeholder="https://images.unsplash.com/... (optional banner)"
                onChange={(e, data) => handleFieldChange('heroBgUrl', data.value)}
              />
              <div className={styles.pickerActionsRow}>
                <Button
                  size="small"
                  appearance="subtle"
                  icon={<FolderOpenRegular />}
                  onClick={() => void handleLaunchNativePicker('hero', false)}
                >
                  Pick Background from SharePoint
                </Button>
              </div>
            </div>
          </>
        );

      case 'link':
        return (
          <>
            <div className={styles.fieldRow}>
              <Label weight="semibold">Link text</Label>
              <Input
                value={formData.linkText || ''}
                placeholder="e.g. Access Commercial Management Portal"
                onChange={(e, data) => handleFieldChange('linkText', data.value)}
              />
            </div>
            <div className={styles.fieldRow}>
              <Label weight="semibold">Target destination URL</Label>
              <Input
                value={formData.linkUrl || ''}
                placeholder="https://..."
                onChange={(e, data) => handleFieldChange('linkUrl', data.value)}
              />
            </div>
          </>
        );

      case 'divider':
        return (
          <>
            <div className={styles.fieldRow}>
              <Label weight="semibold">Line style</Label>
              <Dropdown
                value={formData.dividerStyle || 'solid'}
                onOptionSelect={(e, data) => handleFieldChange('dividerStyle', data.optionValue)}
              >
                <Option value="solid">Solid separator line</Option>
                <Option value="dashed">Dashed divider line</Option>
              </Dropdown>
            </div>
          </>
        );

      case 'quickLinks':
        const links = formData.quickLinks || [];
        return (
          <>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <Label weight="semibold">Quick Links List ({links.length})</Label>
              <Button
                size="small"
                appearance="subtle"
                icon={<AddRegular />}
                onClick={() => {
                  const updated = [...links, { label: 'New Link', url: '#' }];
                  handleFieldChange('quickLinks', updated);
                }}
              >
                Add Link
              </Button>
            </div>
            <div className={styles.listContainer}>
              {links.map((lnk, idx) => (
                <div key={idx} className={styles.listItemRow}>
                  <Input
                    size="small"
                    value={lnk.label}
                    placeholder="Link Label"
                    onChange={(e, data) => {
                      const updated = [...links];
                      updated[idx] = { ...updated[idx], label: data.value };
                      handleFieldChange('quickLinks', updated);
                    }}
                  />
                  <Input
                    size="small"
                    value={lnk.url}
                    placeholder="https://..."
                    onChange={(e, data) => {
                      const updated = [...links];
                      updated[idx] = { ...updated[idx], url: data.value };
                      handleFieldChange('quickLinks', updated);
                    }}
                  />
                  <Button
                    size="small"
                    appearance="subtle"
                    icon={<ArrowUpRegular />}
                    disabled={idx === 0}
                    onClick={() => {
                      if (idx === 0) return;
                      const updated = [...links];
                      const temp = updated[idx - 1];
                      updated[idx - 1] = updated[idx];
                      updated[idx] = temp;
                      handleFieldChange('quickLinks', updated);
                    }}
                    title="Move link up"
                  />
                  <Button
                    size="small"
                    appearance="subtle"
                    icon={<ArrowDownRegular />}
                    disabled={idx === links.length - 1}
                    onClick={() => {
                      if (idx === links.length - 1) return;
                      const updated = [...links];
                      const temp = updated[idx + 1];
                      updated[idx + 1] = updated[idx];
                      updated[idx] = temp;
                      handleFieldChange('quickLinks', updated);
                    }}
                    title="Move link down"
                  />
                  <Button
                    size="small"
                    appearance="subtle"
                    icon={<DeleteRegular />}
                    onClick={() => {
                      const updated = links.filter((_, i) => i !== idx);
                      handleFieldChange('quickLinks', updated);
                    }}
                  />
                </div>
              ))}
            </div>
          </>
        );

      case 'gallery':
        const images = formData.galleryImages || [];
        return (
          <>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <Label weight="semibold">Gallery Images ({images.length})</Label>
              <div style={{ display: 'flex', gap: '6px' }}>
                <Button
                  size="small"
                  appearance="primary"
                  icon={<FolderOpenRegular />}
                  onClick={() => void handleLaunchNativePicker('gallery', true)}
                >
                  Pick from SharePoint
                </Button>
                <Button
                  size="small"
                  appearance="subtle"
                  icon={<AddRegular />}
                  onClick={() => {
                    const updated = [...images, { url: 'https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?w=400', caption: 'Image' }];
                    handleFieldChange('galleryImages', updated);
                  }}
                >
                  Add URL
                </Button>
              </div>
            </div>
            <div className={styles.listContainer}>
              {images.map((img, idx) => (
                <div key={idx} className={styles.listItemRow}>
                  <Input
                    size="small"
                    value={img.url}
                    placeholder="https://... Image URL"
                    onChange={(e, data) => {
                      const updated = [...images];
                      updated[idx] = { ...updated[idx], url: data.value };
                      handleFieldChange('galleryImages', updated);
                    }}
                  />
                  <Input
                    size="small"
                    value={img.caption || ''}
                    placeholder="Caption (optional)"
                    onChange={(e, data) => {
                      const updated = [...images];
                      updated[idx] = { ...updated[idx], caption: data.value };
                      handleFieldChange('galleryImages', updated);
                    }}
                  />
                  <Button
                    size="small"
                    appearance="subtle"
                    icon={<ArrowUpRegular />}
                    disabled={idx === 0}
                    onClick={() => {
                      if (idx === 0) return;
                      const updated = [...images];
                      const temp = updated[idx - 1];
                      updated[idx - 1] = updated[idx];
                      updated[idx] = temp;
                      handleFieldChange('galleryImages', updated);
                    }}
                    title="Move image up"
                  />
                  <Button
                    size="small"
                    appearance="subtle"
                    icon={<ArrowDownRegular />}
                    disabled={idx === images.length - 1}
                    onClick={() => {
                      if (idx === images.length - 1) return;
                      const updated = [...images];
                      const temp = updated[idx + 1];
                      updated[idx + 1] = updated[idx];
                      updated[idx] = temp;
                      handleFieldChange('galleryImages', updated);
                    }}
                    title="Move image down"
                  />
                  <Button
                    size="small"
                    appearance="subtle"
                    icon={<DeleteRegular />}
                    onClick={() => {
                      const updated = images.filter((_, i) => i !== idx);
                      handleFieldChange('galleryImages', updated);
                    }}
                  />
                </div>
              ))}
            </div>
          </>
        );

      case 'liveData':
        const liveCfg = formData.liveDataConfig || { apiUrl: '', jsonPath: 'value', prefix: '£' };
        return (
          <>
            <div className={styles.fieldRow}>
              <Label weight="semibold">API Endpoint URL</Label>
              <Input
                value={liveCfg.apiUrl || ''}
                placeholder="https://api... or demo-api/burnDown"
                onChange={(e, data) => {
                  handleFieldChange('liveDataConfig', { ...liveCfg, apiUrl: data.value });
                }}
              />
            </div>
            <div className={styles.twoColRow}>
              <div className={styles.fieldRow}>
                <Label weight="semibold">JSON extraction path</Label>
                <Input
                  value={liveCfg.jsonPath || ''}
                  placeholder="data.metric"
                  onChange={(e, data) => {
                    handleFieldChange('liveDataConfig', { ...liveCfg, jsonPath: data.value });
                  }}
                />
              </div>
              <div className={styles.fieldRow}>
                <Label weight="semibold">Currency / Unit prefix</Label>
                <Input
                  value={liveCfg.prefix || ''}
                  placeholder="£"
                  onChange={(e, data) => {
                    handleFieldChange('liveDataConfig', { ...liveCfg, prefix: data.value });
                  }}
                />
              </div>
            </div>
            <div className={styles.fieldRow}>
              <Label weight="semibold">Refresh interval (seconds)</Label>
              <Input
                type="number"
                value={String(liveCfg.refreshIntervalSeconds || 30)}
                onChange={(e, data) => {
                  handleFieldChange('liveDataConfig', { ...liveCfg, refreshIntervalSeconds: Number(data.value) || 30 });
                }}
              />
            </div>
          </>
        );

      case 'termStoreTags':
        return (
          <>
            <div className={styles.fieldRow}>
              <Label weight="semibold">Term Store Tags Selection</Label>
              <Caption1 style={{ color: tokens.colorNeutralForeground3, marginBottom: '6px' }}>
                Browse or select taxonomy tags from the Global Intranet Term Store.
              </Caption1>
              <TermStorePicker
                selectedTags={formData.termStoreTags || []}
                onChange={(tags) => handleFieldChange('termStoreTags', tags)}
                isEditMode={true}
              />
            </div>
          </>
        );

      case 'filterButtons':
        const filterBtns: IFilterButtonItem[] = formData.filterButtons || [];
        return (
          <>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
              <Label weight="semibold">Filter Buttons ({filterBtns.length})</Label>
              <Button
                size="small"
                appearance="subtle"
                icon={<AddRegular />}
                onClick={() => {
                  const updated: IFilterButtonItem[] = [
                    ...filterBtns,
                    {
                      id: `btn-${Date.now()}-${Math.random().toString(36).substr(2, 4)}`,
                      label: 'New filter'
                    }
                  ];
                  handleFieldChange('filterButtons', updated);
                }}
              >
                Add Button
              </Button>
            </div>
            <Caption1 style={{ color: tokens.colorNeutralForeground3, marginBottom: '8px', display: 'block' }}>
              Clicking these pill buttons will filter cards whose title, summary, or tags match the button title.
            </Caption1>
            <div className={styles.listContainer}>
              {filterBtns.map((btn, idx) => (
                <div key={btn.id || idx} className={styles.listItemRow}>
                  <Input
                    size="small"
                    value={btn.label}
                    placeholder="Button Title (e.g. Programme advisory)"
                    onChange={(e, data) => {
                      const updated = [...filterBtns];
                      updated[idx] = { ...updated[idx], label: data.value };
                      handleFieldChange('filterButtons', updated);
                    }}
                  />
                  <Input
                    size="small"
                    value={btn.filterValue || ''}
                    placeholder="Filter criterion (optional)"
                    onChange={(e, data) => {
                      const updated = [...filterBtns];
                      updated[idx] = { ...updated[idx], filterValue: data.value };
                      handleFieldChange('filterButtons', updated);
                    }}
                  />
                  <Button
                    size="small"
                    appearance="subtle"
                    icon={<ArrowUpRegular />}
                    disabled={idx === 0}
                    onClick={() => {
                      if (idx === 0) return;
                      const updated = [...filterBtns];
                      const temp = updated[idx - 1];
                      updated[idx - 1] = updated[idx];
                      updated[idx] = temp;
                      handleFieldChange('filterButtons', updated);
                    }}
                    title="Move button up"
                  />
                  <Button
                    size="small"
                    appearance="subtle"
                    icon={<ArrowDownRegular />}
                    disabled={idx === filterBtns.length - 1}
                    onClick={() => {
                      if (idx === filterBtns.length - 1) return;
                      const updated = [...filterBtns];
                      const temp = updated[idx + 1];
                      updated[idx + 1] = updated[idx];
                      updated[idx] = temp;
                      handleFieldChange('filterButtons', updated);
                    }}
                    title="Move button down"
                  />
                  <Button
                    size="small"
                    appearance="subtle"
                    icon={<DeleteRegular />}
                    onClick={() => {
                      const updated = filterBtns.filter((_, i) => i !== idx);
                      handleFieldChange('filterButtons', updated);
                    }}
                  />
                </div>
              ))}
            </div>
          </>
        );

      case 'processModel':
        const steps: IProcessStepItem[] = formData.processSteps || [];
        return (
          <>
            {/* Live Visual Mini-Preview */}
            <div className={styles.miniPreviewContainer}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <EyeRegular style={{ fontSize: '14px', color: tokens.colorBrandForeground1 }} />
                  <span style={{ fontSize: '0.75rem', fontWeight: 600, color: tokens.colorNeutralForeground2, textTransform: 'uppercase', letterSpacing: '0.04em' }}>
                    Live Flow Preview ({steps.length} Stages)
                  </span>
                </div>
                <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                  Click a step card below to edit details
                </Caption1>
              </div>
              <div className={styles.miniPreviewBar}>
                {steps.map((st, i) => (
                  <React.Fragment key={st.id || i}>
                    <div
                      className={styles.miniChevron}
                      style={{
                        borderColor: expandedStepId === (st.id || `step-${i}`) ? tokens.colorBrandStroke1 : tokens.colorNeutralStroke2,
                        backgroundColor: expandedStepId === (st.id || `step-${i}`) ? tokens.colorBrandBackground2 : tokens.colorNeutralBackground1,
                        cursor: 'pointer'
                      }}
                      onClick={() => setExpandedStepId(st.id || `step-${i}`)}
                    >
                      <span style={{ fontWeight: 700, color: tokens.colorBrandForeground1, fontSize: '0.68rem' }}>
                        {st.stageNumber || `Stage ${i + 1}`}
                      </span>
                      <span style={{ fontWeight: 600, color: tokens.colorNeutralForeground1, fontSize: '0.72rem' }}>
                        {st.title || 'Untitled'}
                      </span>
                    </div>
                    {i < steps.length - 1 && (
                      <span style={{ color: tokens.colorNeutralForeground3, fontSize: '0.68rem', margin: '0 1px' }}>▶</span>
                    )}
                  </React.Fragment>
                ))}
              </div>
            </div>

            {/* Stages Header & Add Button */}
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <Label weight="semibold" size="medium">Process Stages & Capabilities</Label>
                <Badge appearance="tint" color="brand" size="small">{steps.length} steps</Badge>
              </div>
              <Button
                size="small"
                appearance="primary"
                icon={<AddRegular />}
                onClick={() => {
                  const newIndex = steps.length + 1;
                  const newId = `step-${Date.now()}-${Math.random().toString(36).substr(2, 4)}`;
                  const updated: IProcessStepItem[] = [
                    ...steps,
                    {
                      id: newId,
                      stageNumber: `Stage ${newIndex}`,
                      title: `Stage ${newIndex} Title`,
                      description: 'Define stage objective and actions',
                      metricBadge: `${newIndex} capabilities`,
                      actionType: 'filter'
                    }
                  ];
                  handleFieldChange('processSteps', updated);
                  setExpandedStepId(newId);
                }}
              >
                Add Stage
              </Button>
            </div>

            {/* Collapsible Step Cards List */}
            <div className={styles.listContainer} style={{ maxHeight: '420px', overflowY: 'auto', paddingRight: '4px' }}>
              {steps.map((step, idx) => {
                const stepKey = step.id || `step-${idx}`;
                const isExpanded = expandedStepId === stepKey || (expandedStepId === null && idx === 0);

                return (
                  <div
                    key={stepKey}
                    className={`${styles.stepCard} ${isExpanded ? styles.stepCardExpanded : ''}`}
                  >
                    {/* Collapsible Card Header Bar */}
                    <div
                      className={styles.stepHeader}
                      onClick={() => setExpandedStepId(isExpanded ? '' : stepKey)}
                      title={isExpanded ? 'Click to collapse' : 'Click to expand and edit'}
                    >
                      <div className={styles.stepHeaderLeft}>
                        {isExpanded ? (
                          <ChevronDownRegular style={{ fontSize: '14px', color: tokens.colorBrandForeground1 }} />
                        ) : (
                          <ChevronRightRegular style={{ fontSize: '14px', color: tokens.colorNeutralForeground3 }} />
                        )}
                        <Badge
                          appearance={isExpanded ? 'filled' : 'tint'}
                          color={isExpanded ? 'brand' : 'informative'}
                          size="small"
                          style={{ fontWeight: 700 }}
                        >
                          {step.stageNumber || `Stage ${idx + 1}`}
                        </Badge>
                        <span style={{ fontWeight: 600, fontSize: '0.85rem', color: tokens.colorNeutralForeground1 }}>
                          {step.title || 'Untitled Stage'}
                        </span>
                        {step.metricBadge && (
                          <Badge appearance="tint" color="informative" size="small" style={{ fontSize: '0.7rem' }}>
                            {step.metricBadge}
                          </Badge>
                        )}
                      </div>

                      <div className={styles.stepHeaderRight}>
                        <span style={{ fontSize: '0.72rem', color: tokens.colorNeutralForeground3 }}>
                          {step.actionType === 'navigate' ? '🔗 URL' : '🔍 Filter'}
                        </span>
                        <Button
                          size="small"
                          appearance="subtle"
                          icon={<ArrowUpRegular />}
                          disabled={idx === 0}
                          onClick={(e) => {
                            e.stopPropagation();
                            if (idx === 0) return;
                            const updated = [...steps];
                            const temp = updated[idx - 1];
                            updated[idx - 1] = updated[idx];
                            updated[idx] = temp;
                            handleFieldChange('processSteps', updated);
                          }}
                          title="Move stage up"
                        />
                        <Button
                          size="small"
                          appearance="subtle"
                          icon={<ArrowDownRegular />}
                          disabled={idx === steps.length - 1}
                          onClick={(e) => {
                            e.stopPropagation();
                            if (idx === steps.length - 1) return;
                            const updated = [...steps];
                            const temp = updated[idx + 1];
                            updated[idx + 1] = updated[idx];
                            updated[idx] = temp;
                            handleFieldChange('processSteps', updated);
                          }}
                          title="Move stage down"
                        />
                        <Button
                          size="small"
                          appearance="subtle"
                          icon={<DeleteRegular />}
                          onClick={(e) => {
                            e.stopPropagation();
                            const updated = steps.filter((_, i) => i !== idx);
                            handleFieldChange('processSteps', updated);
                            if (expandedStepId === stepKey) {
                              setExpandedStepId(updated.length > 0 ? (updated[0].id || 'step-0') : null);
                            }
                          }}
                          title="Delete stage"
                        />
                      </div>
                    </div>

                    {/* Expandable Step Body */}
                    {isExpanded && (
                      <div className={styles.stepBody}>
                        {/* Section 1: Stage Identity & Titles */}
                        <div className={styles.twoColRow} style={{ marginBottom: 0 }}>
                          <div>
                            <Label size="small" weight="semibold">Stage label</Label>
                            <Input
                              size="medium"
                              value={step.stageNumber || ''}
                              placeholder="e.g. Stage 1"
                              onChange={(e, data) => {
                                const updated = [...steps];
                                updated[idx] = { ...updated[idx], stageNumber: data.value };
                                handleFieldChange('processSteps', updated);
                              }}
                            />
                            <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>Displayed as the top step pill</Caption1>
                          </div>
                          <div>
                            <Label size="small" weight="semibold">Stage title</Label>
                            <Input
                              size="medium"
                              value={step.title || ''}
                              placeholder="e.g. Shape"
                              onChange={(e, data) => {
                                const updated = [...steps];
                                updated[idx] = { ...updated[idx], title: data.value };
                                handleFieldChange('processSteps', updated);
                              }}
                            />
                            <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>Main headline for this milestone</Caption1>
                          </div>
                        </div>

                        {/* Section 2: Description & Case for acting */}
                        <div>
                          <Label size="small" weight="semibold">Description / Case for acting</Label>
                          <Textarea
                            rows={2}
                            size="medium"
                            value={step.description || ''}
                            placeholder="e.g. Define the problem and the case for acting..."
                            onChange={(e, data) => {
                              const updated = [...steps];
                              updated[idx] = { ...updated[idx], description: data.value };
                              handleFieldChange('processSteps', updated);
                            }}
                          />
                        </div>

                        {/* Section 3: Capabilities & Interactivity */}
                        <div className={styles.stepSectionDivider}>
                          <span>Interactivity & Action</span>
                          <Divider style={{ flexGrow: 1 }} />
                        </div>

                        <div className={styles.threeColRow} style={{ marginBottom: 0 }}>
                          <div>
                            <Label size="small" weight="semibold">Capabilities badge</Label>
                            <Input
                              size="medium"
                              value={step.metricBadge || ''}
                              placeholder="e.g. 4 capabilities"
                              onChange={(e, data) => {
                                const updated = [...steps];
                                updated[idx] = { ...updated[idx], metricBadge: data.value };
                                handleFieldChange('processSteps', updated);
                              }}
                            />
                          </div>
                          <div>
                            <Label size="small" weight="semibold">Action on click</Label>
                            <Dropdown
                              size="medium"
                              value={step.actionType === 'navigate' ? 'Follow URL link' : 'Filter Cards'}
                              onOptionSelect={(e, data) => {
                                const updated = [...steps];
                                updated[idx] = {
                                  ...updated[idx],
                                  actionType: data.optionValue as 'filter' | 'navigate'
                                };
                                handleFieldChange('processSteps', updated);
                              }}
                            >
                              <Option value="filter">Filter Cards</Option>
                              <Option value="navigate">Follow URL link</Option>
                            </Dropdown>
                          </div>
                          <div>
                            {step.actionType === 'navigate' ? (
                              <>
                                <Label size="small" weight="semibold">Destination URL</Label>
                                <Input
                                  size="medium"
                                  value={step.url || ''}
                                  placeholder="https://..."
                                  onChange={(e, data) => {
                                    const updated = [...steps];
                                    updated[idx] = { ...updated[idx], url: data.value };
                                    handleFieldChange('processSteps', updated);
                                  }}
                                />
                              </>
                            ) : (
                              <>
                                <Label size="small" weight="semibold">Custom filter term</Label>
                                <Input
                                  size="medium"
                                  value={step.filterValue || ''}
                                  placeholder="Defaults to stage title"
                                  onChange={(e, data) => {
                                    const updated = [...steps];
                                    updated[idx] = { ...updated[idx], filterValue: data.value };
                                    handleFieldChange('processSteps', updated);
                                  }}
                                />
                              </>
                            )}
                          </div>
                        </div>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </>
        );

      case 'text':
      default:
        return (
          <>
            <div className={styles.fieldRow}>
              <Label weight="semibold">Text content</Label>
              <Textarea
                value={formData.text || ''}
                placeholder="Enter rich paragraph text..."
                rows={5}
                onChange={(e, data) => handleFieldChange('text', data.value)}
              />
              <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                You can also format this text inline directly on the card canvas using the floating toolbar.
              </Caption1>
            </div>
          </>
        );
    }
  };

  return (
    <>
      <Portal>
        <Dialog open={isOpen} onOpenChange={(_, data) => !data.open && onDismiss()}>
          <DialogSurface className={styles.surface}>
            <div className={styles.headerRow}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <span style={{ fontSize: '20px', color: tokens.colorBrandForeground1, display: 'flex' }}>
                  {getItemIcon(formData.type)}
                </span>
                <DialogTitle>Configure {formData.type.toUpperCase()} Item</DialogTitle>
              </div>
              <Button
                appearance="subtle"
                icon={<DismissRegular />}
                onClick={onDismiss}
                aria-label="Close"
              />
            </div>

            <DialogBody>
              <DialogContent>
                {/* Alignment Selector across composable items - Icon buttons */}
                <div className={styles.fieldRow} style={{ marginBottom: '16px' }}>
                  <Label weight="semibold">Content alignment</Label>
                  <div style={{ display: 'inline-flex', gap: '4px', backgroundColor: tokens.colorNeutralBackground3, padding: '3px', borderRadius: tokens.borderRadiusMedium, width: 'fit-content' }}>
                    <Button
                      size="small"
                      appearance={formData.alignment === 'left' || !formData.alignment ? 'primary' : 'subtle'}
                      icon={<TextAlignLeftRegular />}
                      onClick={() => handleFieldChange('alignment', 'left')}
                      title="Align left"
                    >
                      Left
                    </Button>
                    <Button
                      size="small"
                      appearance={formData.alignment === 'center' ? 'primary' : 'subtle'}
                      icon={<TextAlignCenterRegular />}
                      onClick={() => handleFieldChange('alignment', 'center')}
                      title="Align center"
                    >
                      Centre
                    </Button>
                    <Button
                      size="small"
                      appearance={formData.alignment === 'right' ? 'primary' : 'subtle'}
                      icon={<TextAlignRightRegular />}
                      onClick={() => handleFieldChange('alignment', 'right')}
                      title="Align right"
                    >
                      Right
                    </Button>
                  </div>
                  <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                    Align this content block to the left, centre, or right of the container.
                  </Caption1>
                </div>

                {renderTypeSpecificFields()}
              </DialogContent>

              <DialogActions style={{ marginTop: '16px' }}>
                <Button appearance="secondary" onClick={onDismiss}>
                  Cancel
                </Button>
                <Button
                  appearance="primary"
                  icon={<SaveRegular />}
                  onClick={() => {
                    onSave(formData);
                    onDismiss();
                  }}
                >
                  Apply Changes
                </Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      </Portal>

      {/* Pure Fluent UI 2 Site Asset Explorer Dialog */}
      {assetPickerService && (
        <FluentAssetExplorerDialog
          isOpen={isFluentExplorerOpen}
          assetService={assetPickerService}
          title={`Select ${explorerMode.type === 'video' ? 'Video' : 'Image'} Asset`}
          itemType={explorerMode.type === 'video' ? 'video' : 'image'}
          allowMultiple={explorerMode.allowMultiple}
          acceptedExtensions={
            explorerMode.type === 'video'
              ? ['.mp4', '.webm', '.mov']
              : ['.png', '.jpg', '.jpeg', '.gif', '.svg', '.webp']
          }
          onDismiss={() => setIsFluentExplorerOpen(false)}
          onSelect={(results) => {
            applySelectedAssets(explorerMode.type, results);
            setIsFluentExplorerOpen(false);
          }}
        />
      )}
    </>
  );
};
