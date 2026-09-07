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
  Button,
  Input,
  Label,
  Textarea,
  Dropdown,
  Option,
  Divider,
  Caption1,
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
  FolderOpenRegular
} from '@fluentui/react-icons';
import { ICardItem, ICardItemType } from '../models/IContainerModels';
import { TermStorePicker } from './TermStorePicker';
import { IAssetPickerService, IFilePickerResult } from '../services/IAssetPickerService';
import { FluentAssetExplorerDialog } from './FluentAssetExplorerDialog';

const useStyles = makeStyles({
  surface: {
    maxWidth: '540px',
    width: '94vw',
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    ...shorthands.padding('20px')
  },
  headerRow: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: '16px'
  },
  fieldRow: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap(tokens.spacingVerticalXXS),
    marginBottom: '12px'
  },
  twoColRow: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    ...shorthands.gap(tokens.spacingHorizontalM),
    marginBottom: '12px'
  },
  listContainer: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
    marginTop: '6px'
  },
  listItemRow: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr auto',
    gap: '6px',
    alignItems: 'center',
    padding: '6px',
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
  const [isFluentExplorerOpen, setIsFluentExplorerOpen] = React.useState<boolean>(false);
  const [explorerMode, setExplorerMode] = React.useState<{
    type: 'image' | 'video' | 'gallery' | 'hero' | 'cta' | 'editorial';
    allowMultiple: boolean;
  }>({ type: 'image', allowMultiple: false });

  React.useEffect(() => {
    if (item) {
      setFormData(JSON.parse(JSON.stringify(item)));
    } else {
      setFormData(null);
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
