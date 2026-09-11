/**
 * @file FluentAssetExplorerDialog.tsx
 * @description Pure Fluent UI 2 (v9) modal for selecting assets across SharePoint Document Libraries,
 * uploading new local files, or entering external URLs.
 * Renders cleanly in a Fluent UI 2 <Dialog> mounted in a <Portal> with zero legacy v8 dependencies.
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
  TabList,
  Tab,
  SelectTabData,
  SelectTabEvent,
  makeStyles,
  tokens,
  shorthands,
  Caption1,
  Spinner,
  Badge
} from '@fluentui/react-components';
import {
  DismissRegular,
  FolderRegular,
  DocumentRegular,
  ImageRegular,
  VideoRegular,
  ArrowUploadRegular,
  GlobeRegular,
  CheckmarkRegular,
  ArrowLeftRegular,
  SearchRegular,
  SparkleRegular
} from '@fluentui/react-icons';
import {
  IAssetPickerService,
  IFilePickerResult,
  ISiteLibraryItem
} from '../services/IAssetPickerService';

const useStyles = makeStyles({
  surface: {
    maxWidth: '840px',
    width: '92vw',
    height: '80vh',
    maxHeight: '680px',
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    ...shorthands.padding('20px'),
    display: 'flex',
    flexDirection: 'column'
  },
  header: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: '12px'
  },
  tabList: {
    marginBottom: '16px',
    ...shorthands.borderBottom('1px', 'solid', tokens.colorNeutralStroke2)
  },
  contentArea: {
    flex: 1,
    overflowY: 'auto',
    display: 'flex',
    flexDirection: 'column',
    // ...shorthands.gap(tokens.spacingVerticalM) // [USER_TEST: comment out vertical spacing in flex containers]
  },
  grid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fill, minmax(170px, 1fr))',
    gap: '12px',
    padding: '4px'
  },
  itemTile: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorNeutralBackground1,
    cursor: 'pointer',
    position: 'relative',
    overflow: 'hidden',
    textAlign: 'center',
    transition: 'all 0.15s ease',
    ':hover': {
      ...shorthands.borderColor(tokens.colorBrandStroke1),
      transform: 'translateY(-2px)',
      boxShadow: tokens.shadow4
    }
  },
  itemTileSelected: {
    ...shorthands.borderColor(tokens.colorBrandStroke1),
    backgroundColor: tokens.colorBrandBackground2,
    boxShadow: `0 0 0 2px ${tokens.colorBrandStroke1}`
  },
  thumbnailContainer: {
    width: '100%',
    height: '110px',
    backgroundColor: tokens.colorNeutralBackground2,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    overflow: 'hidden'
  },
  thumbnailImg: {
    width: '100%',
    height: '100%',
    objectFit: 'cover'
  },
  itemInfo: {
    padding: '8px',
    width: '100%',
    boxSizing: 'border-box'
  },
  itemName: {
    fontSize: '0.8rem',
    fontWeight: 600,
    color: tokens.colorNeutralForeground1,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    display: 'block'
  },
  uploadDropzone: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    minHeight: '220px',
    ...shorthands.border('2px', 'dashed', tokens.colorBrandStroke2),
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    backgroundColor: tokens.colorNeutralBackground2,
    padding: '24px',
    textAlign: 'center',
    cursor: 'pointer',
    transition: 'all 0.2s ease',
    ':hover': {
      backgroundColor: tokens.colorBrandBackground2,
      ...shorthands.borderColor(tokens.colorBrandStroke1)
    }
  }
});

export interface IFluentAssetExplorerDialogProps {
  isOpen: boolean;
  assetService: IAssetPickerService;
  title?: string;
  itemType?: 'image' | 'video' | 'link' | 'document' | 'all';
  acceptedExtensions?: string[];
  allowMultiple?: boolean;
  onDismiss: () => void;
  onSelect: (results: IFilePickerResult[]) => void;
}

export const FluentAssetExplorerDialog: React.FC<IFluentAssetExplorerDialogProps> = ({
  isOpen,
  assetService,
  title = 'Select Asset',
  itemType = 'image',
  acceptedExtensions,
  allowMultiple = false,
  onDismiss,
  onSelect
}) => {
  const styles = useStyles();

  const [activeTab, setActiveTab] = React.useState<'site' | 'upload' | 'url'>('site');
  const [libraries, setLibraries] = React.useState<Array<{ title: string; serverRelativeUrl: string }>>([]);
  const [currentFolder, setCurrentFolder] = React.useState<string | null>(null);
  const [folderHistory, setFolderHistory] = React.useState<string[]>([]);
  const [items, setItems] = React.useState<ISiteLibraryItem[]>([]);
  const [selectedItems, setSelectedItems] = React.useState<ISiteLibraryItem[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [searchQuery, setSearchQuery] = React.useState<string>('');
  const [manualUrl, setManualUrl] = React.useState<string>('');
  const [isUploading, setIsUploading] = React.useState<boolean>(false);
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  // Load site libraries on open
  React.useEffect(() => {
    if (isOpen) {
      void (async () => {
        setIsLoading(true);
        try {
          const libs = await assetService.listSiteLibraries();
          setLibraries(libs);
          if (libs.length > 0) {
            // Auto open Site Assets or first library
            const siteAssets = libs.find(l => l.title.toLowerCase().includes('site assets')) || libs[0];
            setCurrentFolder(siteAssets.serverRelativeUrl);
            const contents = await assetService.listLibraryContents(siteAssets.serverRelativeUrl, acceptedExtensions);
            setItems(contents);
          }
        } finally {
          setIsLoading(false);
        }
      })();
    } else {
      setSelectedItems([]);
      setFolderHistory([]);
      setSearchQuery('');
      setManualUrl('');
    }
  }, [isOpen, assetService]);

  if (!isOpen) return null;

  const navigateToFolder = async (folderUrl: string): Promise<void> => {
    setIsLoading(true);
    try {
      if (currentFolder) {
        setFolderHistory(prev => [...prev, currentFolder]);
      }
      setCurrentFolder(folderUrl);
      const contents = await assetService.listLibraryContents(folderUrl, acceptedExtensions);
      setItems(contents);
    } finally {
      setIsLoading(false);
    }
  };

  const navigateBack = async (): Promise<void> => {
    if (folderHistory.length === 0) return;
    const prevFolder = folderHistory[folderHistory.length - 1];
    setFolderHistory(prev => prev.slice(0, prev.length - 1));
    setIsLoading(true);
    try {
      setCurrentFolder(prevFolder);
      const contents = await assetService.listLibraryContents(prevFolder, acceptedExtensions);
      setItems(contents);
    } finally {
      setIsLoading(false);
    }
  };

  const toggleSelectItem = (item: ISiteLibraryItem): void => {
    if (item.isFolder) {
      void navigateToFolder(item.serverRelativeUrl);
      return;
    }
    if (allowMultiple) {
      const exists = selectedItems.some(i => i.id === item.id);
      if (exists) {
        setSelectedItems(selectedItems.filter(i => i.id !== item.id));
      } else {
        setSelectedItems([...selectedItems, item]);
      }
    } else {
      setSelectedItems([item]);
    }
  };

  const handleApply = (): void => {
    if (activeTab === 'url') {
      if (!manualUrl.trim()) return;
      const fileName = manualUrl.split('/').pop() || 'Web resource';
      onSelect([{
        fileName,
        fileAbsoluteUrl: manualUrl,
        serverRelativeUrl: manualUrl,
        thumbnailUrl: manualUrl
      }]);
      onDismiss();
      return;
    }

    if (selectedItems.length === 0) return;

    const results: IFilePickerResult[] = selectedItems.map(i => ({
      fileName: i.name,
      fileAbsoluteUrl: i.absoluteUrl,
      serverRelativeUrl: i.serverRelativeUrl,
      thumbnailUrl: i.thumbnailUrl || i.absoluteUrl,
      fileType: i.fileType,
      fileSize: i.size
    }));
    onSelect(results);
    onDismiss();
  };

  const handleFileUpload = async (files: FileList | null): Promise<void> => {
    if (!files || files.length === 0) return;
    const targetFolder = currentFolder || libraries[0]?.serverRelativeUrl;
    if (!targetFolder) return;

    setIsUploading(true);
    try {
      const uploadedResults: IFilePickerResult[] = [];
      for (let i = 0; i < files.length; i++) {
        const res = await assetService.uploadFile(targetFolder, files[i]);
        uploadedResults.push(res);
      }
      onSelect(uploadedResults);
      onDismiss();
    } catch (err) {
      console.error('[FluentAssetExplorerDialog] Upload failed:', err);
      alert(`Upload error: ${(err as Error).message}`);
    } finally {
      setIsUploading(false);
    }
  };

  const filteredItems = items.filter(i => {
    if (!searchQuery) return true;
    return i.name.toLowerCase().includes(searchQuery.toLowerCase());
  });

  return (
    <Dialog open={isOpen} onOpenChange={(_, data) => !data.open && onDismiss()}>
      <DialogSurface className={styles.surface}>
        <div className={styles.header}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <span style={{ fontSize: '20px', color: tokens.colorBrandForeground1 }}>
              {itemType === 'video' ? <VideoRegular /> : <ImageRegular />}
            </span>
            <DialogTitle>{title}</DialogTitle>
          </div>
          <Button
            appearance="subtle"
            icon={<DismissRegular />}
            onClick={onDismiss}
            aria-label="Close"
          />
        </div>

        <TabList
          className={styles.tabList}
          selectedValue={activeTab}
          onTabSelect={(_, data) => setActiveTab(data.value as 'site' | 'upload' | 'url')}
        >
          <Tab value="site" icon={<FolderRegular />}>
            Site & Document Libraries
          </Tab>
          <Tab value="upload" icon={<ArrowUploadRegular />}>
            Upload from this device
          </Tab>
          <Tab value="url" icon={<GlobeRegular />}>
            From a web link
          </Tab>
        </TabList>

        <DialogBody className={styles.contentArea}>
          <DialogContent>
            {activeTab === 'site' && (
              <>
                {/* Explorer Toolbar */}
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px', gap: '12px', flexWrap: 'wrap' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    {folderHistory.length > 0 && (
                      <Button
                        size="small"
                        icon={<ArrowLeftRegular />}
                        appearance="subtle"
                        onClick={() => void navigateBack()}
                      >
                        Back
                      </Button>
                    )}
                    <Caption1 style={{ color: tokens.colorNeutralForeground3, fontWeight: 600 }}>
                      📁 {currentFolder ? currentFolder.split('/').pop() : 'Root'}
                    </Caption1>
                  </div>
                  <Input
                    size="small"
                    placeholder="Search folder..."
                    contentBefore={<SearchRegular />}
                    value={searchQuery}
                    onChange={(_, d) => setSearchQuery(d.value)}
                  />
                </div>

                {isLoading ? (
                  <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '200px' }}>
                    <Spinner label="Loading SharePoint contents..." />
                  </div>
                ) : filteredItems.length === 0 ? (
                  <div style={{ textAlign: 'center', padding: '40px', color: tokens.colorNeutralForeground3 }}>
                    No files found in this folder matching the required format.
                  </div>
                ) : (
                  <div className={styles.grid}>
                    {filteredItems.map(item => {
                      const isSelected = selectedItems.some(i => i.id === item.id);
                      return (
                        <div
                          key={item.id}
                          className={`${styles.itemTile} ${isSelected ? styles.itemTileSelected : ''}`}
                          onClick={() => toggleSelectItem(item)}
                        >
                          <div className={styles.thumbnailContainer}>
                            {item.isFolder ? (
                              <FolderRegular style={{ fontSize: '48px', color: tokens.colorBrandForeground1 }} />
                            ) : itemType === 'video' ? (
                              <VideoRegular style={{ fontSize: '40px', color: tokens.colorNeutralForeground2 }} />
                            ) : (
                              <img
                                src={item.thumbnailUrl || item.absoluteUrl}
                                alt={item.name}
                                className={styles.thumbnailImg}
                                onError={(e) => {
                                  // Fallback to placeholder if thumbnail endpoint fails
                                  (e.target as HTMLElement).style.display = 'none';
                                }}
                              />
                            )}
                          </div>
                          <div className={styles.itemInfo}>
                            <span className={styles.itemName} title={item.name}>
                              {item.name}
                            </span>
                          </div>
                          {isSelected && (
                            <div style={{ position: 'absolute', top: '6px', right: '6px', backgroundColor: tokens.colorBrandBackground, color: '#FFF', borderRadius: '50%', padding: '2px', display: 'flex' }}>
                              <CheckmarkRegular style={{ fontSize: '12px' }} />
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </>
            )}

            {activeTab === 'upload' && (
              <div
                className={styles.uploadDropzone}
                onClick={() => fileInputRef.current?.click()}
              >
                <input
                  type="file"
                  ref={fileInputRef}
                  style={{ display: 'none' }}
                  multiple={allowMultiple}
                  accept={acceptedExtensions?.join(',')}
                  onChange={(e) => void handleFileUpload(e.target.files)}
                />
                <ArrowUploadRegular style={{ fontSize: '48px', color: tokens.colorBrandForeground1, marginBottom: '12px' }} />
                <div style={{ fontWeight: 600, fontSize: '1rem', color: tokens.colorNeutralForeground1 }}>
                  {isUploading ? 'Uploading to SharePoint...' : 'Click or drag files here to upload'}
                </div>
                <Caption1 style={{ color: tokens.colorNeutralForeground3, marginTop: '6px' }}>
                  Target folder: {currentFolder || 'Site Assets'}
                </Caption1>
              </div>
            )}

            {activeTab === 'url' && (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '12px', padding: '12px 0' }}>
                <Label weight="semibold">Direct Web or Asset URL</Label>
                <Input
                  value={manualUrl}
                  placeholder="https://... or /sites/.../image.png"
                  onChange={(_, d) => setManualUrl(d.value)}
                />
                <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                  Paste any absolute web URL or SharePoint server-relative path.
                </Caption1>
              </div>
            )}
          </DialogContent>
        </DialogBody>

        <DialogActions style={{ marginTop: '16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div>
            {selectedItems.length > 0 && (
              <Badge appearance="filled" color="brand">
                {selectedItems.length} selected
              </Badge>
            )}
          </div>
          <div style={{ display: 'flex', gap: '8px' }}>
            <Button appearance="secondary" onClick={onDismiss}>
              Cancel
            </Button>
            <Button
              appearance="primary"
              disabled={activeTab === 'url' ? !manualUrl.trim() : selectedItems.length === 0}
              onClick={handleApply}
            >
              Select & Insert
            </Button>
          </div>
        </DialogActions>
      </DialogSurface>
    </Dialog>
  );
};
