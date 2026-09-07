/**
 * @file IAssetPickerService.ts
 * @description Strongly typed interface for SharePoint asset & file picking services.
 * Decouples presentational Fluent UI 2 components from WebPartContext.
 */

export interface IFilePickerResult {
  fileName: string;
  fileAbsoluteUrl: string;
  serverRelativeUrl: string;
  thumbnailUrl?: string;
  fileType?: string;
  fileSize?: number;
}

export interface IAssetPickerOptions {
  /**
   * Title displayed in the file picker header.
   */
  title?: string;
  /**
   * Accepted file extensions, e.g. ['.png', '.jpg', '.jpeg', '.svg', '.gif'] or ['.mp4', '.webm']
   */
  acceptedExtensions?: string[];
  /**
   * Type of item being picked.
   */
  itemType?: 'image' | 'video' | 'link' | 'document' | 'all';
  /**
   * Whether multi-selection is enabled (e.g. for image galleries).
   */
  allowMultiple?: boolean;
}

export interface ISiteLibraryItem {
  id: string;
  name: string;
  serverRelativeUrl: string;
  absoluteUrl: string;
  isFolder: boolean;
  fileType?: string;
  size?: number;
  timeLastModified?: string;
  thumbnailUrl?: string;
}

export interface IAssetPickerService {
  /**
   * Open the native SharePoint / OneDrive File Picker experience (ODSP FilePicker SDK or SharePoint FilePicker).
   */
  openNativeFilePicker(options?: IAssetPickerOptions): Promise<IFilePickerResult[] | null>;

  /**
   * Browse document libraries and files within the current SharePoint site (Pure Fluent UI 2 fallback).
   */
  listSiteLibraries(): Promise<Array<{ title: string; serverRelativeUrl: string }>>;

  /**
   * List files and folders within a given site library path.
   */
  listLibraryContents(folderServerRelativeUrl: string, extensions?: string[]): Promise<ISiteLibraryItem[]>;

  /**
   * Direct HTML5 upload of a file into the site's 'Site Assets' or target library.
   */
  uploadFile(targetFolderRelativeUrl: string, file: File): Promise<IFilePickerResult>;
}
