/**
 * @file DashboardStorageService.ts
 * @description Service for backing up and restoring dashboard configurations to a custom SharePoint
 * Document Library called 'Dashboards'. Manages folder hierarchies ('Templates' and 'Backups' with
 * dashboard subfolders), automated version control, cross-site JSON portability, and offline mock fallback.
 */

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IContainerSection, LayoutMode, ContainerStyle } from '../models/IContainerModels';

/**
 * Self-contained cross-site portable dashboard package schema.
 */
export interface IDashboardPackage {
  schemaVersion: '1.0.0';
  exportedAt: string;
  exportedBy?: string;
  sourceSiteUrl?: string;
  dashboardTitle: string;
  dashboardSubtitle?: string;
  layoutMode: LayoutMode;
  containerStyle?: ContainerStyle;
  gridColumns?: number;
  gridRows?: number;
  cardHeightMode?: string;
  webPartBackgroundColor?: string;
  sectionBackgroundColor?: string;
  sections: IContainerSection[];
  metadata?: {
    totalSections: number;
    totalCards: number;
    totalFinancialMetrics: number;
    tagsUsed: string[];
    trigger: 'manual' | 'auto';
    folderType: 'Backups' | 'Templates';
  };
}

/**
 * Information about a saved dashboard file in the library.
 */
export interface IBackupFileInfo {
  name: string;
  serverRelativeUrl: string;
  timeLastModified: string;
  length?: number;
  dashboardTitle: string;
  folderType: 'Backups' | 'Templates';
  subFolder: string;
  packageData?: IDashboardPackage;
}

/**
 * Options for saving a dashboard backup snapshot.
 */
export interface ISaveBackupOptions {
  folderType: 'Backups' | 'Templates';
  dashboardTitle: string;
  dashboardSubtitle?: string;
  layoutMode: LayoutMode;
  containerStyle?: ContainerStyle;
  gridColumns?: number;
  gridRows?: number;
  cardHeightMode?: string;
  webPartBackgroundColor?: string;
  sectionBackgroundColor?: string;
  sections: IContainerSection[];
  trigger?: 'manual' | 'auto';
  customFileName?: string;
  /**
   * Optional callback function invoked when the 'Dashboards' library does not exist.
   * Prompts the user whether to provision the library. Returning true provisions it; false cancels.
   */
  promptLibraryCreation?: (libraryName: string) => Promise<boolean> | boolean;
}

/**
 * Result returned upon saving a dashboard snapshot.
 */
export interface ISaveBackupResult {
  success: boolean;
  fileName: string;
  folderPath: string;
  serverRelativeUrl: string;
  message: string;
  savedAt: string;
  versionId?: string;
}

export class DashboardStorageService {
  private _spHttpClient?: SPHttpClient;
  private _siteServerRelativeUrl: string;
  private _libraryTitle: string = 'Dashboards';
  private static _localStorageKey: string = 'spfx_dashboards_library';

  /**
   * Initializes the storage service.
   * @param spHttpClient Optional SPFx SPHttpClient instance for live SharePoint Online environments.
   * @param siteServerRelativeUrl Server-relative URL of the current SharePoint site.
   * @param libraryTitle Name of the custom SharePoint document library (default: 'Dashboards').
   */
  constructor(spHttpClient?: SPHttpClient, siteServerRelativeUrl: string = '', libraryTitle: string = 'Dashboards') {
    this._spHttpClient = spHttpClient;
    this._siteServerRelativeUrl = siteServerRelativeUrl.replace(/\/+$/, '');
    this._libraryTitle = libraryTitle || 'Dashboards';
  }

  /**
   * Sanitizes a folder or file name by replacing illegal SharePoint characters.
   */
  public sanitizeName(name: string): string {
    if (!name || !name.trim()) return 'Untitled_Dashboard';
    return name
      .trim()
      .replace(/[\\/:*?"<>|#%~&{}]/g, '_')
      .replace(/\s+/g, ' ');
  }

  /**
   * Generates a timestamped version filename for the dashboard.
   */
  public generateVersionFileName(dashboardTitle: string, trigger: 'manual' | 'auto' = 'manual'): string {
    const safeTitle = this.sanitizeName(dashboardTitle).replace(/\s+/g, '_');
    const now = new Date();
    const pad = (n: number): string => (n < 10 ? `0${n}` : `${n}`);
    const dateStr = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}`;
    const timeStr = `${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
    const suffix = trigger === 'auto' ? '_autosave' : '';
    return `${safeTitle}_${dateStr}_${timeStr}${suffix}.json`;
  }

  /**
   * Assembles a 100% self-contained portable dashboard package.
   */
  public createPackage(options: ISaveBackupOptions): IDashboardPackage {
    const sections = options.sections || [];
    let totalCards = 0;
    let totalFinancialMetrics = 0;
    const tagsSet = new Set<string>();

    sections.forEach((sec) => {
      (sec.blocks || []).forEach((blk) => {
        totalCards++;
        if (blk.type === 'metric' || blk.metricValue) {
          totalFinancialMetrics++;
        }
        if (blk.tags) {
          blk.tags.forEach((t) => tagsSet.add(t));
        }
        if (blk.termStoreTags) {
          blk.termStoreTags.forEach((t) => tagsSet.add(t.label));
        }
      });
    });

    return {
      schemaVersion: '1.0.0',
      exportedAt: new Date().toISOString(),
      sourceSiteUrl: this._siteServerRelativeUrl,
      dashboardTitle: options.dashboardTitle || 'Dashboard title',
      dashboardSubtitle: options.dashboardSubtitle || '',
      layoutMode: options.layoutMode || 'tabs',
      containerStyle: options.containerStyle || 'standard',
      gridColumns: options.gridColumns !== undefined ? options.gridColumns : 0,
      gridRows: options.gridRows !== undefined ? options.gridRows : 0,
      cardHeightMode: options.cardHeightMode || 'default',
      webPartBackgroundColor: options.webPartBackgroundColor || '',
      sectionBackgroundColor: options.sectionBackgroundColor || '',
      sections: JSON.parse(JSON.stringify(sections)),
      metadata: {
        totalSections: sections.length,
        totalCards,
        totalFinancialMetrics,
        tagsUsed: Array.from(tagsSet),
        trigger: options.trigger || 'manual',
        folderType: options.folderType
      }
    };
  }

  /**
   * Encodes special characters in a SharePoint path (like single quotes and hashes)
   * while strictly preserving path delimiters ('/').
   */
  public escapeSpPath(path: string): string {
    if (!path) return '';
    return path.replace(/'/g, "''");
  }

  /**
   * Saves a dashboard JSON snapshot to the target folder in the 'Dashboards' library.
   * Path: Dashboards/{folderType}/{DashboardTitle}/{FileName}.json
   */
  public async saveDashboardBackup(options: ISaveBackupOptions): Promise<ISaveBackupResult> {
    const safeDashboardFolder = this.sanitizeName(options.dashboardTitle);
    const fileName = options.customFileName || this.generateVersionFileName(options.dashboardTitle, options.trigger);
    const pkg = this.createPackage(options);
    const jsonContent = JSON.stringify(pkg, null, 2);

    const targetFolderRelPath = `${this._libraryTitle}/${options.folderType}/${safeDashboardFolder}`;
    const fullServerRelFolderPath = this._siteServerRelativeUrl
      ? `${this._siteServerRelativeUrl}/${targetFolderRelPath}`
      : `/${targetFolderRelPath}`;
    const fullFileServerRelUrl = `${fullServerRelFolderPath}/${fileName}`;

    if (this._spHttpClient && this._siteServerRelativeUrl) {
      // Ensure the full folder path exists: Dashboards/Backups|Templates/DashboardTitle
      await this._ensureSharePointFolderPath(targetFolderRelPath, options.promptLibraryCreation);

      const escapedFolderPath = this.escapeSpPath(fullServerRelFolderPath);
      // Files/Add requires raw body — do NOT set Content-Type to application/json.
      // SharePoint REST treats the body as octet-stream and will corrupt the content otherwise.
      const addFileEndpoint = `${this._siteServerRelativeUrl}/_api/web/GetFolderByServerRelativeUrl('${escapedFolderPath}')/Files/Add(url='${encodeURIComponent(fileName)}', overwrite=true)`;

      console.log(`[DashboardStorageService] Uploading to: ${addFileEndpoint}`);

      const response: SPHttpClientResponse = await this._spHttpClient.post(
        addFileEndpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata'
            // Note: Content-Type intentionally omitted — SPHttpClient sends as octet-stream
          },
          body: jsonContent
        }
      );

      if (!response.ok) {
        const errText = await response.text();
        throw new Error(`SharePoint file upload HTTP ${response.status}: ${errText}`);
      }

      console.log(`[DashboardStorageService] Successfully saved: ${fullFileServerRelUrl}`);
      return {
        success: true,
        fileName,
        folderPath: fullServerRelFolderPath,
        serverRelativeUrl: fullFileServerRelUrl,
        message: `Saved snapshot to SharePoint: ${this._libraryTitle}/${options.folderType}/${safeDashboardFolder}/${fileName}`,
        savedAt: new Date().toISOString()
      };
    } else {
      return this._saveToLocalMock(options.folderType, safeDashboardFolder, fileName, fullFileServerRelUrl, pkg);
    }
  }

  /**
   * Lists all saved dashboard snapshots from the specified folder.
   */
  public async listDashboardBackups(
    folderType: 'Backups' | 'Templates',
    dashboardTitle?: string
  ): Promise<IBackupFileInfo[]> {
    const safeDashboardFolder = dashboardTitle ? this.sanitizeName(dashboardTitle) : undefined;
    const targetFolderRelPath = safeDashboardFolder
      ? `${this._libraryTitle}/${folderType}/${safeDashboardFolder}`
      : `${this._libraryTitle}/${folderType}`;

    const fullServerRelFolderPath = this._siteServerRelativeUrl
      ? `${this._siteServerRelativeUrl}/${targetFolderRelPath}`
      : `/${targetFolderRelPath}`;

    if (this._spHttpClient && this._siteServerRelativeUrl) {
      try {
        const escapedFolderPath = this.escapeSpPath(fullServerRelFolderPath);
        const queryUrl = `${this._siteServerRelativeUrl}/_api/web/GetFolderByServerRelativeUrl('${escapedFolderPath}')/Files?$select=Name,ServerRelativeUrl,TimeLastModified,Length&$orderby=TimeLastModified desc`;

        const response = await this._spHttpClient.get(queryUrl, SPHttpClient.configurations.v1, {
          headers: { Accept: 'application/json;odata=nometadata' }
        });

        if (response.ok) {
          const data = await response.json();
          const files = data.value || [];
          return files
            .filter((f: { Name: string }) => f.Name.toLowerCase().endsWith('.json'))
            .map((f: { Name: string; ServerRelativeUrl: string; TimeLastModified: string; Length: number }) => ({
              name: f.Name,
              serverRelativeUrl: f.ServerRelativeUrl,
              timeLastModified: f.TimeLastModified,
              length: f.Length,
              dashboardTitle: safeDashboardFolder || 'Dashboard',
              folderType,
              subFolder: safeDashboardFolder || ''
            }));
        }
      } catch (err) {
        console.warn('[DashboardStorageService] Failed to query SharePoint files, reading from local fallback:', err);
      }
    }

    return this._listFromLocalMock(folderType, safeDashboardFolder);
  }

  /**
   * Loads and deserializes a dashboard package from a file URL.
   */
  public async loadDashboardBackup(serverRelativeFileUrl: string): Promise<IDashboardPackage> {
    if (this._spHttpClient && this._siteServerRelativeUrl) {
      try {
        const escapedFileUrl = this.escapeSpPath(serverRelativeFileUrl);
        const getFileEndpoint = `${this._siteServerRelativeUrl}/_api/web/GetFileByServerRelativeUrl('${escapedFileUrl}')/$value`;

        const response = await this._spHttpClient.get(getFileEndpoint, SPHttpClient.configurations.v1, {
          headers: { Accept: 'application/json' }
        });

        if (response.ok) {
          const jsonText = await response.text();
          return JSON.parse(jsonText) as IDashboardPackage;
        }
      } catch (err) {
        console.warn('[DashboardStorageService] SharePoint file load failed, checking local storage:', err);
      }
    }

    return this._loadFromLocalMock(serverRelativeFileUrl);
  }

  /**
   * Checks if the target Document Library exists on the SharePoint site.
   * If missing, prompts the user (or uses promptCallback) to confirm creation before provisioning.
   * Automatically provisions root 'Backups' and 'Templates' folders if created.
   */
  private async _ensureDocumentLibraryExists(promptLibraryCreation?: (libraryName: string) => Promise<boolean> | boolean): Promise<void> {
    if (!this._spHttpClient || !this._siteServerRelativeUrl) return;

    const checkEndpoint = `${this._siteServerRelativeUrl}/_api/web/lists/getByTitle('${encodeURIComponent(this._libraryTitle)}')?$select=Id,Title`;
    const res = await this._spHttpClient.get(checkEndpoint, SPHttpClient.configurations.v1, {
      headers: { Accept: 'application/json;odata=nometadata' }
    });

    if (res.ok) {
      return; // Document library already exists
    }

    // If library does not exist (404), prompt the user before provisioning
    if (res.status === 404) {
      let shouldCreate = false;
      if (typeof promptLibraryCreation === 'function') {
        shouldCreate = await Promise.resolve(promptLibraryCreation(this._libraryTitle));
      } else if (typeof window !== 'undefined' && typeof window.confirm === 'function') {
        shouldCreate = window.confirm(
          `The '${this._libraryTitle}' document library was not found on this SharePoint site.\n\n` +
          `Would you like to create the '${this._libraryTitle}' library now with 'Backups' and 'Templates' folders to store your dashboard configurations?`
        );
      } else {
        shouldCreate = true;
      }

      if (!shouldCreate) {
        throw new Error(`Creation of '${this._libraryTitle}' document library was cancelled by user.`);
      }

      console.log(`[DashboardStorageService] User confirmed. Provisioning Document Library '${this._libraryTitle}'...`);
      const createEndpoint = `${this._siteServerRelativeUrl}/_api/web/lists`;
      const createRes = await this._spHttpClient.post(createEndpoint, SPHttpClient.configurations.v1, {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=verbose'
        },
        body: JSON.stringify({
          '__metadata': { 'type': 'SP.List' },
          'BaseTemplate': 101,
          'Description': 'Storage repository for Full-Width Dashboard backups, templates, and snapshot configurations.',
          'Title': this._libraryTitle
        })
      });

      if (!createRes.ok) {
        const errText = await createRes.text();
        throw new Error(`Failed to create Document Library '${this._libraryTitle}': ${errText}`);
      }

      console.log(`[DashboardStorageService] Successfully provisioned Document Library '${this._libraryTitle}'. Initialising root folders...`);
      // Automatically ensure root 'Backups' and 'Templates' folders exist in the newly provisioned library
      await this._ensureRootLibraryFolder('Backups');
      await this._ensureRootLibraryFolder('Templates');
    }
  }

  /**
   * Helper to ensure a root folder (e.g. Backups, Templates) exists directly under the Dashboards library.
   * Sends the full server-relative path in the body as required by the SharePoint REST /folders endpoint.
   */
  private async _ensureRootLibraryFolder(folderName: string): Promise<void> {
    if (!this._spHttpClient || !this._siteServerRelativeUrl) return;
    const parentFullServerRel = `${this._siteServerRelativeUrl}/${this._libraryTitle}`;
    const newFolderFullPath = `${parentFullServerRel}/${folderName}`;
    try {
      const escapedParentUrl = this.escapeSpPath(parentFullServerRel);
      const endpoint = `${this._siteServerRelativeUrl}/_api/web/GetFolderByServerRelativeUrl('${escapedParentUrl}')/folders`;
      await this._spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata'
        },
        body: JSON.stringify({ ServerRelativeUrl: newFolderFullPath })
      });
      console.log(`[DashboardStorageService] Root folder ensured: ${newFolderFullPath}`);
    } catch {
      // Ignore — folder likely already exists
    }
  }


  /**
   * Iteratively creates folders in the SharePoint Document Library if they don't already exist.
   * Traverses from library root down into subfolders ('Backups'/'Templates' -> DashboardTitle).
   * Uses the correct SharePoint REST folder-add body: { ServerRelativeUrl: fullServerRelativePath }.
   */
  private async _ensureSharePointFolderPath(
    relativeFolderPath: string,
    promptLibraryCreation?: (libraryName: string) => Promise<boolean> | boolean
  ): Promise<void> {
    if (!this._spHttpClient || !this._siteServerRelativeUrl) return;

    // Ensure library itself exists first (with prompt if missing)
    await this._ensureDocumentLibraryExists(promptLibraryCreation);

    const segments = relativeFolderPath.split('/').filter(Boolean);
    if (segments.length === 0) return;

    // segments[0] is the library name — already created above.
    // Iterate through sub-folders: e.g. ['Dashboards', 'Backups', 'My_Dashboard_Title']
    let currentFullServerRelPath = `${this._siteServerRelativeUrl}/${segments[0]}`;

    for (let i = 1; i < segments.length; i++) {
      const segment = segments[i];
      const targetFullPath = `${currentFullServerRelPath}/${segment}`;

      // Check if folder exists first
      const checkUrl = `${this._siteServerRelativeUrl}/_api/web/GetFolderByServerRelativeUrl('${this.escapeSpPath(targetFullPath)}')`;
      try {
        const checkRes = await this._spHttpClient.get(checkUrl, SPHttpClient.configurations.v1, {
          headers: { Accept: 'application/json;odata=nometadata' }
        });
        if (checkRes.ok) {
          // Folder already exists — move on
          currentFullServerRelPath = targetFullPath;
          continue;
        }
      } catch {
        // Treat as missing — attempt to create
      }

      // Create the folder using the parent folder's /folders endpoint
      // Body must contain the FULL server-relative URL of the new folder
      const createEndpoint = `${this._siteServerRelativeUrl}/_api/web/GetFolderByServerRelativeUrl('${this.escapeSpPath(currentFullServerRelPath)}')/folders`;
      try {
        const createRes = await this._spHttpClient.post(createEndpoint, SPHttpClient.configurations.v1, {
          headers: {
            Accept: 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata'
          },
          body: JSON.stringify({ ServerRelativeUrl: targetFullPath })
        });
        if (!createRes.ok && createRes.status !== 409) {
          const errText = await createRes.text();
          console.warn(`[DashboardStorageService] Folder create returned ${createRes.status}: ${errText}`);
        } else {
          console.log(`[DashboardStorageService] Folder ensured: ${targetFullPath}`);
        }
      } catch (folderErr) {
        console.warn(`[DashboardStorageService] Could not create folder ${targetFullPath}:`, folderErr);
      }

      currentFullServerRelPath = targetFullPath;
    }
  }

  /**
   * LocalStorage Mock: Saves snapshot locally.
   */
  private _saveToLocalMock(
    folderType: 'Backups' | 'Templates',
    subFolder: string,
    fileName: string,
    serverRelativeUrl: string,
    pkg: IDashboardPackage
  ): ISaveBackupResult {
    try {
      const raw = localStorage.getItem(DashboardStorageService._localStorageKey);
      const store = raw ? JSON.parse(raw) : { Backups: {}, Templates: {} };
      if (!store[folderType]) store[folderType] = {};
      if (!store[folderType][subFolder]) store[folderType][subFolder] = [];

      store[folderType][subFolder].unshift({
        name: fileName,
        serverRelativeUrl,
        timeLastModified: new Date().toISOString(),
        dashboardTitle: pkg.dashboardTitle,
        folderType,
        subFolder,
        packageData: pkg
      });

      localStorage.setItem(DashboardStorageService._localStorageKey, JSON.stringify(store));
    } catch {
      // Ignore quota errors
    }

    return {
      success: true,
      fileName,
      folderPath: `${this._libraryTitle}/${folderType}/${subFolder}`,
      serverRelativeUrl,
      message: `Saved snapshot to ${this._libraryTitle}/${folderType}/${subFolder}/${fileName}`,
      savedAt: new Date().toISOString()
    };
  }

  /**
   * LocalStorage Mock: Lists snapshots locally.
   */
  private _listFromLocalMock(folderType: 'Backups' | 'Templates', subFolder?: string): IBackupFileInfo[] {
    try {
      const raw = localStorage.getItem(DashboardStorageService._localStorageKey);
      if (!raw) return [];
      const store = JSON.parse(raw);
      const typeBucket = store[folderType] || {};

      if (subFolder) {
        return (typeBucket[subFolder] || []).map((item: IBackupFileInfo) => ({
          name: item.name,
          serverRelativeUrl: item.serverRelativeUrl,
          timeLastModified: item.timeLastModified,
          length: JSON.stringify(item.packageData || {}).length,
          dashboardTitle: item.dashboardTitle || subFolder,
          folderType,
          subFolder
        }));
      }

      const allFiles: IBackupFileInfo[] = [];
      Object.keys(typeBucket).forEach((sf) => {
        (typeBucket[sf] || []).forEach((item: IBackupFileInfo) => {
          allFiles.push({
            name: item.name,
            serverRelativeUrl: item.serverRelativeUrl,
            timeLastModified: item.timeLastModified,
            length: JSON.stringify(item.packageData || {}).length,
            dashboardTitle: item.dashboardTitle || sf,
            folderType,
            subFolder: sf
          });
        });
      });
      return allFiles.sort(
        (a, b) => new Date(b.timeLastModified).getTime() - new Date(a.timeLastModified).getTime()
      );
    } catch {
      return [];
    }
  }

  /**
   * LocalStorage Mock: Loads a package locally.
   */
  private _loadFromLocalMock(serverRelativeFileUrl: string): IDashboardPackage {
    const raw = localStorage.getItem(DashboardStorageService._localStorageKey);
    if (raw) {
      const store = JSON.parse(raw);
      for (const fType of ['Backups', 'Templates'] as const) {
        const typeBucket = store[fType] || {};
        for (const sub of Object.keys(typeBucket)) {
          const found = (typeBucket[sub] || []).find(
            (f: IBackupFileInfo) => f.serverRelativeUrl === serverRelativeFileUrl || f.name === serverRelativeFileUrl
          );
          if (found && found.packageData) {
            return found.packageData;
          }
        }
      }
    }
    throw new Error(`Dashboard backup file not found: ${serverRelativeFileUrl}`);
  }
}
