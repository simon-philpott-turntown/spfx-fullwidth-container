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
      try {
        await this._ensureSharePointFolderPath(targetFolderRelPath);

        const addFileEndpoint = `${this._siteServerRelativeUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(
          fullServerRelFolderPath
        )}')/Files/Add(url='${encodeURIComponent(fileName)}', overwrite=true)`;

        const response: SPHttpClientResponse = await this._spHttpClient.post(addFileEndpoint, SPHttpClient.configurations.v1, {
          headers: {
            Accept: 'application/json;odata=nometadata',
            'Content-Type': 'application/json;charset=utf-8'
          },
          body: jsonContent
        });

        if (!response.ok) {
          const errText = await response.text();
          throw new Error(`Failed to upload dashboard snapshot: ${response.status} ${errText}`);
        }

        return {
          success: true,
          fileName,
          folderPath: fullServerRelFolderPath,
          serverRelativeUrl: fullFileServerRelUrl,
          message: `Saved snapshot to ${this._libraryTitle}/${options.folderType}/${safeDashboardFolder}/${fileName}`,
          savedAt: new Date().toISOString()
        };
      } catch (err) {
        console.warn('[DashboardStorageService] SharePoint save failed, saving to local fallback:', err);
        return this._saveToLocalMock(options.folderType, safeDashboardFolder, fileName, fullFileServerRelUrl, pkg);
      }
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
        const queryUrl = `${this._siteServerRelativeUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(
          fullServerRelFolderPath
        )}')/Files?$select=Name,ServerRelativeUrl,TimeLastModified,Length&$orderby=TimeLastModified desc`;

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
        const getFileEndpoint = `${this._siteServerRelativeUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(
          serverRelativeFileUrl
        )}')/$value`;

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
   * Iteratively creates folders in the SharePoint Document Library if they don't already exist.
   */
  private async _ensureSharePointFolderPath(relativeFolderPath: string): Promise<void> {
    if (!this._spHttpClient || !this._siteServerRelativeUrl) return;

    const segments = relativeFolderPath.split('/').filter(Boolean);
    let currentPath = '';

    for (const segment of segments) {
      const parentPath = currentPath;
      currentPath = currentPath ? `${currentPath}/${segment}` : segment;
      const fullParentServerRel = this._siteServerRelativeUrl
        ? `${this._siteServerRelativeUrl}/${parentPath}`
        : `/${parentPath}`;

      try {
        const endpoint = parentPath
          ? `${this._siteServerRelativeUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(
              fullParentServerRel
            )}')/folders`
          : `${this._siteServerRelativeUrl}/_api/web/folders`;

        await this._spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
          headers: {
            Accept: 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=verbose'
          },
          body: JSON.stringify({ ServerRelativeUrl: segment })
        });
      } catch {
        // Ignore if folder already exists
      }
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
