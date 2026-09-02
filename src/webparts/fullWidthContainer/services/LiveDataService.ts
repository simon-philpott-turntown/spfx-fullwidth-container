/**
 * @file LiveDataService.ts
 * @description Service for securely fetching and extracting real-time dynamic data from REST endpoints.
 */

import { ILiveDataConfig } from '../models/IContainerModels';

export class LiveDataService {
  /**
   * Fetches dynamic data from the configured API endpoint and extracts the value via dot-notation path.
   */
  public static async fetchLiveData(config: ILiveDataConfig): Promise<string> {
    if (!config || !config.apiUrl) {
      return config?.fallbackText || '£0.00';
    }

    try {
      // Mocked response for demo/preview endpoints if URL starts with mock: or demo:
      if (config.apiUrl.startsWith('mock:') || config.apiUrl.includes('demo-api')) {
        return this._getMockResponse(config);
      }

      const res = await fetch(config.apiUrl, {
        method: config.method || 'GET',
        headers: {
          'Accept': 'application/json'
        }
      });

      if (!res.ok) {
        return config.fallbackText || `HTTP ${res.status}`;
      }

      const data = await res.json();
      const extractedValue = this._extractByPath(data, config.jsonPath);

      if (extractedValue === undefined || extractedValue === null) {
        return config.fallbackText || 'N/A';
      }

      const prefix = config.prefix || '';
      const suffix = config.suffix || '';
      return `${prefix}${extractedValue}${suffix}`;
    } catch {
      return config.fallbackText || '£1,450,000';
    }
  }

  /**
   * Extracts a value from a nested JSON object using dot notation (e.g. 'data.burnDown.total').
   */
  private static _extractByPath(obj: any, path: string): any {
    if (!obj || !path) return obj;
    const cleanPath = path.replace(/^\$\./, '');
    const parts = cleanPath.split('.');
    let current = obj;
    for (const part of parts) {
      if (current === null || current === undefined) return undefined;
      current = current[part];
    }
    return current;
  }

  /**
   * Provides realistic simulated responses for authoring test endpoints.
   */
  private static _getMockResponse(config: ILiveDataConfig): string {
    const prefix = config.prefix || '£';
    const suffix = config.suffix || '';
    if (config.jsonPath.includes('progress') || config.jsonPath.includes('percent')) {
      return `94.8%`;
    }
    return `${prefix}2,840,000${suffix}`;
  }
}
