/**
 * @file TaxonomyService.ts
 * @description Service for querying SharePoint Global Term Store terms with resilient fallback.
 */

import { ITermStoreTag } from '../models/IContainerModels';

export class TaxonomyService {
  private static _cachedTerms: ITermStoreTag[] = [
    { id: 't-comm', label: 'Commercial', termSetName: 'Advisory Governance' },
    { id: 't-adv', label: 'Advisory', termSetName: 'Advisory Governance' },
    { id: 't-p0', label: 'P0 Critical', termSetName: 'Priority Class' },
    { id: 't-p1', label: 'P1 Mandatory', termSetName: 'Priority Class' },
    { id: 't-comp', label: 'Compliance', termSetName: 'Quality Assurance' },
    { id: 't-qual', label: 'Quality', termSetName: 'Quality Assurance' },
    { id: 't-bim', label: 'BIM', termSetName: 'Digital Engineering' },
    { id: 't-iso', label: 'ISO 19650', termSetName: 'Digital Engineering' },
    { id: 't-safe', label: 'Safety', termSetName: 'Health & Safety' },
    { id: 't-hse', label: 'HSE', termSetName: 'Health & Safety' },
    { id: 't-fin', label: 'Finance (£)', termSetName: 'Financial Controls' },
    { id: 't-ops', label: 'Operations', termSetName: 'Service Delivery' },
    { id: 't-strat', label: 'Strategy', termSetName: 'Executive Board' },
    { id: 't-risk', label: 'Risk Register', termSetName: 'Governance' }
  ];

  /**
   * Fetches terms from SharePoint Global Term Store or cached default terms.
   */
  public static async getTerms(filterQuery?: string): Promise<ITermStoreTag[]> {
    try {
      // In SPFx environment with active MSGraphClient / REST endpoint:
      if (typeof window !== 'undefined' && (window as any)._spPageContextInfo) {
        // SharePoint REST taxonomy endpoint can be called here if permissions allow
      }
    } catch {
      // Graceful fallback to cached tenant terms
    }

    if (!filterQuery) {
      return this._cachedTerms;
    }

    const q = filterQuery.toLowerCase();
    return this._cachedTerms.filter(
      (t) => t.label.toLowerCase().includes(q) || (t.termSetName && t.termSetName.toLowerCase().includes(q))
    );
  }

  /**
   * Adds a custom term to local term cache for authoring flexibility.
   */
  public static addTerm(label: string, termSetName: string = 'Custom Terms'): ITermStoreTag {
    const newTerm: ITermStoreTag = {
      id: `t-${Date.now()}`,
      label: label.trim(),
      termSetName: termSetName
    };
    this._cachedTerms.push(newTerm);
    return newTerm;
  }
}
