/**
 * @file TaxonomyService.ts
 * @description Service for querying SharePoint Global Term Store terms with resilient fallback.
 */

import { ITermStoreTag } from '../models/IContainerModels';

export class TaxonomyService {
  private static _cachedTerms: ITermStoreTag[] = [
    { id: 't-comm', label: 'Commercial advisory', termSetName: 'Commercial governance' },
    { id: 't-adv', label: 'Advisory services', termSetName: 'Commercial governance' },
    { id: 't-p0', label: 'Priority 1 mandatory', termSetName: 'Priority class' },
    { id: 't-p1', label: 'Standard priority', termSetName: 'Priority class' },
    { id: 't-comp', label: 'Compliance assurance', termSetName: 'Quality assurance' },
    { id: 't-qual', label: 'Quality standard', termSetName: 'Quality assurance' },
    { id: 't-bim', label: 'Digital engineering', termSetName: 'Technical standards' },
    { id: 't-iso', label: 'ISO 19650 standard', termSetName: 'Technical standards' },
    { id: 't-safe', label: 'Health and safety', termSetName: 'Safety governance' },
    { id: 't-hse', label: 'Safety compliance', termSetName: 'Safety governance' },
    { id: 't-fin', label: 'Financial controls (£)', termSetName: 'Financial governance' },
    { id: 't-ops', label: 'Operational delivery', termSetName: 'Service delivery' },
    { id: 't-strat', label: 'Executive strategy', termSetName: 'Executive board' },
    { id: 't-risk', label: 'Risk management', termSetName: 'Risk register' }
  ];

  private static _isFetched: boolean = false;

  /**
   * Fetches terms from SharePoint Global Term Store on the tenant or fallback terms.
   */
  public static async getTerms(filterQuery?: string): Promise<ITermStoreTag[]> {
    if (!this._isFetched && typeof window !== 'undefined') {
      try {
        const spCtx = (window as unknown as { _spPageContextInfo?: { webAbsoluteUrl?: string } })._spPageContextInfo;
        if (spCtx && spCtx.webAbsoluteUrl) {
          const endpoint = `${spCtx.webAbsoluteUrl}/_api/v2.1/termStore/sets?$expand=terms`;
          const response = await fetch(endpoint, {
            headers: {
              Accept: 'application/json;odata=nometadata',
              'Content-Type': 'application/json'
            },
            credentials: 'include'
          });

          if (response.ok) {
            const data = await response.json();
            if (data && Array.isArray(data.value)) {
              const liveTerms: ITermStoreTag[] = [];
              data.value.forEach((set: { id: string; localizedNames?: Array<{ name: string }>; terms?: Array<{ id: string; labels?: Array<{ name: string }> }> }) => {
                const setName = set.localizedNames && set.localizedNames[0] ? set.localizedNames[0].name : 'Global terms';
                if (Array.isArray(set.terms)) {
                  set.terms.forEach((term) => {
                    const termLabel = term.labels && term.labels[0] ? term.labels[0].name : 'Term';
                    liveTerms.push({
                      id: term.id,
                      label: termLabel,
                      termSetName: setName
                    });
                  });
                }
              });

              if (liveTerms.length > 0) {
                this._cachedTerms = liveTerms;
              }
            }
          }
        }
      } catch {
        // Fallback gracefully to tenant defaults
      } finally {
        this._isFetched = true;
      }
    }

    if (!filterQuery) {
      return this._cachedTerms;
    }

    const q = filterQuery.toLowerCase();
    return this._cachedTerms.filter(
      (t) => t.label.toLowerCase().indexOf(q) !== -1 || (t.termSetName && t.termSetName.toLowerCase().indexOf(q) !== -1)
    );
  }

  /**
   * Adds a custom term to local term cache for authoring flexibility.
   */
  public static addTerm(label: string, termSetName: string = 'Custom terms'): ITermStoreTag {
    const newTerm: ITermStoreTag = {
      id: `t-${Date.now()}`,
      label: label.trim(),
      termSetName: termSetName
    };
    this._cachedTerms.push(newTerm);
    return newTerm;
  }
}
