/**
 * @file TaxonomyService.ts
 * @description Service for querying SharePoint Global Term Store terms with resilient fallback.
 */

import { ITermStoreTag } from '../models/IContainerModels';

export class TaxonomyService {
  private static _cachedTerms: ITermStoreTag[] = [
    // Intranet -> Our Sectors
    { id: 'sec-infra', label: 'Infrastructure', termSetName: 'Our Sectors' },
    { id: 'sec-energy', label: 'Energy and utilities', termSetName: 'Our Sectors' },
    { id: 'sec-prop', label: 'Real estate and property', termSetName: 'Our Sectors' },
    { id: 'sec-def', label: 'Defence and security', termSetName: 'Our Sectors' },
    { id: 'sec-av', label: 'Aviation and transport', termSetName: 'Our Sectors' },
    { id: 'sec-health', label: 'Healthcare and science', termSetName: 'Our Sectors' },
    { id: 'sec-gov', label: 'Government and public sector', termSetName: 'Our Sectors' },
    { id: 'sec-tech', label: 'Technology and digital', termSetName: 'Our Sectors' },

    // Intranet -> Our Segments
    { id: 'seg-comm', label: 'Commercial advisory', termSetName: 'Our Segments' },
    { id: 'seg-prog', label: 'Programme management', termSetName: 'Our Segments' },
    { id: 'seg-proj', label: 'Project controls', termSetName: 'Our Segments' },
    { id: 'seg-cost', label: 'Cost and commercial management', termSetName: 'Our Segments' },
    { id: 'seg-strat', label: 'Strategic advisory', termSetName: 'Our Segments' },
    { id: 'seg-esg', label: 'Sustainability, ESG and net zero', termSetName: 'Our Segments' },
    { id: 'seg-proc', label: 'Procurement and supply chain', termSetName: 'Our Segments' },
    { id: 'seg-disp', label: 'Dispute resolution', termSetName: 'Our Segments' }
  ];

  private static _isFetched: boolean = false;

  /**
   * Fetches terms from SharePoint Global Term Store on the tenant or fallback terms.
   */
  public static async getTerms(filterQuery?: string, termSetFilter?: string): Promise<ITermStoreTag[]> {
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
                const setName = set.localizedNames && set.localizedNames[0] ? set.localizedNames[0].name : 'Intranet';
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

    let results = this._cachedTerms;

    if (termSetFilter && termSetFilter !== 'all') {
      results = results.filter((t) => t.termSetName?.toLowerCase() === termSetFilter.toLowerCase());
    }

    if (!filterQuery) {
      return results;
    }

    const q = filterQuery.toLowerCase();
    return results.filter(
      (t) => t.label.toLowerCase().indexOf(q) !== -1 || (t.termSetName && t.termSetName.toLowerCase().indexOf(q) !== -1)
    );
  }

  /**
   * Returns list of unique term sets available in the Intranet taxonomy group.
   */
  public static getAvailableTermSets(): string[] {
    const sets = new Set<string>();
    this._cachedTerms.forEach((t) => {
      if (t.termSetName) sets.add(t.termSetName);
    });
    return Array.from(sets);
  }

  /**
   * Adds a custom term to local term cache for authoring flexibility.
   */
  public static addTerm(label: string, termSetName: string = 'Our Segments'): ITermStoreTag {
    const newTerm: ITermStoreTag = {
      id: `t-${Date.now()}`,
      label: label.trim(),
      termSetName: termSetName
    };
    this._cachedTerms.push(newTerm);
    return newTerm;
  }
}
