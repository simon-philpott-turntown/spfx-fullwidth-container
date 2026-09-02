/**
 * @file TaxonomyService.ts
 * @description Service for querying SharePoint Global Term Store taxonomy directly in SharePoint Online (SPFx)
 * with hierarchical term groups, term sets, live REST/v2.1 endpoints, and comprehensive tenant fallback cache.
 */

import { ITermStoreTag } from '../models/IContainerModels';

export interface ITermGroup {
  id: string;
  name: string;
  termSets: ITermSet[];
}

export interface ITermSet {
  id: string;
  name: string;
  groupId: string;
  groupName: string;
  terms: ITermStoreTag[];
}

export class TaxonomyService {
  private static _cachedGroups: ITermGroup[] = [
    {
      id: 'grp-intranet',
      name: 'Intranet',
      termSets: [
        {
          id: 'set-our-sectors',
          name: 'Our Sectors',
          groupId: 'grp-intranet',
          groupName: 'Intranet',
          terms: [
            { id: 'sec-clean-energy', label: 'Clean energy', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Clean energy' },
            { id: 'sec-conv-power', label: 'Conventional and low carbon power', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Conventional and low carbon power' },
            { id: 'sec-data-centres', label: 'Data centres', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Data centres' },
            { id: 'sec-defense', label: 'Defense', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Defense' },
            { id: 'sec-education', label: 'Education', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Education' },
            { id: 'sec-elec-trans', label: 'Electrical transmission and distribution', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Electrical transmission and distribution' },
            { id: 'sec-energy-res', label: 'Energy and resources', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Energy and resources' },
            { id: 'sec-environment', label: 'Environment', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Environment' },
            { id: 'sec-fin-prof', label: 'Finance and professional services', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Finance and professional services' },
            { id: 'sec-health', label: 'Health', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Health' },
            { id: 'sec-infra', label: 'Infrastructure', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Infrastructure' },
            { id: 'sec-prop', label: 'Real estate and property', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Real estate and property' },
            { id: 'sec-av', label: 'Aviation and transport', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Aviation and transport' },
            { id: 'sec-gov', label: 'Government and public sector', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Government and public sector' },
            { id: 'sec-tech', label: 'Technology and digital', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Technology and digital' },
            { id: 'sec-mining', label: 'Mining and metals', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Mining and metals' },
            { id: 'sec-life-sci', label: 'Life sciences and pharmaceuticals', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Life sciences and pharmaceuticals' }
          ]
        },
        {
          id: 'set-our-segments',
          name: 'Our Segments',
          groupId: 'grp-intranet',
          groupName: 'Intranet',
          terms: [
            { id: 'seg-comm', label: 'Commercial advisory', termSetName: 'Our Segments', path: 'Intranet > Our Segments > Commercial advisory' },
            { id: 'seg-prog', label: 'Programme management', termSetName: 'Our Segments', path: 'Intranet > Our Segments > Programme management' },
            { id: 'seg-proj', label: 'Project controls', termSetName: 'Our Segments', path: 'Intranet > Our Segments > Project controls' },
            { id: 'seg-cost', label: 'Cost and commercial management', termSetName: 'Our Segments', path: 'Intranet > Our Segments > Cost and commercial management' },
            { id: 'seg-strat', label: 'Strategic advisory', termSetName: 'Our Segments', path: 'Intranet > Our Segments > Strategic advisory' },
            { id: 'seg-esg', label: 'Sustainability, ESG and net zero', termSetName: 'Our Segments', path: 'Intranet > Our Segments > Sustainability, ESG and net zero' },
            { id: 'seg-proc', label: 'Procurement and supply chain', termSetName: 'Our Segments', path: 'Intranet > Our Segments > Procurement and supply chain' },
            { id: 'seg-disp', label: 'Dispute resolution', termSetName: 'Our Segments', path: 'Intranet > Our Segments > Dispute resolution' },
            { id: 'seg-asset-ops', label: 'Asset management and operations', termSetName: 'Our Segments', path: 'Intranet > Our Segments > Asset management and operations' }
          ]
        }
      ]
    },
    {
      id: 'grp-amcl',
      name: 'AMCL Terms',
      termSets: [
        {
          id: 'set-amcl-core',
          name: 'Core Capabilities',
          groupId: 'grp-amcl',
          groupName: 'AMCL Terms',
          terms: [
            { id: 'amcl-frameworks', label: 'Asset management frameworks', termSetName: 'Core Capabilities', path: 'AMCL Terms > Core Capabilities > Asset management frameworks' },
            { id: 'amcl-iso55000', label: 'ISO 55000 compliance', termSetName: 'Core Capabilities', path: 'AMCL Terms > Core Capabilities > ISO 55000 compliance' },
            { id: 'amcl-readiness', label: 'Operational readiness', termSetName: 'Core Capabilities', path: 'AMCL Terms > Core Capabilities > Operational readiness' }
          ]
        }
      ]
    }
  ];

  private static _isFetched: boolean = false;

  /**
   * Queries the live SharePoint Online Term Store v2.1 REST API if available on the current site context.
   */
  public static async initializeFromSharePoint(siteUrl?: string): Promise<void> {
    if (this._isFetched) return;

    if (typeof window !== 'undefined') {
      try {
        const spCtx = (window as unknown as { _spPageContextInfo?: { webAbsoluteUrl?: string } })._spPageContextInfo;
        const targetUrl = siteUrl || (spCtx && spCtx.webAbsoluteUrl) || '';

        if (targetUrl) {
          // Attempt 1: v2.1 TermStore Groups with Sets and Terms expanded
          const endpoint = `${targetUrl}/_api/v2.1/termStore/groups?$expand=sets($expand=terms)`;
          const response = await fetch(endpoint, {
            headers: {
              Accept: 'application/json;odata=nometadata',
              'Content-Type': 'application/json'
            },
            credentials: 'include'
          });

          if (response.ok) {
            const data = await response.json();
            if (data && Array.isArray(data.value) && data.value.length > 0) {
              const liveGroups: ITermGroup[] = [];
              data.value.forEach((g: { id: string; displayName?: string; name?: string; sets?: Array<{ id: string; localizedNames?: Array<{ name: string }>; terms?: Array<{ id: string; labels?: Array<{ name: string }> }> }> }) => {
                const groupName = g.displayName || g.name || 'Global Group';
                const termSets: ITermSet[] = [];

                if (Array.isArray(g.sets)) {
                  g.sets.forEach((s) => {
                    const setName = s.localizedNames && s.localizedNames[0] ? s.localizedNames[0].name : 'Term Set';
                    const terms: ITermStoreTag[] = [];

                    if (Array.isArray(s.terms)) {
                      s.terms.forEach((t) => {
                        const termLabel = t.labels && t.labels[0] ? t.labels[0].name : 'Term';
                        terms.push({
                          id: t.id,
                          label: termLabel,
                          termSetId: s.id,
                          termSetName: setName,
                          path: `${groupName} > ${setName} > ${termLabel}`
                        });
                      });
                    }

                    termSets.push({
                      id: s.id,
                      name: setName,
                      groupId: g.id,
                      groupName: groupName,
                      terms: terms
                    });
                  });
                }

                liveGroups.push({
                  id: g.id,
                  name: groupName,
                  termSets: termSets
                });
              });

              if (liveGroups.length > 0) {
                this._cachedGroups = liveGroups;
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
  }

  /**
   * Returns all hierarchical term groups and their term sets.
   */
  public static async getTermGroups(siteUrl?: string): Promise<ITermGroup[]> {
    await this.initializeFromSharePoint(siteUrl);
    return this._cachedGroups;
  }

  /**
   * Fetches flattened list of terms matching optional search query and term set filter.
   */
  public static async getTerms(filterQuery?: string, termSetFilter?: string, siteUrl?: string): Promise<ITermStoreTag[]> {
    await this.initializeFromSharePoint(siteUrl);

    const allTerms: ITermStoreTag[] = [];
    this._cachedGroups.forEach((g) => {
      g.termSets.forEach((s) => {
        s.terms.forEach((t) => {
          allTerms.push(t);
        });
      });
    });

    let results = allTerms;

    if (termSetFilter && termSetFilter !== 'all') {
      const targetFilter = termSetFilter.toLowerCase().trim();
      results = results.filter(
        (t) => t.termSetName?.toLowerCase() === targetFilter || (t.path && t.path.toLowerCase().indexOf(targetFilter) !== -1)
      );
    }

    if (!filterQuery || !filterQuery.trim()) {
      return results;
    }

    const q = filterQuery.toLowerCase().trim();
    return results.filter((t) => {
      const labelMatch = t.label.toLowerCase().indexOf(q) !== -1;
      const setMatch = t.termSetName ? t.termSetName.toLowerCase().indexOf(q) !== -1 : false;
      const pathMatch = t.path ? t.path.toLowerCase().indexOf(q) !== -1 : false;
      return labelMatch || setMatch || pathMatch;
    });
  }

  /**
   * Returns list of unique term sets available across all term groups.
   */
  public static getAvailableTermSets(): string[] {
    const sets = new Set<string>();
    this._cachedGroups.forEach((g) => {
      g.termSets.forEach((s) => {
        sets.add(s.name);
      });
    });
    return Array.from(sets);
  }

  /**
   * Adds a new custom term to the active term group/set.
   */
  public static addTerm(label: string, termSetName: string = 'Our Sectors', groupName: string = 'Intranet'): ITermStoreTag {
    const trimmed = label.trim();
    const newTerm: ITermStoreTag = {
      id: `t-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`,
      label: trimmed,
      termSetName: termSetName,
      path: `${groupName} > ${termSetName} > ${trimmed}`
    };

    let group = this._cachedGroups.find((g) => g.name.toLowerCase() === groupName.toLowerCase());
    if (!group) {
      group = { id: `grp-${Date.now()}`, name: groupName, termSets: [] };
      this._cachedGroups.push(group);
    }

    let set = group.termSets.find((s) => s.name.toLowerCase() === termSetName.toLowerCase());
    if (!set) {
      set = { id: `set-${Date.now()}`, name: termSetName, groupId: group.id, groupName: group.name, terms: [] };
      group.termSets.push(set);
    }

    set.terms.push(newTerm);
    return newTerm;
  }
}
