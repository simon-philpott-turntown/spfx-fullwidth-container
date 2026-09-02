/**
 * @file TermStorePicker.tsx
 * @description SharePoint Global Term Store tag picker component with Fluent UI 2 tags,
 * real-time search, category filters, and hierarchical Term Store tree browser.
 */

import * as React from 'react';
import {
  makeStyles,
  tokens,
  shorthands,
  Tag,
  TagGroup,
  Input,
  Button,
  Caption1,
  Skeleton,
  SkeletonItem
} from '@fluentui/react-components';
import {
  TagRegular,
  AddRegular,
  SearchRegular,
  FolderRegular,
  FolderOpenRegular,
  CheckmarkRegular,
  DeleteRegular
} from '@fluentui/react-icons';
import { ITermStoreTag } from '../models/IContainerModels';
import { TaxonomyService, ITermGroup } from '../services/TaxonomyService';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
    width: '100%'
  },
  tagContainer: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '6px',
    alignItems: 'center'
  },
  pickerBox: {
    ...shorthands.padding('12px'),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorNeutralBackground2,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke1),
    display: 'flex',
    flexDirection: 'column',
    gap: '10px',
    marginTop: '6px'
  },
  termSetFilterRow: {
    display: 'flex',
    gap: '6px',
    flexWrap: 'wrap',
    marginBottom: '2px'
  },
  filterPill: {
    fontSize: '0.75rem',
    cursor: 'pointer'
  },
  suggestionsList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
    maxHeight: '180px',
    overflowY: 'auto',
    ...shorthands.padding('4px'),
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.borderRadius(tokens.borderRadiusSmall),
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2)
  },
  suggestionItem: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    ...shorthands.padding('6px', '8px'),
    ...shorthands.borderRadius(tokens.borderRadiusSmall),
    cursor: 'pointer',
    transitionProperty: 'background-color, color',
    transitionDuration: '120ms',
    ':hover': {
      backgroundColor: tokens.colorBrandBackground2,
      color: tokens.colorBrandForeground1
    }
  },
  treeContainer: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
    maxHeight: '200px',
    overflowY: 'auto',
    ...shorthands.padding('6px'),
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.borderRadius(tokens.borderRadiusSmall),
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2)
  },
  treeNode: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    ...shorthands.padding('4px', '6px'),
    ...shorthands.borderRadius(tokens.borderRadiusSmall),
    cursor: 'pointer',
    fontSize: '0.85rem',
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground2
    }
  },
  treeLeaf: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginLeft: '20px',
    ...shorthands.padding('3px', '6px'),
    ...shorthands.borderRadius(tokens.borderRadiusSmall),
    cursor: 'pointer',
    fontSize: '0.8rem',
    ':hover': {
      backgroundColor: tokens.colorBrandBackground2,
      color: tokens.colorBrandForeground1
    }
  },
  addNewRow: {
    display: 'flex',
    gap: '8px',
    alignItems: 'center',
    marginTop: '4px'
  }
});

export interface ITermStorePickerProps {
  selectedTags?: ITermStoreTag[];
  onChange: (tags: ITermStoreTag[]) => void;
  isEditMode?: boolean;
}

export const TermStorePicker: React.FC<ITermStorePickerProps> = ({
  selectedTags = [],
  onChange,
  isEditMode = false
}) => {
  const styles = useStyles();
  const [isPickerOpen, setIsPickerOpen] = React.useState<boolean>(false);
  const [searchQuery, setSearchQuery] = React.useState<string>('');
  const [selectedTermSet, setSelectedTermSet] = React.useState<string>('all');
  const [viewMode, setViewMode] = React.useState<'search' | 'tree'>('search');
  const [allTerms, setAllTerms] = React.useState<ITermStoreTag[]>([]);
  const [termGroups, setTermGroups] = React.useState<ITermGroup[]>([]);
  const [expandedSets, setExpandedSets] = React.useState<Record<string, boolean>>({
    'set-our-sectors': true,
    'set-our-segments': false
  });
  const [newTermInput, setNewTermInput] = React.useState<string>('');
  const [isLoading, setIsLoading] = React.useState<boolean>(true);

  React.useEffect(() => {
    setIsLoading(true);
    Promise.all([
      TaxonomyService.getTerms(),
      TaxonomyService.getTermGroups()
    ]).then(([terms, groups]) => {
      setAllTerms(terms);
      setTermGroups(groups);
      setIsLoading(false);
    }).catch(() => {
      setIsLoading(false);
    });
  }, []);

  const handleSelectTerm = (term: ITermStoreTag): void => {
    if (!selectedTags.some((t) => t.id === term.id || t.label.toLowerCase() === term.label.toLowerCase())) {
      onChange([...selectedTags, term]);
    }
  };

  const handleRemoveTag = (tagId: string): void => {
    onChange(selectedTags.filter((t) => t.id !== tagId));
  };

  const handleCreateTerm = (): void => {
    if (!newTermInput.trim()) return;
    const targetSet = selectedTermSet !== 'all' ? selectedTermSet : 'Our Sectors';
    const created = TaxonomyService.addTerm(newTermInput.trim(), targetSet);
    setAllTerms([...allTerms, created]);
    void TaxonomyService.getTermGroups().then((groups) => setTermGroups([...groups]));
    handleSelectTerm(created);
    setNewTermInput('');
  };

  const toggleSetExpansion = (setId: string): void => {
    setExpandedSets((prev) => ({
      ...prev,
      [setId]: !prev[setId]
    }));
  };

  const filteredSuggestions = allTerms.filter((t) => {
    const isAlreadySelected = selectedTags.some((st) => st.label.toLowerCase() === t.label.toLowerCase());
    if (isAlreadySelected) return false;

    const matchesSet =
      selectedTermSet === 'all' ||
      (t.termSetName && t.termSetName.toLowerCase().includes(selectedTermSet.toLowerCase()));
    if (!matchesSet) return false;

    if (!searchQuery.trim()) return true;
    const q = searchQuery.toLowerCase().trim();
    return (
      t.label.toLowerCase().includes(q) ||
      (t.termSetName && t.termSetName.toLowerCase().includes(q)) ||
      (t.path && t.path.toLowerCase().includes(q))
    );
  });

  return (
    <div className={styles.root}>
      {/* Selected Tags Display */}
      <TagGroup onDismiss={(e, data) => handleRemoveTag(data.value)}>
        <div className={styles.tagContainer}>
          {selectedTags.map((tag) => (
            <Tag
              key={tag.id}
              value={tag.id}
              shape="rounded"
              appearance="outline"
              size="extra-small"
              icon={<TagRegular style={{ fontSize: '10px', width: '10px', height: '10px' }} />}
              dismissible={isEditMode}
              dismissIcon={isEditMode ? <DeleteRegular style={{ fontSize: '10px', width: '10px', height: '10px', cursor: 'pointer' }} /> : undefined}
              style={{
                fontSize: '0.68rem',
                lineHeight: '0.9rem',
                padding: '1px 5px',
                height: '19px'
              }}
            >
              {tag.label}
            </Tag>
          ))}

          {isEditMode && (
            <Button
              size="small"
              appearance="subtle"
              style={{ fontSize: '0.72rem', height: '20px', padding: '0 6px' }}
              icon={isPickerOpen ? <CheckmarkRegular style={{ fontSize: '11px' }} /> : <AddRegular style={{ fontSize: '11px' }} />}
              onClick={() => setIsPickerOpen(!isPickerOpen)}
            >
              {isPickerOpen ? 'Done' : '+ Add Global Term Store Tag'}
            </Button>
          )}
        </div>
      </TagGroup>

      {/* Interactive Global Term Store Taxonomy Picker */}
      {isEditMode && isPickerOpen && (
        <div className={styles.pickerBox}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <div>
              <Caption1 style={{ color: tokens.colorNeutralForeground1, fontWeight: 600 }}>
                Global Term Store • Intranet Taxonomy
              </Caption1>
              <Caption1 style={{ color: tokens.colorNeutralForeground4, display: 'block', fontSize: '0.72rem' }}>
                Connected to SharePoint Online Term Store Admin Center
              </Caption1>
            </div>
            <div style={{ display: 'flex', gap: '4px' }}>
              <Button
                size="small"
                appearance={viewMode === 'search' ? 'primary' : 'subtle'}
                onClick={() => setViewMode('search')}
              >
                Search
              </Button>
              <Button
                size="small"
                appearance={viewMode === 'tree' ? 'primary' : 'subtle'}
                onClick={() => setViewMode('tree')}
              >
                Browse Tree
              </Button>
            </div>
          </div>

          {viewMode === 'search' ? (
            <>
              {/* Term Set Quick Filter Pills */}
              <div className={styles.termSetFilterRow}>
                <Button
                  size="small"
                  appearance={selectedTermSet === 'all' ? 'primary' : 'secondary'}
                  className={styles.filterPill}
                  onClick={() => setSelectedTermSet('all')}
                >
                  All Terms
                </Button>
                <Button
                  size="small"
                  appearance={selectedTermSet === 'our sectors' ? 'primary' : 'secondary'}
                  className={styles.filterPill}
                  onClick={() => setSelectedTermSet('our sectors')}
                >
                  🏢 Our Sectors
                </Button>
                <Button
                  size="small"
                  appearance={selectedTermSet === 'our segments' ? 'primary' : 'secondary'}
                  className={styles.filterPill}
                  onClick={() => setSelectedTermSet('our segments')}
                >
                  📊 Our Segments
                </Button>
                <Button
                  size="small"
                  appearance={selectedTermSet === 'amcl' ? 'primary' : 'secondary'}
                  className={styles.filterPill}
                  onClick={() => setSelectedTermSet('amcl')}
                >
                  📁 AMCL Terms
                </Button>
              </div>

              {/* Live Search Input */}
              <Input
                size="small"
                contentBefore={<SearchRegular />}
                placeholder="Search Global Term Store (e.g. Infrastructure, Clean energy, Data centres)..."
                value={searchQuery}
                onChange={(_, data) => setSearchQuery(data.value)}
              />

              {/* Results List */}
              <div className={styles.suggestionsList}>
                {isLoading ? (
                  <Skeleton aria-label="Loading taxonomy terms" style={{ display: 'flex', flexDirection: 'column', gap: '8px', padding: '8px' }}>
                    <SkeletonItem style={{ height: '22px', width: '80%', borderRadius: '4px' }} />
                    <SkeletonItem style={{ height: '22px', width: '65%', borderRadius: '4px' }} />
                    <SkeletonItem style={{ height: '22px', width: '90%', borderRadius: '4px' }} />
                  </Skeleton>
                ) : (
                  <>
                    {filteredSuggestions.map((term) => (
                      <div
                        key={term.id}
                        className={styles.suggestionItem}
                        onClick={() => handleSelectTerm(term)}
                      >
                        <div>
                          <span style={{ fontWeight: 600, fontSize: '0.85rem' }}>+ {term.label}</span>
                          {term.path && (
                            <Caption1 style={{ display: 'block', color: tokens.colorNeutralForeground4, fontSize: '0.7rem' }}>
                              {term.path}
                            </Caption1>
                          )}
                        </div>
                        {term.termSetName && (
                          <Tag size="small" shape="rounded" appearance="brand">
                            {term.termSetName}
                          </Tag>
                        )}
                      </div>
                    ))}

                    {filteredSuggestions.length === 0 && (
                      <div style={{ padding: '8px', textAlign: 'center' }}>
                        <Caption1 style={{ color: tokens.colorNeutralForeground4 }}>
                          No matching terms in &ldquo;{selectedTermSet}&rdquo;. Create a custom tag below.
                        </Caption1>
                      </div>
                    )}
                  </>
                )}
              </div>
            </>
          ) : (
            /* Hierarchical Tree Browser (Matching Screenshot 2) */
            <div className={styles.treeContainer}>
              {termGroups.map((grp) => (
                <div key={grp.id} style={{ marginBottom: '6px' }}>
                  <div className={styles.treeNode} style={{ fontWeight: 600, color: tokens.colorNeutralForeground1 }}>
                    <FolderOpenRegular style={{ color: tokens.colorBrandForeground1 }} />
                    <span>Global Group: {grp.name}</span>
                  </div>
                  {grp.termSets.map((set) => {
                    const isExpanded = !!expandedSets[set.id];
                    return (
                      <div key={set.id} style={{ marginLeft: '14px', marginTop: '2px' }}>
                        <div
                          className={styles.treeNode}
                          style={{ fontWeight: 500, color: tokens.colorNeutralForeground2 }}
                          onClick={() => toggleSetExpansion(set.id)}
                        >
                          {isExpanded ? <FolderOpenRegular /> : <FolderRegular />}
                          <span>Term Set: {set.name} ({set.terms.length})</span>
                        </div>
                        {isExpanded && (
                          <div style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                            {set.terms.map((term) => {
                              const isSelected = selectedTags.some(
                                (st) => st.label.toLowerCase() === term.label.toLowerCase()
                              );
                              return (
                                <div
                                  key={term.id}
                                  className={styles.treeLeaf}
                                  style={{
                                    opacity: isSelected ? 0.5 : 1,
                                    cursor: isSelected ? 'default' : 'pointer'
                                  }}
                                  onClick={() => !isSelected && handleSelectTerm(term)}
                                >
                                  <span>{isSelected ? '✓ ' : '+ '}{term.label}</span>
                                  {isSelected && (
                                    <Tag size="small" appearance="outline">Selected</Tag>
                                  )}
                                </div>
                              );
                            })}
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              ))}
            </div>
          )}

          {/* Add New Custom Term Row */}
          <div className={styles.addNewRow}>
            <Input
              size="small"
              placeholder={`New term for ${selectedTermSet !== 'all' ? selectedTermSet : 'Our Sectors'}...`}
              value={newTermInput}
              onChange={(_, data) => setNewTermInput(data.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  e.preventDefault();
                  handleCreateTerm();
                }
              }}
            />
            <Button size="small" appearance="primary" onClick={handleCreateTerm}>
              Add Term
            </Button>
          </div>
        </div>
      )}
    </div>
  );
};
