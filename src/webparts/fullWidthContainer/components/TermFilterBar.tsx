/**
 * @file TermFilterBar.tsx
 * @description Dynamic Global Term Store Filter Bar component built with Fluent UI 2 (@fluentui/react-components v9).
 * Renders multiple taxonomy dropdown filters, dynamically fetches terms from SharePoint Term Store via TaxonomyService,
 * displays an interactive summary indicator (e.g. "✓ updated — showing UK requirements for Infrastructure"),
 * and provides authoring controls in edit mode to add, configure, or remove filter dropdowns.
 */

import * as React from 'react';
import {
  makeStyles,
  tokens,
  shorthands,
  Dropdown,
  Option,
  Button,
  Caption1,
  Spinner,
  SelectionEvents,
  OptionOnSelectData
} from '@fluentui/react-components';
import {
  CheckmarkCircleRegular,
  DismissRegular,
  AddRegular,
  SettingsRegular,
  FilterRegular,
  LocationRegular,
  BuildingRegular,
  GlobeRegular,
  TagRegular,
  BriefcaseRegular,
  PeopleRegular,
  BookOpenRegular,
  SparkleRegular,
  TargetRegular,
  ArrowLeftRegular,
  ArrowRightRegular
} from '@fluentui/react-icons';
import { ITermFilterConfig, ITermStoreTag } from '../models/IContainerModels';
import { TaxonomyService } from '../services/TaxonomyService';
import { TermFilterConfigDialog } from './TermFilterConfigDialog';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    alignItems: 'center',
    flexWrap: 'wrap',
    ...shorthands.gap(tokens.spacingHorizontalM),
    width: '100%',
    ...shorthands.padding(tokens.spacingVerticalXS, '0')
  },
  filterGroup: {
    display: 'flex',
    alignItems: 'center',
    flexWrap: 'wrap',
    ...shorthands.gap(tokens.spacingHorizontalS)
  },
  dropdownWrapper: {
    display: 'flex',
    alignItems: 'center',
    position: 'relative',
    ...shorthands.gap('4px')
  },
  dropdown: {
    minWidth: '160px',
    maxWidth: '240px',
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    boxShadow: tokens.shadow2
  },
  summaryPill: {
    display: 'inline-flex',
    alignItems: 'center',
    ...shorthands.gap('6px'),
    ...shorthands.padding('4px', '10px'),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorPaletteGreenBackground1,
    color: tokens.colorPaletteGreenForeground1,
    border: `1px solid ${tokens.colorPaletteGreenBorder1}`,
    fontWeight: 500,
    fontSize: '0.85rem'
  },
  summaryUpdated: {
    fontWeight: 700,
    color: tokens.colorPaletteGreenForeground1
  },
  summaryText: {
    color: tokens.colorNeutralForeground1
  },
  clearButton: {
    marginLeft: tokens.spacingHorizontalXS,
    cursor: 'pointer',
    display: 'inline-flex',
    alignItems: 'center',
    color: tokens.colorNeutralForeground3,
    ':hover': {
      color: tokens.colorPaletteRedForeground1
    }
  },
  editActionBtn: {
    minWidth: '24px',
    height: '24px',
    ...shorthands.padding('0')
  }
});

export interface ITermFilterBarProps {
  filters: ITermFilterConfig[];
  selectedValues: Record<string, string>; // filterId -> selected term label or empty
  onChangeFilter: (filterId: string, termLabel: string) => void;
  onClearAllFilters: () => void;
  isEditMode?: boolean;
  onUpdateFilters?: (newFilters: ITermFilterConfig[]) => void;
}

function renderFilterIcon(iconName?: string): React.ReactElement {
  switch (iconName) {
    case 'Building':
      return <BuildingRegular fontSize={16} />;
    case 'Globe':
      return <GlobeRegular fontSize={16} />;
    case 'Tag':
      return <TagRegular fontSize={16} />;
    case 'Briefcase':
      return <BriefcaseRegular fontSize={16} />;
    case 'People':
      return <PeopleRegular fontSize={16} />;
    case 'BookAnswers':
      return <BookOpenRegular fontSize={16} />;
    case 'Sparkle':
      return <SparkleRegular fontSize={16} />;
    case 'Target':
      return <TargetRegular fontSize={16} />;
    case 'Pin':
    default:
      return <LocationRegular fontSize={16} />;
  }
}

export const TermFilterBar: React.FC<ITermFilterBarProps> = ({
  filters,
  selectedValues,
  onChangeFilter,
  onClearAllFilters,
  isEditMode,
  onUpdateFilters
}) => {
  const styles = useStyles();

  // Cache terms loaded per term set: termSetName -> ITermStoreTag[]
  const [termsCache, setTermsCache] = React.useState<Record<string, ITermStoreTag[]>>({});
  const [isLoadingTerms, setIsLoadingTerms] = React.useState<boolean>(false);

  // Dialog state for adding/editing a filter
  const [dialogOpen, setDialogOpen] = React.useState<boolean>(false);
  const [editingFilter, setEditingFilter] = React.useState<ITermFilterConfig | undefined>(undefined);

  // Load terms for each configured filter
  React.useEffect(() => {
    let isMounted = true;
    const fetchAllTerms = async (): Promise<void> => {
      setIsLoadingTerms(true);
      try {
        const cache: Record<string, ITermStoreTag[]> = {};
        for (const f of filters) {
          if (!cache[f.termSetName]) {
            const terms = await TaxonomyService.getTermsByTermSet(f.termSetName);
            cache[f.termSetName] = terms;
          }
        }
        if (isMounted) {
          setTermsCache(cache);
        }
      } catch (err) {
        console.warn('TermFilterBar: Failed to load taxonomy terms', err);
      } finally {
        if (isMounted) {
          setIsLoadingTerms(false);
        }
      }
    };

    if (filters && filters.length > 0) {
      void fetchAllTerms();
    } else {
      setTermsCache({});
    }

    return () => {
      isMounted = false;
    };
  }, [filters]);

  const handleOptionSelect = (filterId: string, event: SelectionEvents, data: OptionOnSelectData): void => {
    const selectedVal = data.optionValue || '';
    onChangeFilter(filterId, selectedVal);
  };

  const handleSaveFilterConfig = (savedConfig: ITermFilterConfig): void => {
    if (!onUpdateFilters) return;
    const existingIndex = filters.findIndex((f) => f.id === savedConfig.id);
    let updated: ITermFilterConfig[];
    if (existingIndex >= 0) {
      updated = [...filters];
      updated[existingIndex] = savedConfig;
    } else {
      updated = [...filters, savedConfig];
    }
    onUpdateFilters(updated);
  };

  const handleDeleteFilterConfig = (filterId: string): void => {
    if (!onUpdateFilters) return;
    const updated = filters.filter((f) => f.id !== filterId);
    onUpdateFilters(updated);
    // Clear selection if this filter was active
    if (selectedValues[filterId]) {
      onChangeFilter(filterId, '');
    }
  };

  // Compute active selections
  const activeKeys: string[] = Object.keys(selectedValues).filter((k) => {
    const v = selectedValues[k];
    return !!v && v.trim().length > 0;
  });
  const hasActiveFilters = activeKeys.length > 0;

  // Build the summary sentence matching Screenshot 1:
  // e.g. "✓ updated — showing UK requirements for Infrastructure"
  let summarySentence = '';
  if (hasActiveFilters) {
    const activeTerms = activeKeys.map((filterId: string) => {
      const config = filters.find((f) => f.id === filterId);
      const val = selectedValues[filterId] || '';
      return { label: config?.label || 'Item', value: val };
    });

    if (activeTerms.length === 1) {
      summarySentence = `showing ${activeTerms[0].value}`;
    } else if (activeTerms.length === 2) {
      summarySentence = `showing ${activeTerms[0].value} requirements for ${activeTerms[1].value}`;
    } else {
      const parts = activeTerms.map((t: { label: string; value: string }) => t.value).join(', ');
      summarySentence = `showing ${parts}`;
    }
  }


  if ((!filters || filters.length === 0) && !isEditMode) {
    return null;
  }

  return (
    <div className={styles.root}>
      <div className={styles.filterGroup}>
        {filters.map((filter, idx) => {
          const terms = termsCache[filter.termSetName] || [];
          const currentVal = selectedValues[filter.id] || '';

          return (
            <div key={filter.id} className={styles.dropdownWrapper}>
              <Dropdown
                className={styles.dropdown}
                placeholder={filter.placeholder || `Choose a ${filter.label.toLowerCase()}`}
                value={currentVal || undefined}
                selectedOptions={currentVal ? [currentVal] : []}
                onOptionSelect={(e, data) => handleOptionSelect(filter.id, e, data)}
                size="medium"
                clearable
              >
                {terms.map((t) => (
                  <Option key={t.id} value={t.label} text={t.label}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                      {renderFilterIcon(filter.iconName)}
                      <span>{t.label}</span>
                    </div>
                  </Option>
                ))}
              </Dropdown>

              {/* Edit & Reorder buttons in Edit Mode */}
              {isEditMode && onUpdateFilters && (
                <div style={{ display: 'inline-flex', alignItems: 'center', gap: '2px' }}>
                  <Button
                    appearance="subtle"
                    size="small"
                    icon={<ArrowLeftRegular fontSize={13} />}
                    className={styles.editActionBtn}
                    disabled={idx === 0}
                    title={`Move ${filter.label} filter left`}
                    onClick={() => {
                      if (idx === 0) return;
                      const updated = [...filters];
                      const temp = updated[idx - 1];
                      updated[idx - 1] = updated[idx];
                      updated[idx] = temp;
                      onUpdateFilters(updated);
                    }}
                  />
                  <Button
                    appearance="subtle"
                    size="small"
                    icon={<ArrowRightRegular fontSize={13} />}
                    className={styles.editActionBtn}
                    disabled={idx === filters.length - 1}
                    title={`Move ${filter.label} filter right`}
                    onClick={() => {
                      if (idx === filters.length - 1) return;
                      const updated = [...filters];
                      const temp = updated[idx + 1];
                      updated[idx + 1] = updated[idx];
                      updated[idx] = temp;
                      onUpdateFilters(updated);
                    }}
                  />
                  <Button
                    appearance="subtle"
                    size="small"
                    icon={<SettingsRegular fontSize={14} />}
                    className={styles.editActionBtn}
                    title={`Configure ${filter.label} dropdown`}
                    onClick={() => {
                      setEditingFilter(filter);
                      setDialogOpen(true);
                    }}
                  />
                </div>
              )}
            </div>
          );
        })}

        {/* In Edit Mode: + Add filter dropdown button */}
        {isEditMode && onUpdateFilters && (
          <Button
            appearance="subtle"
            size="medium"
            icon={<AddRegular />}
            onClick={() => {
              setEditingFilter(undefined);
              setDialogOpen(true);
            }}
            style={{
              border: `1px dashed ${tokens.colorBrandStroke2}`,
              color: tokens.colorBrandForeground1
            }}
          >
            Add filter dropdown
          </Button>
        )}
      </div>

      {/* Dynamic Selection Summary Tag matching Screenshot 1 */}
      {hasActiveFilters && (
        <div className={styles.summaryPill}>
          <CheckmarkCircleRegular fontSize={16} style={{ color: tokens.colorPaletteGreenForeground1 }} />
          <span className={styles.summaryUpdated}>updated</span>
          <span className={styles.summaryText}>— {summarySentence}</span>
          <span
            className={styles.clearButton}
            onClick={onClearAllFilters}
            title="Clear all filters"
            role="button"
            tabIndex={0}
            onKeyDown={(e) => {
              if (e.key === 'Enter' || e.key === ' ') onClearAllFilters();
            }}
          >
            <DismissRegular fontSize={14} />
          </span>
        </div>
      )}

      {isLoadingTerms && (
        <Spinner size="extra-tiny" label="Loading taxonomy..." labelPosition="after" />
      )}

      {/* In-place Authoring Dialog */}
      <TermFilterConfigDialog
        isOpen={dialogOpen}
        filterConfig={editingFilter}
        onSave={handleSaveFilterConfig}
        onDelete={handleDeleteFilterConfig}
        onDismiss={() => {
          setDialogOpen(false);
          setEditingFilter(undefined);
        }}
      />
    </div>
  );
};
