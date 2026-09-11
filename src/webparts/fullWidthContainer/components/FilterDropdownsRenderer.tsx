/**
 * @file FilterDropdownsRenderer.tsx
 * @description Renders a horizontal row of multiple filter dropdowns built with Fluent UI 2 (@fluentui/react-components v9).
 * Allows page authors to place as many dropdown filters alongside each other as they want.
 * Dispatches dashboard:card-filter-apply events with composite or single filter keys so parent containers filter cards cleanly.
 */

import * as React from 'react';
import {
  Dropdown,
  Option,
  Label,
  Button,
  Caption1,
  makeStyles,
  tokens
} from '@fluentui/react-components';
import {
  DismissRegular,
  EditRegular
} from '@fluentui/react-icons';
import { IDropdownFilterConfig } from '../models/IContainerModels';
import { TaxonomyService } from '../services/TaxonomyService';
import { renderUnifiedIcon } from './CustomSvgIconRegistry';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
    width: '100%',
    boxSizing: 'border-box',
    paddingTop: '4px',
    paddingBottom: '4px'
  },
  dropdownsRow: {
    display: 'flex',
    alignItems: 'flex-end',
    flexWrap: 'wrap',
    gap: '12px'
  },
  dropdownCol: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
    minWidth: '200px',
    flex: '0 1 auto'
  },
  activeIndicatorRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    flexWrap: 'wrap',
    marginTop: '2px'
  },
  clearLink: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    fontSize: '0.75rem',
    fontWeight: 600,
    color: '#1E4479',
    cursor: 'pointer',
    background: 'none',
    border: 'none',
    padding: 0,
    textDecoration: 'none',
    ':hover': {
      textDecoration: 'underline'
    }
  }
});

interface IDropdownItemRendererProps {
  config: IDropdownFilterConfig;
  onSelectValue: (val: string | undefined) => void;
}

const IndividualFilterDropdown: React.FC<IDropdownItemRendererProps> = ({ config, onSelectValue }) => {
  const [loadedOptions, setLoadedOptions] = React.useState<Array<{ label: string; value: string }>>(config.options || []);

  React.useEffect(() => {
    let isMounted = true;
    if (config.termSetName && config.termSetName.trim()) {
      TaxonomyService.getTermsByTermSet(config.termSetName.trim())
        .then((terms) => {
          if (isMounted && terms && terms.length > 0) {
            setLoadedOptions(terms.map((t) => ({ label: t.label, value: t.label })));
          }
        })
        .catch((err) => console.warn('[FilterDropdownsRenderer] Failed loading termset', err));
    } else {
      setLoadedOptions(config.options || []);
    }
    return () => {
      isMounted = false;
    };
  }, [config.termSetName, config.options]);

  const currentOptions = loadedOptions.length > 0 ? loadedOptions : (config.options || []);
  const activeLabel = currentOptions.find((o) => o.value === config.selectedValue)?.label || '';

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', minWidth: '200px' }}>
      {config.label && (
        <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
          {config.iconName && (
            <span style={{ fontSize: '15px', color: tokens.colorBrandForeground1, display: 'inline-flex' }}>
              {renderUnifiedIcon(config.iconName)}
            </span>
          )}
          <Label size="small" weight="semibold">{config.label}</Label>
        </div>
      )}
      <Dropdown
        size="medium"
        placeholder={config.placeholder || 'Select option...'}
        selectedOptions={config.selectedValue ? [config.selectedValue] : []}
        value={activeLabel}
        onOptionSelect={(_, data) => {
          onSelectValue(data.optionValue || undefined);
        }}
        style={{ minWidth: '200px' }}
      >
        {currentOptions.map((opt) => (
          <Option key={opt.value} value={opt.value}>
            {opt.label}
          </Option>
        ))}
      </Dropdown>
    </div>
  );
};

export interface IFilterDropdownsRendererProps {
  itemId: string;
  dropdowns: IDropdownFilterConfig[];
  alignment?: 'left' | 'center' | 'right';
  isEditMode?: boolean;
  onUpdateDropdowns?: (updated: IDropdownFilterConfig[]) => void;
  onEdit?: () => void;
}

export const FilterDropdownsRenderer: React.FC<IFilterDropdownsRendererProps> = ({
  itemId,
  dropdowns,
  alignment = 'left',
  isEditMode,
  onUpdateDropdowns,
  onEdit
}) => {
  const styles = useStyles();

  const handleSelectDropdownValue = (index: number, val: string | undefined): void => {
    const next = [...dropdowns];
    next[index] = { ...next[index], selectedValue: val };
    if (onUpdateDropdowns) {
      onUpdateDropdowns(next);
    }

    // Active filters summary
    const activeFilters = next
      .filter((d) => !!d.selectedValue)
      .map((d) => d.selectedValue as string);

    window.dispatchEvent(
      new CustomEvent('dashboard:card-filter-apply', {
        detail: {
          sourceItemId: itemId,
          filterValue: activeFilters.join(' ')
        }
      })
    );
  };

  const handleClearAll = (): void => {
    const cleared = dropdowns.map((d) => ({ ...d, selectedValue: undefined }));
    if (onUpdateDropdowns) {
      onUpdateDropdowns(cleared);
    }
    window.dispatchEvent(
      new CustomEvent('dashboard:card-filter-apply', {
        detail: {
          sourceItemId: itemId,
          filterValue: ''
        }
      })
    );
  };

  const justifyVal = alignment === 'center' ? 'center' : alignment === 'right' ? 'flex-end' : 'flex-start';
  const hasActiveSelections = dropdowns.some((d) => !!d.selectedValue);

  return (
    <div className={styles.container}>
      <div className={styles.dropdownsRow} style={{ justifyContent: justifyVal }}>
        {dropdowns.map((drop, idx) => (
          <div key={drop.id || idx} className={styles.dropdownCol}>
            <IndividualFilterDropdown
              config={drop}
              onSelectValue={(val) => handleSelectDropdownValue(idx, val)}
            />
          </div>
        ))}

        {isEditMode && onEdit && (
          <div style={{ marginBottom: '2px' }}>
            <Button
              size="small"
              appearance="subtle"
              icon={<EditRegular />}
              onClick={onEdit}
              title="Configure filter dropdowns"
            >
              Configure dropdowns ({dropdowns.length})
            </Button>
          </div>
        )}
      </div>

      {hasActiveSelections && (
        <div className={styles.activeIndicatorRow} style={{ justifyContent: justifyVal }}>
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
            Filters applied:
          </Caption1>
          {dropdowns
            .filter((d) => !!d.selectedValue)
            .map((d) => (
              <Caption1 key={d.id} style={{ fontWeight: 600, color: tokens.colorBrandForeground1 }}>
                {d.label ? `${d.label}: ` : ''}{d.selectedValue}
              </Caption1>
            ))}
          <button
            type="button"
            className={styles.clearLink}
            onClick={handleClearAll}
            title="Clear all dropdown filters"
          >
            <DismissRegular style={{ fontSize: '12px' }} /> Clear all
          </button>
        </div>
      )}
    </div>
  );
};
