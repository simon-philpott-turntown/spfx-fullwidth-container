/**
 * @file TermFilterConfigDialog.tsx
 * @description In-place modal dialog for configuring Global Term Store filter dropdowns
 * using Fluent UI 2 (@fluentui/react-components v9). Allows dashboard authors to select
 * a term group, term set, placeholder text, label, and custom icon.
 * Wrapped in Fluent UI 2 Portal with zIndex: 1000000 to prevent clipping in SharePoint edit mode.
 */

import * as React from 'react';
import {
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogContent,
  Button,
  Input,
  Field,
  Select,
  Portal,
  makeStyles,
  tokens,
  shorthands
} from '@fluentui/react-components';
import {
  DismissRegular,
  CheckmarkRegular,
  DeleteRegular,
  FilterRegular
} from '@fluentui/react-icons';
import { ITermFilterConfig } from '../models/IContainerModels';
import { TaxonomyService, ITermGroup } from '../services/TaxonomyService';

const useStyles = makeStyles({
  surface: {
    maxWidth: '520px',
    width: '90vw',
    zIndex: 1000000,
    boxShadow: tokens.shadow16
  },
  form: {
    display: 'flex',
    flexDirection: 'column',
    // ...shorthands.gap(tokens.spacingVerticalM), // [USER_TEST: comment out vertical spacing in flex containers]
    marginTop: tokens.spacingVerticalS
  },
  actions: {
    display: 'flex',
    justifyContent: 'space-between',
    width: '100%'
  },
  leftActions: {
    display: 'flex',
    alignItems: 'center'
  },
  rightActions: {
    display: 'flex',
    ...shorthands.gap(tokens.spacingHorizontalS)
  }
});

export interface ITermFilterConfigDialogProps {
  isOpen: boolean;
  filterConfig?: ITermFilterConfig;
  onSave: (config: ITermFilterConfig) => void;
  onDelete?: (filterId: string) => void;
  onDismiss: () => void;
}

const AVAILABLE_ICONS = [
  { id: 'Pin', label: 'Location Pin (Pin)' },
  { id: 'Building', label: 'Building / Sector' },
  { id: 'Globe', label: 'Globe / Global' },
  { id: 'Tag', label: 'Tag / Category' },
  { id: 'Briefcase', label: 'Briefcase / Business' },
  { id: 'People', label: 'People / Team' },
  { id: 'BookAnswers', label: 'Book / Methodology' },
  { id: 'Sparkle', label: 'Sparkle / Innovation' },
  { id: 'Target', label: 'Target / Goals' }
];

export const TermFilterConfigDialog: React.FC<ITermFilterConfigDialogProps> = ({
  isOpen,
  filterConfig,
  onSave,
  onDelete,
  onDismiss
}) => {
  const styles = useStyles();

  const [termGroups, setTermGroups] = React.useState<ITermGroup[]>([]);
  const [selectedGroupName, setSelectedGroupName] = React.useState<string>('Our business');
  const [selectedTermSetName, setSelectedTermSetName] = React.useState<string>('Our segments');
  const [label, setLabel] = React.useState<string>('Segment');
  const [placeholder, setPlaceholder] = React.useState<string>('Choose a segment');
  const [iconName, setIconName] = React.useState<string>('Building');

  // Load term groups and populate defaults
  React.useEffect(() => {
    let isMounted = true;
    const loadGroups = async (): Promise<void> => {
      try {
        const groups = await TaxonomyService.getTermGroups();
        if (isMounted) {
          setTermGroups(groups);
        }
      } catch (err) {
        console.warn('Failed to load taxonomy term groups:', err);
      }
    };
    void loadGroups();

    return () => {
      isMounted = false;
    };
  }, []);

  // Synchronize state when dialog opens or editing filter changes
  React.useEffect(() => {
    if (isOpen) {
      if (filterConfig) {
        setLabel(filterConfig.label || 'Filter');
        setPlaceholder(filterConfig.placeholder || 'Choose an option');
        setSelectedGroupName(filterConfig.termGroupName || 'Our business');
        setSelectedTermSetName(filterConfig.termSetName || 'Our segments');
        setIconName(filterConfig.iconName || 'Pin');
      } else {
        // Defaults for new filter
        setLabel('Segment');
        setPlaceholder('Choose a segment');
        setSelectedGroupName('Our business');
        setSelectedTermSetName('Our segments');
        setIconName('Building');
      }
    }
  }, [isOpen, filterConfig]);

  // Find available term sets under current group
  const activeGroup = termGroups.find(
    (g) => g.name.toLowerCase() === selectedGroupName.toLowerCase()
  ) || (termGroups.length > 0 ? termGroups[0] : undefined);

  const availableSets = activeGroup ? activeGroup.termSets : [];

  const handleGroupChange = (e: React.ChangeEvent<HTMLSelectElement>): void => {
    const newGroup = e.target.value;
    setSelectedGroupName(newGroup);
    const grp = termGroups.find((g) => g.name === newGroup);
    if (grp && grp.termSets.length > 0) {
      setSelectedTermSetName(grp.termSets[0].name);
      setPlaceholder(`Choose a ${grp.termSets[0].name.toLowerCase().replace(/^our\s+/i, '')}`);
    }
  };

  const handleSetChange = (e: React.ChangeEvent<HTMLSelectElement>): void => {
    const newSet = e.target.value;
    setSelectedTermSetName(newSet);
    const cleanName = newSet.replace(/^our\s+/i, '');
    setLabel(cleanName.charAt(0).toUpperCase() + cleanName.slice(1));
    setPlaceholder(`Choose a ${cleanName.toLowerCase()}`);
  };

  const handleSave = (): void => {
    const config: ITermFilterConfig = {
      id: filterConfig ? filterConfig.id : `filter-${Date.now()}`,
      label: label.trim() || 'Filter',
      placeholder: placeholder.trim() || 'Select an option',
      termGroupName: selectedGroupName,
      termSetName: selectedTermSetName,
      iconName: iconName
    };
    onSave(config);
    onDismiss();
  };

  const handleDelete = (): void => {
    if (filterConfig && onDelete) {
      onDelete(filterConfig.id);
      onDismiss();
    }
  };

  return (
    <Portal>
      <Dialog open={isOpen} onOpenChange={(e, data) => { if (!data.open) onDismiss(); }}>
        <DialogSurface className={styles.surface}>
          <DialogTitle
            action={
              <Button
                appearance="subtle"
                aria-label="close"
                icon={<DismissRegular />}
                onClick={onDismiss}
              />
            }
          >
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <FilterRegular fontSize={20} />
              <span>{filterConfig ? 'Configure filter dropdown' : 'Add filter dropdown'}</span>
            </div>
          </DialogTitle>

          <DialogBody>
            <DialogContent className={styles.form}>
              {/* Global Term Store Group */}
              <Field label="Term Store group">
                <Select
                  value={selectedGroupName}
                  onChange={handleGroupChange}
                >
                  {termGroups.map((g) => (
                    <option key={g.id} value={g.name}>
                      {g.name}
                    </option>
                  ))}
                </Select>
              </Field>

              {/* Term Set */}
              <Field label="Source term set">
                <Select
                  value={selectedTermSetName}
                  onChange={handleSetChange}
                >
                  {availableSets.map((s) => (
                    <option key={s.id} value={s.name}>
                      {s.name} ({s.terms.length} terms)
                    </option>
                  ))}
                </Select>
              </Field>

              {/* Label / Filter Category */}
              <Field label="Filter category label">
                <Input
                  value={label}
                  onChange={(e, data) => setLabel(data.value)}
                  placeholder="e.g. Region, Segment, Sector"
                />
              </Field>

              {/* Placeholder Text */}
              <Field label="Dropdown placeholder / holding text">
                <Input
                  value={placeholder}
                  onChange={(e, data) => setPlaceholder(data.value)}
                  placeholder="e.g. Choose a segment"
                />
              </Field>

              {/* Icon */}
              <Field label="Dropdown icon">
                <Select
                  value={iconName}
                  onChange={(e) => setIconName(e.target.value)}
                >
                  {AVAILABLE_ICONS.map((ic) => (
                    <option key={ic.id} value={ic.id}>
                      {ic.label}
                    </option>
                  ))}
                </Select>
              </Field>
            </DialogContent>

            <DialogActions className={styles.actions}>
              <div className={styles.leftActions}>
                {filterConfig && onDelete && (
                  <Button
                    appearance="subtle"
                    icon={<DeleteRegular />}
                    onClick={handleDelete}
                    style={{ color: tokens.colorPaletteRedForeground1 }}
                  >
                    Delete filter
                  </Button>
                )}
              </div>
              <div className={styles.rightActions}>
                <Button appearance="secondary" onClick={onDismiss}>
                  Cancel
                </Button>
                <Button
                  appearance="primary"
                  icon={<CheckmarkRegular />}
                  onClick={handleSave}
                >
                  Save filter
                </Button>
              </div>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </Portal>
  );
};
