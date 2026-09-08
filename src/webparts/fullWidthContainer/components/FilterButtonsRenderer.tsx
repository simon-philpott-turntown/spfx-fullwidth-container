/**
 * @file FilterButtonsRenderer.tsx
 * @description Interactive Filter Buttons (Pill Filters) component built with Fluent UI 2 (@fluentui/react-components v9).
 * Renders rounded filter pills matching the Turner & Townsend UI design system.
 * Clicking a pill toggles its active filtering state and filters cards in the active view.
 */

import * as React from 'react';
import {
  makeStyles,
  tokens,
  shorthands,
  Button,
  Caption1
} from '@fluentui/react-components';
import {
  FilterRegular,
  DismissRegular,
  EditRegular,
  AddRegular
} from '@fluentui/react-icons';
import { IFilterButtonItem } from '../models/IContainerModels';

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
  pillsRow: {
    display: 'flex',
    alignItems: 'center',
    flexWrap: 'wrap',
    gap: '8px'
  },
  pillButton: {
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    borderRadius: '9999px',
    paddingTop: '6px',
    paddingBottom: '6px',
    paddingLeft: '16px',
    paddingRight: '16px',
    fontSize: '0.875rem',
    fontWeight: 500,
    lineHeight: '1.25rem',
    cursor: 'pointer',
    ...shorthands.border('1px', 'solid', '#D2DAE4'),
    backgroundColor: '#FFFFFF',
    color: '#292929',
    transitionProperty: 'all',
    transitionDuration: '150ms',
    transitionTimingFunction: 'ease-in-out',
    boxShadow: '0 1px 2px rgba(0, 0, 0, 0.04)',
    ':hover': {
      backgroundColor: '#F2EEE7',
      ...shorthands.borderColor('#A5B4C9'),
      transform: 'translateY(-1px)',
      boxShadow: '0 2px 4px rgba(0, 0, 0, 0.08)'
    }
  },
  pillButtonActive: {
    backgroundColor: '#1E4479', // Turner & Townsend Deep Blue
    ...shorthands.borderColor('#1E4479'),
    color: '#FFFFFF',
    fontWeight: 600,
    boxShadow: '0 2px 6px rgba(30, 68, 121, 0.35)',
    ':hover': {
      backgroundColor: '#001436',
      ...shorthands.borderColor('#001436'),
      color: '#FFFFFF'
    }
  },
  activeIndicatorRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
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
    textDecoration: 'none',
    ':hover': {
      textDecoration: 'underline'
    }
  },
  editActionsRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    marginTop: '4px'
  }
});

export interface IFilterButtonsRendererProps {
  buttons: IFilterButtonItem[];
  activeFilterId?: string;
  onSelectFilter?: (button: IFilterButtonItem | null) => void;
  isEditMode?: boolean;
  onEdit?: () => void;
}

export const FilterButtonsRenderer: React.FC<IFilterButtonsRendererProps> = ({
  buttons,
  activeFilterId,
  onSelectFilter,
  isEditMode,
  onEdit
}) => {
  const styles = useStyles();

  const handleButtonClick = (button: IFilterButtonItem): void => {
    if (activeFilterId === button.id) {
      // Toggle off
      if (onSelectFilter) {
        onSelectFilter(null);
      }
    } else {
      if (onSelectFilter) {
        onSelectFilter(button);
      }
    }
  };

  const activeBtn = buttons.find((b) => b.id === activeFilterId);

  return (
    <div className={styles.container}>
      <div className={styles.pillsRow} role="toolbar" aria-label="Filter cards">
        {buttons.map((btn) => {
          const isActive = activeFilterId === btn.id;
          return (
            <button
              key={btn.id}
              type="button"
              className={`${styles.pillButton} ${isActive ? styles.pillButtonActive : ''}`}
              onClick={() => handleButtonClick(btn)}
              aria-pressed={isActive}
              title={`Filter cards by "${btn.label}"`}
            >
              {btn.label}
            </button>
          );
        })}

        {isEditMode && onEdit && (
          <Button
            size="small"
            appearance="subtle"
            icon={<EditRegular />}
            onClick={onEdit}
            title="Configure filter buttons"
          >
            Configure buttons
          </Button>
        )}
      </div>

      {activeBtn && (
        <div className={styles.activeIndicatorRow}>
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
            Filtered by: <strong>{activeBtn.label}</strong>
          </Caption1>
          <button
            type="button"
            className={styles.clearLink}
            onClick={() => onSelectFilter && onSelectFilter(null)}
            title="Clear filter"
          >
            <DismissRegular style={{ fontSize: '12px' }} /> Clear
          </button>
        </div>
      )}
    </div>
  );
};
