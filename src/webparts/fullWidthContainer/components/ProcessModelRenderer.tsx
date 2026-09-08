/**
 * @file ProcessModelRenderer.tsx
 * @description Customisable Process Model (Chevron / Stage Flow) component built with Fluent UI 2 (@fluentui/react-components v9).
 * Renders connected process stages (Stage 1: Shape, Stage 2: Plan, Stage 3: Source, Stage 4: Deliver, Stage 5: Realise)
 * with capability badges, descriptions, and interactive click handling (URL navigation or card filtering).
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
  OpenRegular,
  EditRegular,
  DismissRegular,
  ChevronRightRegular
} from '@fluentui/react-icons';
import { IProcessStepItem } from '../models/IContainerModels';

const useStyles = makeStyles({
  container: {
    width: '100%',
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
    boxSizing: 'border-box',
    marginTop: '4px',
    marginBottom: '8px'
  },
  processFlow: {
    display: 'flex',
    flexDirection: 'row',
    width: '100%',
    boxSizing: 'border-box',
    backgroundColor: '#F8F6F2', // Soft mushroom background matching Turner & Townsend palette
    borderRadius: '8px',
    border: '1px solid #E6DFD5',
    overflow: 'hidden',
    position: 'relative',
    '@media (max-width: 768px)': {
      flexDirection: 'column',
      borderRadius: '6px'
    }
  },
  stageItem: {
    flex: 1,
    display: 'flex',
    flexDirection: 'column',
    justifyContent: 'space-between',
    padding: '16px 14px',
    minHeight: '110px',
    cursor: 'pointer',
    position: 'relative',
    backgroundColor: 'transparent',
    border: 'none',
    textAlign: 'left',
    outlineStyle: 'none',
    transitionProperty: 'background-color, color, transform',
    transitionDuration: '180ms',
    ':hover': {
      backgroundColor: '#EFEAE2'
    },
    '@media (max-width: 768px)': {
      borderBottom: '1px solid #E6DFD5',
      minHeight: 'auto'
    }
  },
  stageItemActive: {
    backgroundColor: '#001436 !important', // Deep Turner & Townsend navy
    color: '#FFFFFF !important',
    boxShadow: '0 4px 12px rgba(0, 20, 54, 0.25)',
    zIndex: 2
  },
  stageDivider: {
    position: 'absolute',
    right: 0,
    top: 0,
    bottom: 0,
    width: '14px',
    pointerEvents: 'none',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    color: '#C8C0B4',
    zIndex: 3,
    '@media (max-width: 768px)': {
      display: 'none'
    }
  },
  stageHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    marginBottom: '4px'
  },
  stageDot: {
    width: '7px',
    height: '7px',
    borderRadius: '50%',
    backgroundColor: '#1E4479',
    flexShrink: 0
  },
  stageDotActive: {
    backgroundColor: '#0090DC' // Cyan tint
  },
  stageNumber: {
    fontSize: '0.72rem',
    fontWeight: 600,
    textTransform: 'uppercase',
    letterSpacing: '0.5px',
    color: '#737B80'
  },
  stageNumberActive: {
    color: '#CCE9F8'
  },
  stageTitle: {
    fontSize: '1.05rem',
    fontWeight: 700,
    lineHeight: '1.3',
    color: '#001436',
    marginBottom: '4px'
  },
  stageTitleActive: {
    color: '#FFFFFF'
  },
  stageDescription: {
    fontSize: '0.78rem',
    lineHeight: '1.25',
    color: '#505A60',
    flexGrow: 1,
    marginBottom: '8px'
  },
  stageDescriptionActive: {
    color: '#E2E8F0'
  },
  stageMetric: {
    fontSize: '0.75rem',
    fontWeight: 600,
    color: '#1E4479',
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    marginTop: 'auto'
  },
  stageMetricActive: {
    color: '#80D1F8'
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
    background: 'none',
    border: 'none',
    padding: 0,
    ':hover': {
      textDecoration: 'underline'
    }
  }
});

export interface IProcessModelRendererProps {
  steps: IProcessStepItem[];
  activeStepId?: string;
  alignment?: 'left' | 'center' | 'right';
  onSelectStep?: (step: IProcessStepItem | null) => void;
  isEditMode?: boolean;
  onEdit?: () => void;
}

export const ProcessModelRenderer: React.FC<IProcessModelRendererProps> = ({
  steps,
  activeStepId,
  alignment = 'left',
  onSelectStep,
  isEditMode,
  onEdit
}) => {
  const styles = useStyles();

  const handleStepClick = (step: IProcessStepItem): void => {
    if (step.actionType === 'navigate' && step.url) {
      window.open(step.url, '_blank', 'noopener,noreferrer');
      return;
    }

    // Filter action
    if (activeStepId === step.id) {
      if (onSelectStep) {
        onSelectStep(null);
      }
    } else {
      if (onSelectStep) {
        onSelectStep(step);
      }
    }
  };

  const activeStep = steps.find((s) => s.id === activeStepId);

  return (
    <div className={styles.container}>
      <div className={styles.processFlow} role="navigation" aria-label="Process stages">
        {steps.map((step, idx) => {
          const isActive = activeStepId === step.id;
          const isLast = idx === steps.length - 1;

          return (
            <button
              key={step.id}
              type="button"
              className={`${styles.stageItem} ${isActive ? styles.stageItemActive : ''}`}
              onClick={() => handleStepClick(step)}
              title={
                step.actionType === 'navigate'
                  ? `Navigate to ${step.url || 'link'}`
                  : `Filter cards by "${step.title}"`
              }
              aria-pressed={isActive}
            >
              <div>
                <div className={styles.stageHeader}>
                  <div
                    className={`${styles.stageDot} ${isActive ? styles.stageDotActive : ''}`}
                  />
                  <span
                    className={`${styles.stageNumber} ${
                      isActive ? styles.stageNumberActive : ''
                    }`}
                  >
                    {step.stageNumber || `Stage ${idx + 1}`}
                  </span>
                </div>

                <div
                  className={`${styles.stageTitle} ${
                    isActive ? styles.stageTitleActive : ''
                  }`}
                >
                  {step.title}
                </div>

                {step.description && (
                  <div
                    className={`${styles.stageDescription} ${
                      isActive ? styles.stageDescriptionActive : ''
                    }`}
                  >
                    {step.description}
                  </div>
                )}
              </div>

              {step.metricBadge && (
                <div
                  className={`${styles.stageMetric} ${
                    isActive ? styles.stageMetricActive : ''
                  }`}
                >
                  {step.metricBadge}
                  {step.actionType === 'navigate' && (
                    <OpenRegular style={{ fontSize: '11px', marginLeft: '2px' }} />
                  )}
                </div>
              )}

              {!isLast && (
                <div className={styles.stageDivider}>
                  <ChevronRightRegular fontSize={14} />
                </div>
              )}
            </button>
          );
        })}
      </div>

      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        {activeStep && (
          <div className={styles.activeIndicatorRow}>
            <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
              Process filter: <strong>{activeStep.title}</strong>
            </Caption1>
            <button
              type="button"
              className={styles.clearLink}
              onClick={() => onSelectStep && onSelectStep(null)}
              title="Clear stage filter"
            >
              <DismissRegular style={{ fontSize: '12px' }} /> Clear filter
            </button>
          </div>
        )}

        {isEditMode && onEdit && (
          <Button
            size="small"
            appearance="subtle"
            icon={<EditRegular />}
            onClick={onEdit}
            style={{ marginLeft: 'auto' }}
          >
            Configure process model
          </Button>
        )}
      </div>
    </div>
  );
};
