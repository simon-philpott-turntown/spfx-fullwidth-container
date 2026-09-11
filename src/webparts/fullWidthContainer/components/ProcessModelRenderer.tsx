/**
 * @file ProcessModelRenderer.tsx
 * @description Customisable Process Model (Interlocking Chevron Arrow Flow) component built with Fluent UI 2.
 * Renders connected chevron arrow stages (Stage 1: Shape, Stage 2: Plan, Stage 3: Source, Stage 4: Deliver, Stage 5: Realise)
 * with the first stage having a flat left end cap, the last stage having a flat right end cap, and intermediate stages
 * having interlocking chevron arrow points and notches matching the Turner & Townsend brand styling.
 */

import * as React from 'react';
import {
  makeStyles,
  tokens,
  Caption1
} from '@fluentui/react-components';
import {
  OpenRegular,
  DismissRegular
} from '@fluentui/react-icons';
import { IProcessStepItem } from '../models/IContainerModels';

const useStyles = makeStyles({
  container: {
    width: '100%',
    display: 'flex',
    flexDirection: 'column',
    gap: '10px',
    boxSizing: 'border-box',
    marginTop: '6px',
    marginBottom: '8px'
  },
  processFlow: {
    display: 'flex',
    flexDirection: 'row',
    width: '100%',
    boxSizing: 'border-box',
    backgroundColor: '#FFFFFF',
    borderRadius: '8px',
    border: '1px solid #E2E8F0',
    overflow: 'hidden',
    boxShadow: '0 2px 8px rgba(0, 0, 0, 0.04)',
    '@media (max-width: 820px)': {
      flexDirection: 'column',
      borderRadius: '6px'
    }
  },
  stageItem: {
    flex: 1,
    display: 'flex',
    flexDirection: 'column',
    justifyContent: 'space-between',
    minHeight: '112px',
    cursor: 'pointer',
    position: 'relative',
    backgroundColor: '#FFFFFF',
    border: 'none',
    textAlign: 'left',
    outlineStyle: 'none',
    boxSizing: 'border-box',
    transitionProperty: 'background-color, color, box-shadow',
    transitionDuration: '160ms',
    transitionTimingFunction: 'ease-in-out',
    ':hover': {
      backgroundColor: '#F8FAFC'
    },
    '@media (max-width: 820px)': {
      clipPath: 'none !important' as any,
      marginLeft: '0 !important',
      paddingLeft: '16px !important',
      paddingRight: '16px !important',
      borderBottom: '1px solid #E2E8F0',
      minHeight: 'auto'
    }
  },
  stageItemActive: {
    backgroundColor: '#001436 !important', // Deep Turner & Townsend navy
    color: '#FFFFFF !important',
    zIndex: 4,
    ':hover': {
      backgroundColor: '#001E4D !important'
    }
  },
  stageHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    marginBottom: '4px'
  },
  stageDot: {
    width: '8px',
    height: '8px',
    borderRadius: '50%',
    backgroundColor: '#94A3B8',
    flexShrink: 0
  },
  stageDotActive: {
    backgroundColor: '#EAA023 !important' // Warm amber indicator dot
  },
  stageNumber: {
    fontSize: '0.72rem',
    fontWeight: 600,
    textTransform: 'uppercase',
    letterSpacing: '0.5px',
    color: '#64748B'
  },
  stageNumberActive: {
    color: '#93C5FD'
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
    lineHeight: '1.28',
    color: '#475569',
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
    color: '#93C5FD'
  },
  activeIndicatorRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
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

    // Filter action toggle
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
          const isFirst = idx === 0;
          const isLast = idx === steps.length - 1;
          const isSingle = steps.length === 1;

          // Arrow chevron geometry with flat end caps:
          // 1. Single item: completely flat rectangular card
          // 2. First stage: flat left edge, pointy arrow right edge
          // 3. Middle stages: notched arrow left edge, pointy arrow right edge
          // 4. Last stage: notched arrow left edge, flat right edge
          let clipPath = 'none';
          let marginLeft = '0px';
          let paddingLeft = '18px';
          let paddingRight = '18px';

          if (!isSingle) {
            if (isFirst) {
              // Flat left edge, arrow point on right
              clipPath = 'polygon(0% 0%, calc(100% - 18px) 0%, 100% 50%, calc(100% - 18px) 100%, 0% 100%)';
              marginLeft = '0px';
              paddingLeft = '18px';
              paddingRight = '32px';
            } else if (isLast) {
              // Arrow notch on left, flat right edge
              clipPath = 'polygon(0% 0%, 18px 50%, 0% 100%, 100% 100%, 100% 0%)';
              marginLeft = '-15px';
              paddingLeft = '30px';
              paddingRight = '18px';
            } else {
              // Arrow notch on left, arrow point on right
              clipPath = 'polygon(0% 0%, calc(100% - 18px) 0%, 100% 50%, calc(100% - 18px) 100%, 0% 100%, 18px 50%)';
              marginLeft = '-15px';
              paddingLeft = '30px';
              paddingRight = '32px';
            }
          }

          // Layer order: Active step gets high z-index, otherwise natural cascading z-index so chevrons nest seamlessly
          const zIndex = isActive ? 10 : steps.length - idx;

          return (
            <button
              key={step.id}
              type="button"
              className={`${styles.stageItem} ${isActive ? styles.stageItemActive : ''}`}
              style={{
                clipPath,
                marginLeft,
                paddingTop: '16px',
                paddingBottom: '16px',
                paddingLeft,
                paddingRight,
                zIndex
              }}
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
                    <OpenRegular style={{ fontSize: '11px', marginLeft: '3px' }} />
                  )}
                </div>
              )}
            </button>
          );
        })}
      </div>

      {activeStep && (
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
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
        </div>
      )}
    </div>
  );
};

