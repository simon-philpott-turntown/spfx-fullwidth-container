/**
 * @file LiveDataRenderer.tsx
 * @description Dynamic real-time text and metric field component bound to REST/Graph APIs.
 */

import * as React from 'react';
import {
  makeStyles,
  tokens,
  shorthands,
  Badge,
  Button,
  Spinner,
  Caption1,
  Skeleton,
  SkeletonItem
} from '@fluentui/react-components';
import { ArrowSyncRegular, CloudCheckmarkRegular, WarningRegular } from '@fluentui/react-icons';
import { ILiveDataConfig } from '../models/IContainerModels';
import { LiveDataService } from '../services/LiveDataService';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '6px',
    ...shorthands.padding('10px'),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorNeutralBackground2,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2)
  },
  statusHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between'
  },
  valueDisplay: {
    fontSize: '1.4rem',
    fontWeight: 700,
    color: tokens.colorBrandForeground1,
    fontFamily: tokens.fontFamilyNumeric
  },
  syncButton: {
    minWidth: 'auto',
    ...shorthands.padding('2px', '6px')
  }
});

export interface ILiveDataRendererProps {
  config: ILiveDataConfig;
  isEditMode?: boolean;
}

export const LiveDataRenderer: React.FC<ILiveDataRendererProps> = ({ config, isEditMode = false }) => {
  const styles = useStyles();
  const [dataValue, setDataValue] = React.useState<string>('Loading...');
  const [loading, setLoading] = React.useState<boolean>(true);
  const [lastUpdated, setLastUpdated] = React.useState<string>('');
  const [hasError, setHasError] = React.useState<boolean>(false);

  const loadData = React.useCallback(async () => {
    setLoading(true);
    setHasError(false);
    try {
      const val = await LiveDataService.fetchLiveData(config);
      setDataValue(val);
      const now = new Date();
      const pad2 = (n: number): string => (n < 10 ? '0' + n : '' + n);
      setLastUpdated(`${pad2(now.getHours())}:${pad2(now.getMinutes())}:${pad2(now.getSeconds())}`);
    } catch {
      setHasError(true);
      setDataValue(config.fallbackText || '£0.00');
    } finally {
      setLoading(false);
    }
  }, [config]);

  React.useEffect(() => {
    void loadData();

    if (config.refreshIntervalSeconds && config.refreshIntervalSeconds > 0) {
      const interval = setInterval(() => { void loadData(); }, config.refreshIntervalSeconds * 1000);
      return () => clearInterval(interval);
    }
  }, [loadData, config.refreshIntervalSeconds]);

  return (
    <div className={styles.container}>
      <div className={styles.statusHeader}>
        <Badge
          appearance="tint"
          color={hasError ? 'danger' : 'success'}
          icon={hasError ? <WarningRegular /> : <CloudCheckmarkRegular />}
          size="small"
        >
          {hasError ? 'Offline Fallback' : 'Live Connected'}
        </Badge>
        <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
          {lastUpdated && (
            <Caption1 style={{ color: tokens.colorNeutralForeground4, fontSize: '0.72rem' }}>
              {lastUpdated}
            </Caption1>
          )}
          <Button
            size="small"
            appearance="subtle"
            className={styles.syncButton}
            icon={loading ? <Spinner size="extra-tiny" /> : <ArrowSyncRegular />}
            onClick={loadData}
            title="Refresh Live Data"
          />
        </div>
      </div>

      <div className={styles.valueDisplay}>
        {loading ? (
          <Skeleton aria-label="Loading metric value" style={{ width: '100%', padding: '4px 0' }}>
            <SkeletonItem style={{ height: '24px', width: '55%', borderRadius: '4px' }} />
          </Skeleton>
        ) : (
          dataValue
        )}
      </div>

      {isEditMode && config.apiUrl && (
        <Caption1 style={{ color: tokens.colorNeutralForeground3, fontSize: '0.72rem', wordBreak: 'break-all' }}>
          Endpoint: {config.apiUrl} ({config.jsonPath})
        </Caption1>
      )}
    </div>
  );
};
