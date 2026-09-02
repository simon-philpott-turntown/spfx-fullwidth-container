/**
 * @file CardEditDialog.tsx
 * @description In-place Fluent UI 2 Modal Dialog for editing all card properties directly on canvas.
 * Follows the Microsoft Fluent 2 Design System.
 */

import * as React from 'react';
import {
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogContent,
  DialogActions,
  Button,
  Input,
  Textarea,
  Label,
  Dropdown,
  Option,
  makeStyles,
  shorthands,
  tokens
} from '@fluentui/react-components';
import {
  DismissRegular,
  SaveRegular,
  MoneyRegular,
  ShieldCheckmarkRegular,
  DocumentRegular,
  AppsRegular,
  WrenchRegular,
  GlobeRegular,
  ArrowTrendingLinesRegular,
  LockClosedRegular,
  CheckmarkCircleRegular
} from '@fluentui/react-icons';
import { IContentBlock, BlockType } from '../models/IContainerModels';

const useStyles = makeStyles({
  dialogSurface: {
    maxWidth: '560px',
    width: '90vw',
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: tokens.shadow28
  },
  contentGrid: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap(tokens.spacingVerticalM),
    marginTop: tokens.spacingVerticalS
  },
  fieldRow: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap(tokens.spacingVerticalXS)
  },
  twoColRow: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    ...shorthands.gap(tokens.spacingHorizontalM)
  },
  iconPreviewBox: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap(tokens.spacingHorizontalS),
    ...shorthands.padding(tokens.spacingVerticalXS, tokens.spacingHorizontalS),
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorNeutralBackground2
  }
});

export interface ICardEditDialogProps {
  isOpen: boolean;
  block: IContentBlock | null;
  onSave: (updatedBlock: IContentBlock) => void;
  onDismiss: () => void;
}

const AVAILABLE_ICONS = [
  { key: 'BookAnswers', label: 'Book / Documentation', icon: <DocumentRegular /> },
  { key: 'Financial', label: 'Financial / Money (£)', icon: <MoneyRegular /> },
  { key: 'ComplianceAudit', label: 'Shield / Governance', icon: <ShieldCheckmarkRegular /> },
  { key: 'CheckList', label: 'Checklist / Verified', icon: <CheckmarkCircleRegular /> },
  { key: 'TimelineProgress', label: 'Trending Lines / Metrics', icon: <ArrowTrendingLinesRegular /> },
  { key: 'Calculator', label: 'Wrench / Engineering', icon: <WrenchRegular /> },
  { key: 'AppIconDefault', label: 'Apps / Tool Grid', icon: <AppsRegular /> },
  { key: 'Lock', label: 'Lock / Security', icon: <LockClosedRegular /> },
  { key: 'Globe', label: 'Globe / Portal', icon: <GlobeRegular /> }
];

export const CardEditDialog: React.FC<ICardEditDialogProps> = ({
  isOpen,
  block,
  onSave,
  onDismiss
}) => {
  const styles = useStyles();

  const [formData, setFormData] = React.useState<Partial<IContentBlock>>({});

  React.useEffect(() => {
    if (block) {
      setFormData({ ...block });
    }
  }, [block]);

  if (!isOpen || !block) {
    return null;
  }

  const handleSave = (): void => {
    if (block && formData) {
      onSave({
        ...block,
        ...formData
      } as IContentBlock);
    }
    onDismiss();
  };

  return (
    <Dialog open={isOpen} onOpenChange={(e, data) => { if (!data.open) onDismiss(); }}>
      <DialogSurface className={styles.dialogSurface}>
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
          Edit Card Properties
        </DialogTitle>

        <DialogBody>
          <DialogContent className={styles.contentGrid}>
            {/* Block Type */}
            <div className={styles.twoColRow}>
              <div className={styles.fieldRow}>
                <Label required weight="semibold">Card Type</Label>
                <Dropdown
                  value={
                    formData.type === 'metric'
                      ? 'British Metric Stat (£)'
                      : formData.type === 'embed'
                      ? 'Embed / Tool'
                      : 'Standard Content Card'
                  }
                  onOptionSelect={(e, data) => {
                    setFormData({ ...formData, type: data.optionValue as BlockType });
                  }}
                >
                  <Option value="card">Standard Content Card</Option>
                  <Option value="metric">British Metric Stat (£)</Option>
                  <Option value="embed">Embed / Tool</Option>
                </Dropdown>
              </div>

              {/* Icon Selection */}
              <div className={styles.fieldRow}>
                <Label weight="semibold">Icon</Label>
                <Dropdown
                  value={
                    AVAILABLE_ICONS.find((i) => i.key === formData.iconName)?.label || 'Globe / Portal'
                  }
                  onOptionSelect={(e, data) => {
                    setFormData({ ...formData, iconName: data.optionValue });
                  }}
                >
                  {AVAILABLE_ICONS.map((item) => (
                    <Option key={item.key} value={item.key}>
                      {item.label}
                    </Option>
                  ))}
                </Dropdown>
              </div>
            </div>

            {/* Title */}
            <div className={styles.fieldRow}>
              <Label required weight="semibold">Title</Label>
              <Input
                value={formData.title || ''}
                placeholder="Enter title..."
                onChange={(e, data) => setFormData({ ...formData, title: data.value })}
              />
            </div>

            {/* Description */}
            <div className={styles.fieldRow}>
              <Label weight="semibold">Description / Body Text</Label>
              <Textarea
                rows={3}
                value={formData.description || ''}
                placeholder="Enter description..."
                onChange={(e, data) => setFormData({ ...formData, description: data.value })}
              />
            </div>

            {/* Metric Fields (if type === 'metric') */}
            {formData.type === 'metric' && (
              <div className={styles.twoColRow}>
                <div className={styles.fieldRow}>
                  <Label weight="semibold">Metric Value (GBP £)</Label>
                  <Input
                    value={formData.metricValue || ''}
                    placeholder="e.g. £450,000"
                    onChange={(e, data) => setFormData({ ...formData, metricValue: data.value })}
                  />
                </div>
                <div className={styles.fieldRow}>
                  <Label weight="semibold">Trend Badge</Label>
                  <Input
                    value={formData.metricTrend || ''}
                    placeholder="e.g. +14.2% YoY"
                    onChange={(e, data) => setFormData({ ...formData, metricTrend: data.value })}
                  />
                </div>
              </div>
            )}

            {/* Embed URL (if type === 'embed') */}
            {formData.type === 'embed' && (
              <div className={styles.fieldRow}>
                <Label weight="semibold">iFrame Embed URL</Label>
                <Input
                  value={formData.embedUrl || ''}
                  placeholder="https://..."
                  onChange={(e, data) => setFormData({ ...formData, embedUrl: data.value })}
                />
              </div>
            )}

            {/* Badge & Tags */}
            <div className={styles.twoColRow}>
              <div className={styles.fieldRow}>
                <Label weight="semibold">Status / Category Badge</Label>
                <Input
                  value={formData.badge || ''}
                  placeholder="e.g. Core Framework, New, Guidance"
                  onChange={(e, data) => setFormData({ ...formData, badge: data.value })}
                />
              </div>
              <div className={styles.fieldRow}>
                <Label weight="semibold">Tags (comma-separated)</Label>
                <Input
                  value={formData.tags ? formData.tags.join(', ') : ''}
                  placeholder="e.g. ISO 19650, BIM, Standards"
                  onChange={(e, data) => {
                    const splitTags = data.value.split(',').map((t) => t.trim()).filter((t) => !!t);
                    setFormData({ ...formData, tags: splitTags });
                  }}
                />
              </div>
            </div>

            {/* Link & Action */}
            <div className={styles.twoColRow}>
              <div className={styles.fieldRow}>
                <Label weight="semibold">Button / Link Text</Label>
                <Input
                  value={formData.linkText || ''}
                  placeholder="e.g. Open Playbook, Learn More"
                  onChange={(e, data) => setFormData({ ...formData, linkText: data.value })}
                />
              </div>
              <div className={styles.fieldRow}>
                <Label weight="semibold">Destination URL</Label>
                <Input
                  value={formData.linkUrl || ''}
                  placeholder="https://..."
                  onChange={(e, data) => setFormData({ ...formData, linkUrl: data.value })}
                />
              </div>
            </div>
          </DialogContent>

          <DialogActions>
            <Button appearance="secondary" onClick={onDismiss}>
              Cancel
            </Button>
            <Button appearance="primary" icon={<SaveRegular />} onClick={handleSave}>
              Save Card
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};
