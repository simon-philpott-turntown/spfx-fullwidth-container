/**
 * @file SectionEditDialog.tsx
 * @description In-place Fluent UI 2 Right-Hand Side Property Panel for editing section properties
 * (Title, Badge, Icon, Glyph Color, Background Box, and Background Color).
 * Matches the CardEditDialog slide-out architecture without blocking/blurring underlying dashboard items.
 */

import * as React from 'react';
import {
  Button,
  Input,
  Textarea,
  Label,
  makeStyles,
  shorthands,
  tokens,
  Subtitle2,
  Caption1,
  Checkbox
} from '@fluentui/react-components';
import { renderUnifiedIcon } from './CustomSvgIconRegistry';
import {
  DismissRegular,
  SaveRegular,
  AppsRegular,
  DeleteRegular
} from '@fluentui/react-icons';
import { IContainerSection } from '../models/IContainerModels';
import { FluentIconPicker } from './FluentIconPicker';
import { BrandColorPickerPopover } from './BrandColorPickerPopover';

const useStyles = makeStyles({
  backdrop: {
    position: 'fixed',
    top: 0,
    left: 0,
    width: '100vw',
    height: '100vh',
    backgroundColor: 'transparent',
    zIndex: 999,
    display: 'flex',
    justifyContent: 'flex-end',
    pointerEvents: 'none'
  },
  sidePanel: {
    position: 'relative',
    maxWidth: '92vw',
    height: '100vh',
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.borderLeft('1px', 'solid', tokens.colorNeutralStroke1),
    boxShadow: tokens.shadow28,
    display: 'flex',
    flexDirection: 'column',
    boxSizing: 'border-box',
    pointerEvents: 'auto'
  },
  leftResizeHandle: {
    position: 'absolute',
    left: '-4px',
    top: 0,
    bottom: 0,
    width: '8px',
    cursor: 'ew-resize',
    zIndex: 100,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    ':hover': {
      backgroundColor: tokens.colorBrandStroke1
    }
  },
  panelHeader: {
    padding: '16px 18px 12px 18px',
    ...shorthands.borderBottom('1px', 'solid', tokens.colorNeutralStroke2),
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'flex-start'
  },
  panelBody: {
    flex: 1,
    padding: '16px 18px',
    overflowY: 'auto',
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap(tokens.spacingVerticalM)
  },
  panelFooter: {
    padding: '12px 18px',
    ...shorthands.borderTop('1px', 'solid', tokens.colorNeutralStroke2),
    backgroundColor: tokens.colorNeutralBackground2,
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center'
  },
  fieldRow: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap(tokens.spacingVerticalXXS)
  },
  twoColRow: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    ...shorthands.gap(tokens.spacingHorizontalM)
  },
  iconPreviewBox: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    ...shorthands.padding(tokens.spacingVerticalXS, tokens.spacingHorizontalS),
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke1),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorNeutralBackground2
  },
  iconDisplay: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap(tokens.spacingHorizontalS),
    fontSize: '20px',
    color: tokens.colorBrandForeground1
  }
});

export interface ISectionEditDialogProps {
  isOpen: boolean;
  section?: IContainerSection;
  canDelete?: boolean;
  onSave: (updatedSection: Partial<IContainerSection>) => void;
  onDelete?: () => void;
  onDismiss: () => void;
}

export const SectionEditDialog: React.FC<ISectionEditDialogProps> = ({
  isOpen,
  section,
  canDelete = false,
  onSave,
  onDelete,
  onDismiss
}) => {
  const styles = useStyles();
  const [formData, setFormData] = React.useState<Partial<IContainerSection>>({});
  const [panelWidth, setPanelWidth] = React.useState<number>(440);
  const [isIconPickerOpen, setIsIconPickerOpen] = React.useState<boolean>(false);

  // Drag-to-resize side panel
  const startResizeDrag = (e: React.MouseEvent): void => {
    e.preventDefault();
    const startX = e.clientX;
    const startWidth = panelWidth;

    const onMouseMove = (moveEv: MouseEvent): void => {
      const deltaX = startX - moveEv.clientX;
      const newWidth = Math.max(340, Math.min(800, startWidth + deltaX));
      setPanelWidth(newWidth);
    };

    const onMouseUp = (): void => {
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', onMouseUp);
    };

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
  };

  React.useEffect(() => {
    if (section) {
      setFormData({
        title: section.title || '',
        description: section.description || '',
        badge: section.badge || '',
        iconName: section.iconName || 'BookAnswers',
        iconColor: section.iconColor || '',
        iconBackgroundColor: section.iconBackgroundColor || '',
        showIconBackground: section.showIconBackground !== false,
        backgroundColor: section.backgroundColor || ''
      });
    }
  }, [section, isOpen]);

  if (!isOpen || !section) {
    return null;
  }

  const handleSave = (): void => {
    onSave(formData);
    onDismiss();
  };

  return (
    <>
      <div className={styles.backdrop} onClick={onDismiss}>
        <div
          className={styles.sidePanel}
          style={{ width: `${panelWidth}px` }}
          onClick={(e) => e.stopPropagation()}
        >
          {/* Left-edge draggable resize bar */}
          <div
            className={styles.leftResizeHandle}
            onMouseDown={startResizeDrag}
            title="Drag to adjust sidebar width"
          />

          {/* Header */}
          <div className={styles.panelHeader}>
            <div>
              <Subtitle2 style={{ fontWeight: 700, color: tokens.colorNeutralForeground1, display: 'block' }}>
                Edit Section Properties
              </Subtitle2>
              <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                Configure section title, icon, badge, and custom theme background.
              </Caption1>
            </div>
            <Button
              appearance="subtle"
              size="small"
              icon={<DismissRegular />}
              onClick={onDismiss}
              title="Close panel"
            />
          </div>

          {/* Body */}
          <div className={styles.panelBody}>
            {/* Title & Badge */}
            <div className={styles.fieldRow}>
              <Label weight="semibold" required>Section Title</Label>
              <Input
                value={formData.title || ''}
                placeholder="e.g. Operations & Delivery"
                onChange={(e, data) => setFormData({ ...formData, title: data.value })}
              />
            </div>

            <div className={styles.fieldRow}>
              <Label weight="semibold">Section Badge / Tag</Label>
              <Input
                value={formData.badge || ''}
                placeholder="e.g. Core, Live, 2026 Q3 (or leave blank for none)"
                onChange={(e, data) => setFormData({ ...formData, badge: data.value })}
              />
            </div>

            <div className={styles.fieldRow}>
              <Label weight="semibold">Description / Guidance</Label>
              <Textarea
                value={formData.description || ''}
                placeholder="Guidance text displayed under section header in accordion view..."
                rows={3}
                onChange={(e, data) => setFormData({ ...formData, description: data.value })}
              />
            </div>

            {/* Icon Picker & Glyph Color */}
            <div className={styles.fieldRow}>
              <Label weight="semibold">Section Icon (Corporate & Fluent)</Label>
              <div className={styles.iconPreviewBox}>
                <div className={styles.iconDisplay}>
                  {renderUnifiedIcon(formData.iconName || 'BookAnswers', formData.iconColor || tokens.colorBrandForeground1, '24px')}
                  <span style={{ fontSize: '0.85rem', fontWeight: 600, color: tokens.colorNeutralForeground1 }}>
                    {formData.iconName || 'BookAnswers'}
                  </span>
                </div>
                <Button
                  appearance="secondary"
                  size="small"
                  icon={<AppsRegular />}
                  onClick={() => setIsIconPickerOpen(true)}
                >
                  Change Icon
                </Button>
              </div>
            </div>

            <div className={styles.twoColRow}>
              {/* Icon Glyph Color */}
              <div className={styles.fieldRow}>
                <Label weight="semibold">Icon Glyph Colour</Label>
                <BrandColorPickerPopover
                  selectedColor={formData.iconColor}
                  onChange={(color?: string) => setFormData({ ...formData, iconColor: color })}
                  defaultLabel="Default (Brand Cyan)"
                  defaultColorHex="#0090DC"
                />
              </div>

              {/* Icon Container Background */}
              <div className={styles.fieldRow}>
                <Label weight="semibold">Icon Box Background</Label>
                <BrandColorPickerPopover
                  selectedColor={formData.iconBackgroundColor}
                  onChange={(color?: string) => setFormData({ ...formData, iconBackgroundColor: color })}
                  defaultLabel="Default (Soft Tint Box)"
                  defaultColorHex="#CCE9F8"
                />
              </div>
            </div>

            <div className={styles.fieldRow}>
              <Checkbox
                checked={formData.showIconBackground !== false}
                label="Include rounded icon background container"
                onChange={(ev, data) => setFormData({ ...formData, showIconBackground: !!data.checked })}
              />
            </div>

            {/* Custom Section Background */}
            <div className={styles.fieldRow}>
              <Label weight="semibold">Section Background Colour</Label>
              <BrandColorPickerPopover
                selectedColor={formData.backgroundColor}
                onChange={(color?: string) => setFormData({ ...formData, backgroundColor: color })}
                defaultLabel="Default (Transparent canvas)"
                defaultColorHex="transparent"
              />
            </div>
          </div>

          {/* Footer */}
          <div className={styles.panelFooter}>
            <div>
              {canDelete && onDelete && (
                <Button
                  appearance="subtle"
                  icon={<DeleteRegular />}
                  style={{ color: tokens.colorPaletteRedForeground1 }}
                  onClick={() => {
                    if (window.confirm(`Are you sure you want to delete section "${section.title}"?`)) {
                      onDelete();
                      onDismiss();
                    }
                  }}
                >
                  Delete Section
                </Button>
              )}
            </div>
            <div style={{ display: 'flex', gap: '8px' }}>
              <Button appearance="secondary" onClick={onDismiss}>
                Cancel
              </Button>
              <Button
                appearance="primary"
                icon={<SaveRegular />}
                onClick={handleSave}
              >
                Save Section
              </Button>
            </div>
          </div>
        </div>
      </div>

      {/* Visual Fluent UI 2 & Corporate SVG Icon Picker */}
      <FluentIconPicker
        isOpen={isIconPickerOpen}
        selectedIconKey={formData.iconName || 'BookAnswers'}
        onSelectIcon={(iconKey) => {
          setFormData({ ...formData, iconName: iconKey });
        }}
        onDismiss={() => setIsIconPickerOpen(false)}
      />
    </>
  );
};
