/**
 * @file ToolboxModal.tsx
 * @description SharePoint-style Content Toolbox modal for inserting composable items into cards (Screenshot 1).
 */

import * as React from 'react';
import {
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogContent,
  makeStyles,
  tokens,
  shorthands,
  Button
} from '@fluentui/react-components';
import {
  DismissRegular,
  TextDescriptionRegular,
  CursorClickRegular,
  MegaphoneRegular,
  LineHorizontal1Regular,
  NewsRegular,
  ImageRegular,
  ImageMultipleRegular,
  LinkRegular,
  LinkSquareRegular,
  VideoRegular,
  DataTrendingRegular,
  TagRegular,
  SparkleRegular
} from '@fluentui/react-icons';
import { ICardItemType } from '../models/IContainerModels';

const useStyles = makeStyles({
  surface: {
    maxWidth: '540px',
    width: '95vw',
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    ...shorthands.padding('20px')
  },
  headerRow: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: '16px'
  },
  sectionTitle: {
    fontSize: '0.85rem',
    fontWeight: 600,
    textTransform: 'uppercase',
    letterSpacing: '0.5px',
    color: tokens.colorNeutralForeground3,
    marginBottom: '12px'
  },
  toolboxGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(4, 1fr)',
    gap: '12px',
    marginBottom: '16px',
    '@media (max-width: 500px)': {
      gridTemplateColumns: 'repeat(3, 1fr)'
    }
  },
  toolItem: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    ...shorthands.padding('12px', '8px'),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    backgroundColor: tokens.colorNeutralBackground1,
    cursor: 'pointer',
    textAlign: 'center',
    transitionProperty: 'all',
    transitionDuration: tokens.durationNormal,
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
      ...shorthands.borderColor(tokens.colorBrandStroke1),
      transform: 'translateY(-2px)',
      boxShadow: tokens.shadow4
    }
  },
  toolIcon: {
    fontSize: '24px',
    color: tokens.colorBrandForeground1,
    marginBottom: '8px'
  },
  toolLabel: {
    fontSize: '0.78rem',
    fontWeight: 500,
    color: tokens.colorNeutralForeground1,
    lineHeight: '1.2'
  }
});

export interface IToolboxModalProps {
  isOpen: boolean;
  onDismiss: () => void;
  onSelectItem: (itemType: ICardItemType) => void;
  contextTitle?: string;
}

export const ToolboxModal: React.FC<IToolboxModalProps> = ({
  isOpen,
  onDismiss,
  onSelectItem,
  contextTitle = 'Add content to card'
}) => {
  const styles = useStyles();

  if (!isOpen) return null;

  const contentTools: Array<{ type: ICardItemType; label: string; icon: React.ReactElement }> = [
    { type: 'text', label: 'Text', icon: <TextDescriptionRegular /> },
    { type: 'button', label: 'Button', icon: <CursorClickRegular /> },
    { type: 'cta', label: 'Call to action', icon: <MegaphoneRegular /> },
    { type: 'divider', label: 'Divider', icon: <LineHorizontal1Regular /> },
    { type: 'editorial', label: 'Editorial card', icon: <NewsRegular /> },
    { type: 'hero', label: 'Hero', icon: <SparkleRegular /> },
    { type: 'image', label: 'Image', icon: <ImageRegular /> },
    { type: 'gallery', label: 'Image gallery', icon: <ImageMultipleRegular /> },
    { type: 'link', label: 'Link', icon: <LinkRegular /> },
    { type: 'quickLinks', label: 'Quick links', icon: <LinkSquareRegular /> },
    { type: 'video', label: 'Video', icon: <VideoRegular /> },
    { type: 'liveData', label: 'Live Data API', icon: <DataTrendingRegular /> },
    { type: 'termStoreTags', label: 'Term Store Tags', icon: <TagRegular /> }
  ];

  return (
    <Dialog open={isOpen} onOpenChange={(_, data) => !data.open && onDismiss()}>
      <DialogSurface className={styles.surface}>
        <div className={styles.headerRow}>
          <DialogTitle>{contextTitle}</DialogTitle>
          <Button
            appearance="subtle"
            icon={<DismissRegular />}
            onClick={onDismiss}
            aria-label="Close"
          />
        </div>

        <DialogBody>
          <DialogContent>
            <div className={styles.sectionTitle}>Text, media, and content</div>
            <div className={styles.toolboxGrid}>
              {contentTools.map((tool) => (
                <div
                  key={tool.type}
                  className={styles.toolItem}
                  onClick={() => {
                    onSelectItem(tool.type);
                    onDismiss();
                  }}
                  role="button"
                  tabIndex={0}
                >
                  <div className={styles.toolIcon}>{tool.icon}</div>
                  <div className={styles.toolLabel}>{tool.label}</div>
                </div>
              ))}
            </div>
          </DialogContent>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};
