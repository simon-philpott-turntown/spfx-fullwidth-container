/**
 * @file InsertionBar.tsx
 * @description Hover-activated (+) insertion bar for inserting content items on the canvas (Screenshot 4).
 */

import * as React from 'react';
import { makeStyles, tokens, shorthands } from '@fluentui/react-components';
import { AddRegular } from '@fluentui/react-icons';
import { ToolboxModal } from './ToolboxModal';
import { ICardItemType } from '../models/IContainerModels';

const useStyles = makeStyles({
  container: {
    position: 'relative',
    width: '100%',
    height: '18px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    cursor: 'pointer',
    ...shorthands.margin('4px', '0'),
    opacity: 0,
    transitionProperty: 'opacity',
    transitionDuration: tokens.durationNormal,
    '&:hover': {
      opacity: 1
    }
  },
  containerAlwaysVisible: {
    opacity: 0.7,
    '&:hover': {
      opacity: 1
    }
  },
  line: {
    position: 'absolute',
    left: 0,
    right: 0,
    top: '50%',
    height: '2px',
    backgroundColor: tokens.colorBrandStroke1,
    transform: 'translateY(-50%)',
    zIndex: 1
  },
  addButton: {
    position: 'relative',
    zIndex: 2,
    width: '24px',
    height: '24px',
    ...shorthands.borderRadius('50%'),
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: '14px',
    boxShadow: tokens.shadow4,
    transitionProperty: 'all',
    transitionDuration: tokens.durationFast,
    '&:hover': {
      transform: 'scale(1.15)',
      backgroundColor: tokens.colorBrandBackgroundHover
    }
  }
});

export interface IInsertionBarProps {
  onInsert: (itemType: ICardItemType) => void;
  alwaysVisible?: boolean;
  contextTitle?: string;
}

export const InsertionBar: React.FC<IInsertionBarProps> = ({
  onInsert,
  alwaysVisible = false,
  contextTitle = 'Add content to card'
}) => {
  const styles = useStyles();
  const [isToolboxOpen, setIsToolboxOpen] = React.useState<boolean>(false);

  return (
    <>
      <div
        className={`${styles.container} ${alwaysVisible ? styles.containerAlwaysVisible : ''}`}
        onClick={(e) => {
          e.stopPropagation();
          setIsToolboxOpen(true);
        }}
        title="Click to add content here"
      >
        <div className={styles.line} />
        <div className={styles.addButton}>
          <AddRegular />
        </div>
      </div>

      <ToolboxModal
        isOpen={isToolboxOpen}
        onDismiss={() => setIsToolboxOpen(false)}
        onSelectItem={(itemType) => {
          setIsToolboxOpen(false);
          onInsert(itemType);
        }}
        contextTitle={contextTitle}
      />
    </>
  );
};
