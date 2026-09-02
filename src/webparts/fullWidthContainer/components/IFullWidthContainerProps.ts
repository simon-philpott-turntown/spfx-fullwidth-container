/**
 * @file IFullWidthContainerProps.ts
 * @description Props contract for the FullWidthContainer root React component.
 */

import { IContainerSection, LayoutMode, ContainerStyle } from '../models/IContainerModels';

export interface IFullWidthContainerProps {
  title: string;
  subtitle?: string;
  layoutMode: LayoutMode;
  containerStyle: ContainerStyle;
  accentColor: string;
  enableAnimation: boolean;
  compactPadding: boolean;
  showSearch: boolean;
  gridColumns?: number;
  gridRows?: number;
  sections: IContainerSection[];
  isDarkTheme: boolean;
  userDisplayName: string;
  spfxTheme?: unknown;
  isEditMode?: boolean;
  onOpenPropertyPane?: () => void;
  onTitleChange?: (newTitle: string) => void;
  onSubtitleChange?: (newSubtitle: string) => void;
  onUpdateSection?: (sectionId: string, updatedFields: Partial<IContainerSection>) => void;
  onAddSection?: () => void;
  onDeleteSection?: (sectionId: string) => void;
  onUpdateBlock?: (sectionId: string, blockId: string, updatedFields: Partial<import('../models/IContainerModels').IContentBlock>) => void;
  onAddBlock?: (sectionId: string) => void;
  onDeleteBlock?: (sectionId: string, blockId: string) => void;
  onEditBlockProperties?: (sectionIndex: number, blockIndex: number) => void;
}
