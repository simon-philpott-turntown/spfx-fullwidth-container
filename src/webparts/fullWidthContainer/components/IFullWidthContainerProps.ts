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
  sections: IContainerSection[];
  isDarkTheme: boolean;
  userDisplayName: string;
}
