/**
 * @file CustomSvgIconRegistry.tsx
 * @description Extensible SVG Icon Registry and Unified Icon Renderer for the Full-Width Dashboard.
 * Allows custom corporate SVG icon collections to be seamlessly registered alongside SPFx Fluent UI 2 icons,
 * with dynamic 'currentColor' inheritance, custom icon background styling, and visual color picking.
 */

import * as React from 'react';
import {
  DocumentRegular,
  FolderRegular,
  MoneyRegular,
  ReceiptMoneyRegular,
  ArrowTrendingLinesRegular,
  ChartMultipleRegular,
  ShieldCheckmarkRegular,
  CheckmarkCircleRegular,
  LockClosedRegular,
  GlobeRegular,
  WrenchRegular,
  PeopleRegular,
  BuildingRegular,
  MegaphoneRegular,
  SparkleRegular,
  StarRegular,
  AppsRegular,
  BookOpenRegular,
  TagRegular,
  SearchRegular,
  CalendarLtrRegular,
  LightbulbRegular,
  TargetRegular,
  TrophyRegular,
  GaugeRegular,
  LinkRegular,
  MailRegular,
  PhoneRegular,
  TimerRegular,
  WarningRegular,
  InfoRegular,
  KeyRegular,
  BookmarkRegular,
  HeartRegular,
  ShareRegular
} from '@fluentui/react-icons';

export interface ICustomSvgIcon {
  id: string;
  name: string;
  category: 'Custom Brand SVGs' | 'Fluent UI 2';
  keywords: string[];
  svgPathOrNode: React.ReactNode;
}

/**
 * Registry of custom SVG icons.
 * Any new collection of SVG icons provided can be directly added or registered here.
 */
export const CUSTOM_SVG_ICONS: ICustomSvgIcon[] = [
  {
    id: 'svg-brand-book',
    name: 'Brand Playbook (Screenshot)',
    category: 'Custom Brand SVGs',
    keywords: ['book', 'playbook', 'manual', 'pages', 'guidelines'],
    svgPathOrNode: (
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        <rect x="3.5" y="4.5" width="7" height="15" rx="1.5" stroke="currentColor" strokeWidth="1.8" />
        <rect x="13.5" y="4.5" width="7" height="15" rx="1.5" stroke="currentColor" strokeWidth="1.8" />
      </svg>
    )
  },
  {
    id: 'svg-brand-document',
    name: 'Brand Project Document',
    category: 'Custom Brand SVGs',
    keywords: ['doc', 'file', 'contract', 'statement'],
    svgPathOrNode: (
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M6 3.5H14L18.5 8V20.5H6V3.5Z" stroke="currentColor" strokeWidth="1.8" strokeLinejoin="round" />
        <path d="M14 3.5V8H18.5" stroke="currentColor" strokeWidth="1.8" strokeLinejoin="round" />
        <path d="M9 12H15.5M9 15.5H15.5" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
      </svg>
    )
  },
  {
    id: 'svg-brand-analytics',
    name: 'Brand Analytics Chart',
    category: 'Custom Brand SVGs',
    keywords: ['chart', 'analytics', 'growth', 'trending', 'kpi'],
    svgPathOrNode: (
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M4 20.5H20" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
        <path d="M6.5 16.5L10.5 11.5L14 14.5L18.5 7.5" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
        <circle cx="18.5" cy="7.5" r="1.5" fill="currentColor" />
      </svg>
    )
  },
  {
    id: 'svg-brand-shield',
    name: 'Brand Assurance Shield',
    category: 'Custom Brand SVGs',
    keywords: ['shield', 'compliance', 'security', 'audit', 'assurance'],
    svgPathOrNode: (
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M12 3.5L4.5 7V13C4.5 17.5 12 21 12 21C12 21 19.5 17.5 19.5 13V7L12 3.5Z" stroke="currentColor" strokeWidth="1.8" strokeLinejoin="round" />
        <path d="M9.5 12L11 13.5L15 9.5" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
      </svg>
    )
  },
  {
    id: 'svg-brand-coins',
    name: 'Brand Financial Currency (£)',
    category: 'Custom Brand SVGs',
    keywords: ['finance', 'pound', 'money', 'commercial', 'cost'],
    svgPathOrNode: (
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        <circle cx="12" cy="12" r="8.5" stroke="currentColor" strokeWidth="1.8" />
        <path d="M13.5 8.5C12.5 8.5 11.5 9 11 10C10.5 11 10.5 12 10.5 13.5H14.5M9.5 12H13.5" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
        <path d="M9.5 15.5H14.5" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
      </svg>
    )
  }
];

/**
 * Registers custom SVG icons into the runtime registry.
 */
export function registerCustomSvgIcons(newIcons: ICustomSvgIcon[]): void {
  newIcons.forEach((icon) => {
    const existingIdx = CUSTOM_SVG_ICONS.findIndex((i) => i.id === icon.id);
    if (existingIdx >= 0) {
      CUSTOM_SVG_ICONS[existingIdx] = icon;
    } else {
      CUSTOM_SVG_ICONS.push(icon);
    }
  });
}

/**
 * Unified Icon Renderer:
 * Renders either a custom SVG icon or a standard Fluent UI 2 icon with custom foreground color and size.
 */
export function renderUnifiedIcon(
  iconKey?: string,
  iconColor?: string,
  fontSize: string = '20px'
): React.ReactElement {
  const customIcon = CUSTOM_SVG_ICONS.find((i) => i.id === iconKey || i.name === iconKey);
  if (customIcon) {
    return (
      <span
        style={{
          display: 'inline-flex',
          alignItems: 'center',
          justifyContent: 'center',
          color: iconColor || 'currentColor',
          width: fontSize,
          height: fontSize,
          lineHeight: 1
        }}
      >
        {customIcon.svgPathOrNode}
      </span>
    );
  }

  // Fluent UI 2 Icons
  const iconStyle: React.CSSProperties = {
    color: iconColor || 'currentColor',
    fontSize
  };

  switch (iconKey) {
    case 'Document': return <DocumentRegular style={iconStyle} />;
    case 'Folder': return <FolderRegular style={iconStyle} />;
    case 'Financial': return <MoneyRegular style={iconStyle} />;
    case 'ReceiptMoney': return <ReceiptMoneyRegular style={iconStyle} />;
    case 'TimelineProgress': return <ArrowTrendingLinesRegular style={iconStyle} />;
    case 'ChartMultiple': return <ChartMultipleRegular style={iconStyle} />;
    case 'ComplianceAudit': return <ShieldCheckmarkRegular style={iconStyle} />;
    case 'CheckList': return <CheckmarkCircleRegular style={iconStyle} />;
    case 'Lock': return <LockClosedRegular style={iconStyle} />;
    case 'Globe': return <GlobeRegular style={iconStyle} />;
    case 'Wrench': return <WrenchRegular style={iconStyle} />;
    case 'People': return <PeopleRegular style={iconStyle} />;
    case 'Building': return <BuildingRegular style={iconStyle} />;
    case 'Megaphone': return <MegaphoneRegular style={iconStyle} />;
    case 'Sparkle': return <SparkleRegular style={iconStyle} />;
    case 'Star': return <StarRegular style={iconStyle} />;
    case 'AppIconDefault': return <AppsRegular style={iconStyle} />;
    case 'Tag': return <TagRegular style={iconStyle} />;
    case 'Search': return <SearchRegular style={iconStyle} />;
    case 'Calendar': return <CalendarLtrRegular style={iconStyle} />;
    case 'Lightbulb': return <LightbulbRegular style={iconStyle} />;
    case 'Target': return <TargetRegular style={iconStyle} />;
    case 'Trophy': return <TrophyRegular style={iconStyle} />;
    case 'Gauge': return <GaugeRegular style={iconStyle} />;
    case 'Link': return <LinkRegular style={iconStyle} />;
    case 'Mail': return <MailRegular style={iconStyle} />;
    case 'Phone': return <PhoneRegular style={iconStyle} />;
    case 'Timer': return <TimerRegular style={iconStyle} />;
    case 'Warning': return <WarningRegular style={iconStyle} />;
    case 'Info': return <InfoRegular style={iconStyle} />;
    case 'Key': return <KeyRegular style={iconStyle} />;
    case 'Bookmark': return <BookmarkRegular style={iconStyle} />;
    case 'Heart': return <HeartRegular style={iconStyle} />;
    case 'Share': return <ShareRegular style={iconStyle} />;
    case 'BookAnswers':
    default:
      return <BookOpenRegular style={iconStyle} />;
  }
}
