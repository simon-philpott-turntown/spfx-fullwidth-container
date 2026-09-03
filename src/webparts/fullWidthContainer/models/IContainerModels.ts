/**
 * @file IContainerModels.ts
 * @description Data models and default presets for the Full-Width Container Web Part.
 */

export type LayoutMode = 'tabs' | 'accordion';
export type ContainerStyle = 'standard' | 'glassmorphism' | 'branded' | 'minimal' | 'gradient';
export type TabStyle = 'pills' | 'underline' | 'cards';

export type BlockType = 'card' | 'metric' | 'embed' | 'richText' | 'quickLinks';

/**
 * Inner composable item types for cards (matching Screenshot 1).
 */
export type ICardItemType =
  | 'text'
  | 'button'
  | 'cta'
  | 'divider'
  | 'editorial'
  | 'hero'
  | 'image'
  | 'gallery'
  | 'link'
  | 'quickLinks'
  | 'video'
  | 'liveData'
  | 'termStoreTags';

/**
 * Term Store Tag reference.
 */
export interface ITermStoreTag {
  id: string;
  label: string;
  path?: string;
  termSetId?: string;
  termSetName?: string;
}

/**
 * Dynamic REST / Graph API Live Data Configuration.
 */
export interface ILiveDataConfig {
  apiUrl: string;
  method?: 'GET' | 'POST';
  jsonPath: string;
  prefix?: string;
  suffix?: string;
  refreshIntervalSeconds?: number;
  fallbackText?: string;
}

/**
 * Nested inner item within a composable card.
 */
export interface ICardItem {
  id: string;
  type: ICardItemType;
  text?: string;
  buttonLabel?: string;
  buttonUrl?: string;
  buttonVariant?: 'primary' | 'subtle' | 'outline';
  ctaHeading?: string;
  ctaDescription?: string;
  ctaButtonText?: string;
  ctaButtonUrl?: string;
  dividerStyle?: 'solid' | 'dashed';
  imageUrl?: string;
  imageAlt?: string;
  imageCaption?: string;
  galleryImages?: Array<{ url: string; caption?: string }>;
  quickLinks?: Array<{ label: string; url: string; icon?: string }>;
  videoUrl?: string;
  liveDataConfig?: ILiveDataConfig;
  termStoreTags?: ITermStoreTag[];
}

export type CardHeightMode = 'auto' | 'equal';

/**
 * Interface representing a child content block inside a container section.
 */
export interface IContentBlock {
  id: string;
  type: BlockType;
  title: string;
  description?: string;
  badge?: string;
  iconName?: string;
  iconColor?: string; // Custom icon foreground color
  iconBackgroundColor?: string; // Custom icon container background color
  showIconBackground?: boolean; // Whether icon has a rounded background container (default: true)
  linkUrl?: string;
  linkText?: string;
  metricValue?: string;
  metricTrend?: string;
  metricTrendPositive?: boolean;
  embedUrl?: string;
  richContent?: string;
  tags?: string[];
  termStoreTags?: ITermStoreTag[];
  liveDataConfig?: ILiveDataConfig;
  colSpan?: number; // 1, 2, 3, 4 columns
  rowSpan?: number; // 1, 2 rows
  heightMode?: 'default' | 'auto' | 'equal'; // Auto (fit content) vs Equal (match tallest on row)
  backgroundColor?: string; // Custom card background color (independent)
  textColor?: string; // Custom card text color override
  items?: ICardItem[]; // Composable nested elements
}

/**
 * Interface representing a container section / tab / accordion panel.
 */
export interface IContainerSection {
  id: string;
  title: string;
  iconName?: string;
  badge?: string;
  description?: string;
  backgroundColor?: string; // Custom section background color (independent)
  textColor?: string; // Custom section text color override
  blocks: IContentBlock[];
}

/**
 * Official Turner & Townsend Brand Colors (Version 2.2 — July 2025)
 */
export const TT_BRAND_COLORS = {
  // Primary Brand Colors
  blue: '#1E4479',
  blueDeep: '#001436', // 100+
  cyan: '#0090DC',
  cyanDeep: '#0073A5', // 100+
  grey: '#505A60',
  greyDeep: '#292929', // 100+

  // Secondary Brand Colors
  green: '#00A000',
  greenDeep: '#007E1E', // 100+
  orange: '#D55C17',
  orangeDeep: '#C83700', // 100+

  // Backgrounds
  mushroom: '#F2EEE7',
  white: '#FFFFFF',

  // Primary Tints
  blueTints: {
    t80: '#4B6994',
    t60: '#788FAE',
    t40: '#A5B4C9',
    t20: '#D2DAE4'
  },
  cyanTints: {
    t80: '#33A6E3',
    t60: '#66BCEB',
    t40: '#99D3F1',
    t20: '#CCE9F8'
  },
  greyTints: {
    t80: '#737B80',
    t60: '#969CA0',
    t40: '#B9BDC0',
    t20: '#DCDFE0'
  },

  // Secondary Tints
  greenTints: {
    t80: '#33B333',
    t60: '#66C666',
    t40: '#99D999',
    t20: '#CCEECC'
  },
  orangeTints: {
    t80: '#DD7D45',
    t60: '#E69D74',
    t40: '#EEBEA2',
    t20: '#F7DED1'
  }
};

/**
 * Default sample sections demonstrating descriptive placeholders, British metrics (£), and sentence case syntax.
 */
export const DEFAULT_CONTAINER_SECTIONS: IContainerSection[] = [
  {
    id: 'sec-1',
    title: 'Section title 1',
    iconName: 'BookAnswers',
    badge: 'Badge',
    description: 'Section summary 1',
    blocks: [
      {
        id: 'blk-1',
        type: 'card',
        title: 'Card title 1',
        description: 'Card summary 1',
        badge: 'Badge',
        iconName: 'BookAnswers',
        linkUrl: '#',
        linkText: 'Action link',
        tags: ['Tag 1', 'Tag 2']
      },
      {
        id: 'blk-2',
        type: 'card',
        title: 'Card title 2',
        description: 'Card summary 2',
        badge: 'Badge',
        iconName: 'CheckList',
        linkUrl: '#',
        linkText: 'Action link',
        tags: ['Tag 3', 'Tag 4']
      },
      {
        id: 'blk-3',
        type: 'card',
        title: 'Card title 3',
        description: 'Card summary 3',
        badge: 'Badge',
        iconName: 'Calculator',
        linkUrl: '#',
        linkText: 'Action link',
        tags: ['Tag 5']
      },
      {
        id: 'blk-4',
        type: 'card',
        title: 'Card title 4',
        description: 'Card summary 4',
        badge: 'Badge',
        iconName: 'ComplianceAudit',
        linkUrl: '#',
        linkText: 'Action link',
        tags: ['Tag 6']
      }
    ]
  },
  {
    id: 'sec-2',
    title: 'Section title 2',
    iconName: 'Financial',
    badge: 'Real-time',
    description: 'Section summary 2',
    blocks: [
      {
        id: 'blk-m1',
        type: 'metric',
        title: 'Metric title 1',
        metricValue: '£1,420,000',
        metricTrend: '+6.2% vs baseline',
        metricTrendPositive: true,
        description: 'Metric summary 1',
        iconName: 'Financial'
      },
      {
        id: 'blk-m2',
        type: 'metric',
        title: 'Metric title 2',
        metricValue: '£3,420,000',
        metricTrend: '+18.4% ahead of target',
        metricTrendPositive: true,
        description: 'Metric summary 2',
        iconName: 'Financial'
      },
      {
        id: 'blk-m3',
        type: 'metric',
        title: 'Metric title 3',
        metricValue: '94.8%',
        metricTrend: '+1.5% from last period',
        metricTrendPositive: true,
        description: 'Metric summary 3',
        iconName: 'TimelineProgress'
      },
      {
        id: 'blk-m4',
        type: 'metric',
        title: 'Metric title 4',
        metricValue: '28 / 30',
        metricTrend: '2 pending verification',
        metricTrendPositive: true,
        description: 'Metric summary 4',
        iconName: 'CheckList'
      }
    ]
  },
  {
    id: 'sec-3',
    title: 'Section title 3',
    iconName: 'AppIconDefault',
    badge: 'Interactive',
    description: 'Section summary 3',
    blocks: [
      {
        id: 'blk-e1',
        type: 'embed',
        title: 'Tool title 1',
        description: 'Tool summary 1',
        embedUrl: 'https://en.wikipedia.org/wiki/SharePoint',
        linkUrl: '#',
        linkText: 'Open link'
      },
      {
        id: 'blk-e2',
        type: 'card',
        title: 'Card title 5',
        description: 'Card summary 5',
        badge: 'Badge',
        iconName: 'Calculator',
        linkUrl: '#',
        linkText: 'Action link',
        tags: ['Tag 7', 'Tag 8']
      }
    ]
  }
];

export const PRESET_TEMPLATES: Record<string, { name: string; sections: IContainerSection[] }> = {
  commercial: {
    name: 'Commercial and financial hub (GBP £)',
    sections: DEFAULT_CONTAINER_SECTIONS
  },
  governance: {
    name: 'Assurance and governance portal',
    sections: [
      {
        id: 'sec-gov-1',
        title: 'Section title 1',
        iconName: 'ComplianceAudit',
        badge: 'Core',
        description: 'Section summary 1',
        blocks: [
          {
            id: 'blk-g-1',
            type: 'card',
            title: 'Card title 1',
            description: 'Card summary 1',
            badge: 'Badge',
            iconName: 'Lock',
            linkText: 'Action link',
            linkUrl: '#',
            tags: ['Tag 1', 'Tag 2']
          },
          {
            id: 'blk-g-2',
            type: 'metric',
            title: 'Metric title 1',
            metricValue: '99.4%',
            metricTrend: '+0.8% YoY',
            metricTrendPositive: true,
            description: 'Metric summary 1'
          }
        ]
      },
      {
        id: 'sec-gov-2',
        title: 'Section title 2',
        iconName: 'TimelineProgress',
        badge: 'Real-time',
        description: 'Section summary 2',
        blocks: [
          {
            id: 'blk-r-1',
            type: 'card',
            title: 'Card title 2',
            description: 'Card summary 2',
            badge: 'Badge',
            linkText: 'Action link',
            linkUrl: '#',
            tags: ['Tag 3', 'Tag 4']
          }
        ]
      }
    ]
  },
  starter: {
    name: 'Blank 2-section starter',
    sections: [
      {
        id: 'sec-start-1',
        title: 'Section title 1',
        iconName: 'BookAnswers',
        badge: 'Badge',
        description: 'Section summary 1',
        blocks: [
          {
            id: 'blk-s-1',
            type: 'card',
            title: 'Card title 1',
            description: 'Card summary 1',
            badge: 'Badge',
            linkText: 'Action link',
            linkUrl: '#',
            tags: ['Tag 1', 'Tag 2']
          }
        ]
      },
      {
        id: 'sec-start-2',
        title: 'Section title 2',
        iconName: 'Financial',
        badge: 'Badge',
        description: 'Section summary 2',
        blocks: [
          {
            id: 'blk-s-2',
            type: 'metric',
            title: 'Metric title 1',
            metricValue: '£100,000',
            metricTrend: '+10%',
            metricTrendPositive: true,
            description: 'Metric summary 1'
          }
        ]
      }
    ]
  }
};
