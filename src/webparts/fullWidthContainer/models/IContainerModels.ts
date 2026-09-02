/**
 * @file IContainerModels.ts
 * @description Data models and default presets for the Full-Width Container Web Part.
 */

export type LayoutMode = 'tabs' | 'accordion';
export type ContainerStyle = 'standard' | 'glassmorphism' | 'branded' | 'minimal' | 'gradient';
export type TabStyle = 'pills' | 'underline' | 'cards';

export type BlockType = 'card' | 'metric' | 'embed' | 'richText' | 'quickLinks';

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
  linkUrl?: string;
  linkText?: string;
  metricValue?: string;
  metricTrend?: string;
  metricTrendPositive?: boolean;
  embedUrl?: string;
  richContent?: string;
  tags?: string[];
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
  blocks: IContentBlock[];
}

/**
 * Default sample sections demonstrating Playbooks, Assurance, Metrics (£), and Embedded Tools.
 */
export const DEFAULT_CONTAINER_SECTIONS: IContainerSection[] = [
  {
    id: 'sec-playbooks',
    title: 'Playbooks & Guidance',
    iconName: 'BookAnswers',
    badge: '4 Guides',
    description: 'Centralised operational methodologies, service line playbooks, and best practice frameworks.',
    blocks: [
      {
        id: 'blk-1',
        type: 'card',
        title: 'Commercial Advisory Playbook',
        description: 'Standardized delivery lifecycle, contract governance, and commercial risk matrices.',
        badge: 'Core Framework',
        iconName: 'DocumentManagement',
        linkUrl: '#',
        linkText: 'Open Playbook',
        tags: ['Commercial', 'Advisory', 'P0']
      },
      {
        id: 'blk-2',
        type: 'card',
        title: 'Project Assurance Checklist',
        description: 'Mandatory gateway reviews, milestone verification, and stakeholder sign-off criteria.',
        badge: 'Compliance',
        iconName: 'CheckList',
        linkUrl: '#',
        linkText: 'View Checklist',
        tags: ['Assurance', 'Quality']
      },
      {
        id: 'blk-3',
        type: 'card',
        title: 'Digital Engineering Standards',
        description: 'BIM protocols, ISO 19650 workflows, and common data environment structuring.',
        badge: 'Technical',
        iconName: 'EngineeringGroup',
        linkUrl: '#',
        linkText: 'Read Standards',
        tags: ['BIM', 'ISO19650']
      },
      {
        id: 'blk-4',
        type: 'card',
        title: 'Health & Safety Governance',
        description: 'Zero-harm policy compliance, incident reporting, and on-site audit guidance.',
        badge: 'Mandatory',
        iconName: 'Shield',
        linkUrl: '#',
        linkText: 'Review Policy',
        tags: ['Safety', 'HSE']
      }
    ]
  },
  {
    id: 'sec-metrics',
    title: 'Programme Metrics (£)',
    iconName: 'Financial',
    badge: 'Real-time',
    description: 'Executive overview of programme financials, budget burn-down, and delivery milestones.',
    blocks: [
      {
        id: 'blk-m1',
        type: 'metric',
        title: 'Total Portfolio Budget',
        metricValue: '£42,850,000',
        metricTrend: '+3.2% vs Q2 forecast',
        metricTrendPositive: true,
        description: 'Approved capital allocation across active commissions',
        iconName: 'Money'
      },
      {
        id: 'blk-m2',
        type: 'metric',
        title: 'Efficiency Savings Realised',
        metricValue: '£3,420,000',
        metricTrend: '+18.4% ahead of target',
        metricTrendPositive: true,
        description: 'Value engineering and procurement optimization savings',
        iconName: 'Savings'
      },
      {
        id: 'blk-m3',
        type: 'metric',
        title: 'Commissions on Schedule',
        metricValue: '94.8%',
        metricTrend: '+1.5% from last period',
        metricTrendPositive: true,
        description: 'Critical path milestone tracking across 48 projects',
        iconName: 'TimelineProgress'
      },
      {
        id: 'blk-m4',
        type: 'metric',
        title: 'Active Assurance Audits',
        metricValue: '28 / 30',
        metricTrend: '2 pending gateway 4',
        metricTrendPositive: true,
        description: 'Stage gate compliance rating across regional sectors',
        iconName: 'VerifiedBrand'
      }
    ]
  },
  {
    id: 'sec-embeds',
    title: 'Live Tools & Portals',
    iconName: 'AppIconDefault',
    badge: 'Interactive',
    description: 'Integrated live data feeds, external calculators, and interactive dashboards.',
    blocks: [
      {
        id: 'blk-e1',
        type: 'embed',
        title: 'Power BI Portfolio Performance Dashboard',
        description: 'Live interactive reporting stream directly synchronised with programme datasets.',
        embedUrl: 'https://en.wikipedia.org/wiki/SharePoint',
        linkUrl: '#',
        linkText: 'Launch Fullscreen'
      },
      {
        id: 'blk-e2',
        type: 'card',
        title: 'Cost Estimation Calculator',
        description: 'Real-time NRM1/NRM2 benchmarking model with automated inflation indices.',
        badge: 'Tool',
        iconName: 'Calculator',
        linkUrl: '#',
        linkText: 'Open Calculator',
        tags: ['Cost', 'NRM', 'Live']
      }
    ]
  },
  {
    id: 'sec-governance',
    title: 'Standards & Directory',
    iconName: 'ComplianceAudit',
    badge: 'Verified',
    description: 'Quick links to enterprise policies, client portal logins, and resource directories.',
    blocks: [
      {
        id: 'blk-g1',
        type: 'card',
        title: 'Global Methodology Handbook',
        description: 'Detailed operating models, client engagement principles, and quality standards.',
        badge: 'Edition 2026',
        iconName: 'News',
        linkUrl: '#',
        linkText: 'Download PDF',
        tags: ['Governance', 'Handbook']
      },
      {
        id: 'blk-g2',
        type: 'card',
        title: 'Client Portal & Secure Vault',
        description: 'Encrypted document repository for deliverables, certificates, and commercial registers.',
        badge: 'Secure',
        iconName: 'Lock',
        linkUrl: '#',
        linkText: 'Access Vault',
        tags: ['Security', 'External']
      }
    ]
  }
];

export const PRESET_TEMPLATES: Record<string, { name: string; sections: IContainerSection[] }> = {
  commercial: {
    name: 'Commercial & Financial Hub (Default - GBP)',
    sections: DEFAULT_CONTAINER_SECTIONS
  },
  governance: {
    name: 'Assurance & Governance Portal',
    sections: [
      {
        id: 'sec-gov-1',
        title: 'Policy Frameworks',
        iconName: 'Shield',
        badge: 'P0 Core',
        description: 'Mandatory statutory requirements and enterprise standards.',
        blocks: [
          {
            id: 'blk-g-1',
            type: 'card',
            title: 'ISO 27001 Information Security',
            description: 'Data classification guidelines, access governance, and incident protocols.',
            badge: 'Mandatory',
            iconName: 'Lock',
            linkText: 'Open Policy',
            linkUrl: '#',
            tags: ['Security', 'ISO']
          },
          {
            id: 'blk-g-2',
            type: 'metric',
            title: 'Audit Compliance Score',
            metricValue: '99.4%',
            metricTrend: '+0.8% YoY',
            metricTrendPositive: true,
            description: 'Year-to-date compliance rating across all live business units'
          }
        ]
      },
      {
        id: 'sec-gov-2',
        title: 'Risk Registers & Audits',
        iconName: 'ComplianceAudit',
        badge: 'Real-time',
        description: 'Live mitigation tracking and scheduled assurance reviews.',
        blocks: [
          {
            id: 'blk-r-1',
            type: 'card',
            title: 'Q1 Enterprise Risk Matrix',
            description: 'Active commercial, operational, and supply chain risk registers.',
            badge: 'Q1 2026',
            linkText: 'View Matrix',
            linkUrl: '#',
            tags: ['Risk', 'Governance']
          }
        ]
      }
    ]
  },
  starter: {
    name: 'Blank 2-Section Starter',
    sections: [
      {
        id: 'sec-start-1',
        title: 'Section title',
        iconName: 'BookAnswers',
        badge: 'Badge',
        description: 'Section description summary',
        blocks: [
          {
            id: 'blk-s-1',
            type: 'card',
            title: 'Card title',
            description: 'Card description summary',
            badge: 'Badge',
            linkText: 'Learn more',
            linkUrl: '#',
            tags: ['Tag 1', 'Tag 2']
          }
        ]
      },
      {
        id: 'sec-start-2',
        title: 'Metric section title',
        iconName: 'Financial',
        badge: 'Badge',
        description: 'Metric section description summary',
        blocks: [
          {
            id: 'blk-s-2',
            type: 'metric',
            title: 'Metric title',
            metricValue: '£100,000',
            metricTrend: '+10%',
            metricTrendPositive: true,
            description: 'Metric description summary'
          }
        ]
      }
    ]
  }
};
