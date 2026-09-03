/**
 * @file FluentIconPicker.tsx
 * @description Visual Fluent UI 2 Icon Picker modal with instant keyword search and categories.
 * Follows the Microsoft Fluent 2 Design System.
 */

import * as React from 'react';
import {
  makeStyles,
  shorthands,
  tokens,
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogContent,
  DialogActions,
  Button,
  Input,
  Badge,
  Caption1,
  Subtitle2
} from '@fluentui/react-components';
import {
  SearchRegular,
  DismissRegular,
  CheckmarkRegular,
  DocumentRegular,
  FolderRegular,
  MoneyRegular,
  ArrowTrendingLinesRegular,
  ShieldCheckmarkRegular,
  GlobeRegular,
  AppsRegular,
  LockClosedRegular,
  CheckmarkCircleRegular,
  WrenchRegular,
  MegaphoneRegular,
  BookOpenRegular,
  TagRegular,
  PeopleRegular,
  CalendarLtrRegular,
  ChartMultipleRegular,
  BuildingRegular,
  LightbulbRegular,
  SparkleRegular,
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
  StarRegular,
  HeartRegular,
  ShareRegular,
  FilterRegular,
  CalculatorRegular,
  ReceiptMoneyRegular,
  DesktopRegular,
  CloudRegular,
  RocketRegular
} from '@fluentui/react-icons';

const useStyles = makeStyles({
  dialogSurface: {
    maxWidth: '680px',
    width: '92vw',
    maxHeight: '85vh',
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: tokens.shadow28
  },
  headerRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: tokens.spacingVerticalS
  },
  searchBar: {
    width: '100%',
    marginBottom: tokens.spacingVerticalM
  },
  categoryTabs: {
    display: 'flex',
    flexWrap: 'wrap',
    ...shorthands.gap(tokens.spacingHorizontalXS),
    marginBottom: tokens.spacingVerticalM
  },
  categoryPill: {
    cursor: 'pointer',
    fontSize: '0.8rem',
    ...shorthands.padding('4px', '10px'),
    ...shorthands.borderRadius(tokens.borderRadiusCircular),
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke1),
    backgroundColor: tokens.colorNeutralBackground2,
    color: tokens.colorNeutralForeground2,
    transitionProperty: 'background-color, color, border-color',
    transitionDuration: '150ms',
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
      ...shorthands.borderColor(tokens.colorBrandStroke1)
    }
  },
  categoryPillActive: {
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
    ...shorthands.borderColor(tokens.colorBrandStroke1),
    fontWeight: tokens.fontWeightSemibold
  },
  iconGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fill, minmax(100px, 1fr))',
    ...shorthands.gap(tokens.spacingHorizontalS),
    maxHeight: '360px',
    overflowY: 'auto',
    ...shorthands.padding(tokens.spacingVerticalXS)
  },
  iconItem: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalXS),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    backgroundColor: tokens.colorNeutralBackground1,
    cursor: 'pointer',
    textAlign: 'center',
    transitionProperty: 'transform, border-color, background-color, box-shadow',
    transitionDuration: '150ms',
    ':hover': {
      transform: 'translateY(-2px)',
      ...shorthands.borderColor(tokens.colorBrandStroke1),
      backgroundColor: tokens.colorNeutralBackground1Hover,
      boxShadow: tokens.shadow4
    }
  },
  iconItemActive: {
    backgroundColor: tokens.colorBrandBackground2,
    ...shorthands.borderColor(tokens.colorBrandStroke1),
    boxShadow: tokens.shadow4
  },
  iconGlyph: {
    fontSize: '24px',
    color: tokens.colorBrandForeground1,
    marginBottom: tokens.spacingVerticalXXS,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center'
  },
  iconLabel: {
    fontSize: '0.75rem',
    color: tokens.colorNeutralForeground2,
    lineHeight: '1rem',
    wordBreak: 'break-word'
  },
  previewFooter: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    ...shorthands.padding(tokens.spacingVerticalM, tokens.spacingHorizontalNone),
    ...shorthands.borderTop('1px', 'solid', tokens.colorNeutralStroke2),
    marginTop: tokens.spacingVerticalS
  },
  selectedPreviewBox: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap(tokens.spacingHorizontalS)
  }
});

import { CUSTOM_SVG_ICONS, renderUnifiedIcon } from './CustomSvgIconRegistry';

export interface IFluentIconDefinition {
  key: string;
  label: string;
  category: string;
  keywords: string[];
  icon: React.ReactElement;
}

export const FLUENT_ICONS_LIST: IFluentIconDefinition[] = [
  // Custom Brand Corporate SVGs + 224 TT icons — renderUnifiedIcon handles both svgPathOrNode and svgMarkup
  ...CUSTOM_SVG_ICONS.map((svg) => ({
    key: svg.id,
    label: svg.name,
    category: svg.category,
    keywords: svg.keywords,
    icon: renderUnifiedIcon(svg.id, undefined, '24px')
  })),

  // Common & Content
  { key: 'BookAnswers', label: 'Documentation', category: 'common', keywords: ['book', 'docs', 'manual', 'policy'], icon: <BookOpenRegular /> },
  { key: 'Document', label: 'Document', category: 'common', keywords: ['file', 'doc', 'text', 'paper'], icon: <DocumentRegular /> },
  { key: 'Folder', label: 'Folder', category: 'common', keywords: ['directory', 'archive', 'files'], icon: <FolderRegular /> },
  { key: 'Star', label: 'Star', category: 'common', keywords: ['favorite', 'rating', 'star', 'feature'], icon: <StarRegular /> },
  { key: 'Bookmark', label: 'Bookmark', category: 'common', keywords: ['save', 'ribbon', 'pin'], icon: <BookmarkRegular /> },
  { key: 'Tag', label: 'Tag', category: 'common', keywords: ['metadata', 'label', 'term'], icon: <TagRegular /> },
  { key: 'Heart', label: 'Heart', category: 'common', keywords: ['like', 'health', 'wellness'], icon: <HeartRegular /> },
  { key: 'Lightbulb', label: 'Idea / Insight', category: 'common', keywords: ['bulb', 'innovation', 'idea', 'tip'], icon: <LightbulbRegular /> },
  { key: 'Sparkle', label: 'AI / Smart', category: 'common', keywords: ['ai', 'copilot', 'sparkle', 'magic', 'assist'], icon: <SparkleRegular /> },
  { key: 'Rocket', label: 'Launch / Project', category: 'common', keywords: ['rocket', 'launch', 'fast', 'speed'], icon: <RocketRegular /> },

  // Finance & Metrics
  { key: 'Financial', label: 'Pound (£) Finance', category: 'finance', keywords: ['money', 'finance', 'pound', 'gbp', 'commercial', 'cost'], icon: <MoneyRegular /> },
  { key: 'ReceiptMoney', label: 'Budget & Billing', category: 'finance', keywords: ['receipt', 'invoice', 'payment', 'cash'], icon: <ReceiptMoneyRegular /> },
  { key: 'Calculator', label: 'Calculator', category: 'finance', keywords: ['math', 'accounting', 'calculator', 'formula'], icon: <CalculatorRegular /> },
  { key: 'TimelineProgress', label: 'Trending Metrics', category: 'analytics', keywords: ['trend', 'chart', 'growth', 'kpi', 'progress'], icon: <ArrowTrendingLinesRegular /> },
  { key: 'ChartMultiple', label: 'Analytics Chart', category: 'analytics', keywords: ['report', 'graph', 'data', 'bar', 'dashboard'], icon: <ChartMultipleRegular /> },
  { key: 'Gauge', label: 'Performance Gauge', category: 'analytics', keywords: ['speed', 'gauge', 'meter', 'kpi'], icon: <GaugeRegular /> },
  { key: 'Target', label: 'Target / Goal', category: 'analytics', keywords: ['goal', 'objective', 'bullseye', 'okr'], icon: <TargetRegular /> },
  { key: 'Trophy', label: 'Milestone / Award', category: 'analytics', keywords: ['trophy', 'win', 'achievement', 'success'], icon: <TrophyRegular /> },

  // Security & Governance
  { key: 'ComplianceAudit', label: 'Shield Governance', category: 'governance', keywords: ['shield', 'security', 'compliance', 'audit', 'protection'], icon: <ShieldCheckmarkRegular /> },
  { key: 'CheckList', label: 'Verified Checklist', category: 'governance', keywords: ['check', 'done', 'checklist', 'approved', 'pass'], icon: <CheckmarkCircleRegular /> },
  { key: 'Lock', label: 'Lock / Privacy', category: 'governance', keywords: ['lock', 'security', 'confidential', 'restricted'], icon: <LockClosedRegular /> },
  { key: 'Key', label: 'Access Key', category: 'governance', keywords: ['key', 'permissions', 'auth', 'pass'], icon: <KeyRegular /> },
  { key: 'Warning', label: 'Risk / Warning', category: 'governance', keywords: ['alert', 'warning', 'risk', 'hazard'], icon: <WarningRegular /> },
  { key: 'Info', label: 'Information', category: 'governance', keywords: ['info', 'help', 'details', 'notice'], icon: <InfoRegular /> },

  // Tools & Navigation
  { key: 'AppIconDefault', label: 'Apps & Tools', category: 'tools', keywords: ['apps', 'grid', 'portal', 'suite'], icon: <AppsRegular /> },
  { key: 'Globe', label: 'Globe / Intranet', category: 'tools', keywords: ['web', 'portal', 'globe', 'worldwide', 'internet'], icon: <GlobeRegular /> },
  { key: 'Wrench', label: 'Tools & Engineering', category: 'tools', keywords: ['wrench', 'settings', 'fix', 'build', 'engineering'], icon: <WrenchRegular /> },
  { key: 'Desktop', label: 'Workstation / App', category: 'tools', keywords: ['screen', 'pc', 'desktop', 'monitor'], icon: <DesktopRegular /> },
  { key: 'Cloud', label: 'Cloud Services', category: 'tools', keywords: ['cloud', 'azure', 'server', 'online'], icon: <CloudRegular /> },
  { key: 'Link', label: 'Quick Link', category: 'tools', keywords: ['link', 'url', 'hyperlink', 'chain'], icon: <LinkRegular /> },
  { key: 'Filter', label: 'Filter / Search', category: 'tools', keywords: ['filter', 'query', 'sort'], icon: <FilterRegular /> },
  { key: 'Timer', label: 'Schedule / SLA', category: 'tools', keywords: ['clock', 'time', 'timer', 'sla', 'deadline'], icon: <TimerRegular /> },

  // Communication & People
  { key: 'People', label: 'Team / Organisation', category: 'communication', keywords: ['team', 'people', 'users', 'staff', 'hr'], icon: <PeopleRegular /> },
  { key: 'Building', label: 'Enterprise / Sector', category: 'communication', keywords: ['office', 'building', 'company', 'premises'], icon: <BuildingRegular /> },
  { key: 'Megaphone', label: 'Announcements', category: 'communication', keywords: ['broadcast', 'megaphone', 'news', 'shoutout'], icon: <MegaphoneRegular /> },
  { key: 'Mail', label: 'Email / Comms', category: 'communication', keywords: ['mail', 'message', 'inbox', 'contact'], icon: <MailRegular /> },
  { key: 'Phone', label: 'Contact / Support', category: 'communication', keywords: ['call', 'phone', 'support', 'helpdesk'], icon: <PhoneRegular /> },
  { key: 'Calendar', label: 'Calendar / Events', category: 'communication', keywords: ['calendar', 'date', 'schedule', 'meeting'], icon: <CalendarLtrRegular /> },
  { key: 'Share', label: 'Collaboration', category: 'communication', keywords: ['share', 'network', 'send'], icon: <ShareRegular /> }
];

export interface IFluentIconPickerProps {
  isOpen: boolean;
  selectedIconKey: string;
  onSelectIcon: (iconKey: string) => void;
  onDismiss: () => void;
}

export const FluentIconPicker: React.FC<IFluentIconPickerProps> = ({
  isOpen,
  selectedIconKey,
  onSelectIcon,
  onDismiss
}) => {
  const styles = useStyles();
  const [searchTerm, setSearchTerm] = React.useState<string>('');
  const [activeCategory, setActiveCategory] = React.useState<string>('all');
  const [tempSelectedKey, setTempSelectedKey] = React.useState<string>(selectedIconKey || 'BookAnswers');

  React.useEffect(() => {
    setTempSelectedKey(selectedIconKey || 'BookAnswers');
    setSearchTerm('');
    setActiveCategory('all');
  }, [isOpen, selectedIconKey]);

  if (!isOpen) return null;

  const categories = [
    { key: 'all', label: 'All Icons' },
    { key: 'Custom Brand SVGs', label: '✨ Brand SVGs' },
    { key: 'TT Built Environment', label: '🏗 Built Environment' },
    { key: 'TT Sustainability & Energy', label: '🌿 Sustainability' },
    { key: 'TT Finance & Data', label: '💷 Finance & Data' },
    { key: 'TT People & Teams', label: '🤝 People & Teams' },
    { key: 'TT Technology & Innovation', label: '💡 Technology' },
    { key: 'TT Awards & Learning', label: '🏆 Awards & Learning' },
    { key: 'TT Locations & Travel', label: '🌍 Locations' },
    { key: 'TT Construction & Tools', label: '🔧 Construction' },
    { key: 'TT Operations', label: '📋 Operations' },
    { key: 'TT General', label: '⭐ General' },
    { key: 'common', label: 'Fluent: Common' },
    { key: 'finance', label: 'Fluent: Finance' },
    { key: 'analytics', label: 'Fluent: Analytics' },
    { key: 'governance', label: 'Fluent: Governance' },
    { key: 'tools', label: 'Fluent: Tools' },
    { key: 'communication', label: 'Fluent: People' }
  ];

  const filteredIcons = FLUENT_ICONS_LIST.filter((def) => {
    const matchesCategory = activeCategory === 'all' || def.category === activeCategory;
    if (!matchesCategory) return false;
    if (!searchTerm.trim()) return true;
    const q = searchTerm.toLowerCase().trim();
    return (
      def.label.toLowerCase().includes(q) ||
      def.key.toLowerCase().includes(q) ||
      def.keywords.some((k) => k.toLowerCase().includes(q))
    );
  });

  const currentSelectionDef = FLUENT_ICONS_LIST.find((d) => d.key === tempSelectedKey) || FLUENT_ICONS_LIST[0];

  const handleApply = (): void => {
    onSelectIcon(tempSelectedKey);
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
          Select Fluent UI 2 Icon
        </DialogTitle>
        <DialogBody>
          <DialogContent>
            {/* Search Box */}
            <Input
              className={styles.searchBar}
              contentBefore={<SearchRegular />}
              placeholder="Search icons by keyword (e.g., pound, shield, analytics, team)..."
              value={searchTerm}
              onChange={(e, data) => setSearchTerm(data.value)}
            />

            {/* Category Filter Pills */}
            <div className={styles.categoryTabs}>
              {categories.map((cat) => (
                <button
                  key={cat.key}
                  type="button"
                  className={`${styles.categoryPill} ${activeCategory === cat.key ? styles.categoryPillActive : ''}`}
                  onClick={() => setActiveCategory(cat.key)}
                >
                  {cat.label}
                </button>
              ))}
            </div>

            {/* Icon Grid */}
            <div className={styles.iconGrid}>
              {filteredIcons.map((def) => {
                const isSelected = def.key === tempSelectedKey;
                return (
                  <div
                    key={def.key}
                    className={`${styles.iconItem} ${isSelected ? styles.iconItemActive : ''}`}
                    onClick={() => setTempSelectedKey(def.key)}
                    title={`${def.label} (${def.key})`}
                  >
                    <div className={styles.iconGlyph}>{def.icon}</div>
                    <span className={styles.iconLabel}>{def.label}</span>
                    {isSelected && (
                      <Badge
                        appearance="filled"
                        color="brand"
                        size="extra-small"
                        icon={<CheckmarkRegular />}
                        style={{ marginTop: '4px' }}
                      />
                    )}
                  </div>
                );
              })}
              {filteredIcons.length === 0 && (
                <div style={{ gridColumn: '1 / -1', padding: '32px', textAlign: 'center', color: tokens.colorNeutralForeground3 }}>
                  No Fluent icons matching "{searchTerm}". Try another keyword or switch category.
                </div>
              )}
            </div>

            {/* Selected Icon Preview Bar */}
            <div className={styles.previewFooter}>
              <div className={styles.selectedPreviewBox}>
                <div style={{ fontSize: '24px', color: tokens.colorBrandForeground1, display: 'flex' }}>
                  {currentSelectionDef.icon}
                </div>
                <div>
                  <Subtitle2>{currentSelectionDef.label}</Subtitle2>
                  <Caption1 style={{ color: tokens.colorNeutralForeground3, display: 'block' }}>
                    Key: {currentSelectionDef.key} • Category: {currentSelectionDef.category}
                  </Caption1>
                </div>
              </div>
            </div>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={onDismiss}>
              Cancel
            </Button>
            <Button appearance="primary" onClick={handleApply}>
              Use Selected Icon
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};
