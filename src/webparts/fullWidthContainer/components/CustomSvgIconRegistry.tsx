/**
 * @file CustomSvgIconRegistry.tsx
 * @description Extensible SVG Icon Registry and Unified Icon Renderer for the Full-Width Dashboard.
 * Allows custom corporate SVG icon collections to be seamlessly registered alongside SPFx Fluent UI 2 icons,
 * with dynamic "currentColor" inheritance, custom icon background styling, and visual colour picking.
 *
 * Includes 224 Turner & Townsend brand corporate icons sourced from D:\Playbook\Icons\_Blue,
 * categorised into: TT Built Environment, TT Sustainability & Energy, TT Finance & Data,
 * TT People & Teams, TT Technology & Innovation, TT Awards & Learning, TT Locations & Travel,
 * TT Construction & Tools, TT Operations, TT General.
 */

import * as React from "react";
import {
  DocumentRegular, FolderRegular, MoneyRegular, ReceiptMoneyRegular,
  ArrowTrendingLinesRegular, ChartMultipleRegular, ShieldCheckmarkRegular,
  CheckmarkCircleRegular, LockClosedRegular, GlobeRegular, WrenchRegular,
  PeopleRegular, BuildingRegular, MegaphoneRegular, SparkleRegular,
  StarRegular, AppsRegular, BookOpenRegular, TagRegular, SearchRegular,
  CalendarLtrRegular, LightbulbRegular, TargetRegular, TrophyRegular,
  GaugeRegular, LinkRegular, MailRegular, PhoneRegular, TimerRegular,
  WarningRegular, InfoRegular, KeyRegular, BookmarkRegular, HeartRegular,
  ShareRegular
} from "@fluentui/react-icons";

import { TT_SVG_ICONS } from "./TtSvgIconCollection";

export interface ICustomSvgIcon {
  id: string;
  name: string;
  /** Category shown as a filter tab in the icon picker. */
  category: string;
  keywords: string[];
  /** JSX node for inline React SVGs (original brand icons). */
  svgPathOrNode?: React.ReactNode;
  /** Raw SVG markup string for TT corporate icons (rendered via dangerouslySetInnerHTML). */
  svgMarkup?: string;
}

/** Original handcrafted brand SVG icons kept in JSX for design fidelity. */
const BRAND_SVG_ICONS: ICustomSvgIcon[] = [
  {
    id: "svg-brand-book",
    name: "Brand Playbook",
    category: "Custom Brand SVGs",
    keywords: ["book", "playbook", "manual", "pages", "guidelines"],
    svgPathOrNode: React.createElement("svg", { width: 24, height: 24, viewBox: "0 0 24 24", fill: "none", xmlns: "http://www.w3.org/2000/svg" },
      React.createElement("rect", { x: 3.5, y: 4.5, width: 7, height: 15, rx: 1.5, stroke: "currentColor", strokeWidth: 1.8 }),
      React.createElement("rect", { x: 13.5, y: 4.5, width: 7, height: 15, rx: 1.5, stroke: "currentColor", strokeWidth: 1.8 })
    )
  },
  {
    id: "svg-brand-document",
    name: "Brand Project Document",
    category: "Custom Brand SVGs",
    keywords: ["doc", "file", "contract", "statement"],
    svgPathOrNode: React.createElement("svg", { width: 24, height: 24, viewBox: "0 0 24 24", fill: "none", xmlns: "http://www.w3.org/2000/svg" },
      React.createElement("path", { d: "M6 3.5H14L18.5 8V20.5H6V3.5Z", stroke: "currentColor", strokeWidth: 1.8, strokeLinejoin: "round" }),
      React.createElement("path", { d: "M14 3.5V8H18.5", stroke: "currentColor", strokeWidth: 1.8, strokeLinejoin: "round" }),
      React.createElement("path", { d: "M9 12H15.5M9 15.5H15.5", stroke: "currentColor", strokeWidth: 1.8, strokeLinecap: "round" })
    )
  },
  {
    id: "svg-brand-analytics",
    name: "Brand Analytics Chart",
    category: "Custom Brand SVGs",
    keywords: ["chart", "analytics", "growth", "trending", "kpi"],
    svgPathOrNode: React.createElement("svg", { width: 24, height: 24, viewBox: "0 0 24 24", fill: "none", xmlns: "http://www.w3.org/2000/svg" },
      React.createElement("path", { d: "M4 20.5H20", stroke: "currentColor", strokeWidth: 1.8, strokeLinecap: "round" }),
      React.createElement("path", { d: "M6.5 16.5L10.5 11.5L14 14.5L18.5 7.5", stroke: "currentColor", strokeWidth: 1.8, strokeLinecap: "round", strokeLinejoin: "round" }),
      React.createElement("circle", { cx: 18.5, cy: 7.5, r: 1.5, fill: "currentColor" })
    )
  },
  {
    id: "svg-brand-shield",
    name: "Brand Assurance Shield",
    category: "Custom Brand SVGs",
    keywords: ["shield", "compliance", "security", "audit", "assurance"],
    svgPathOrNode: React.createElement("svg", { width: 24, height: 24, viewBox: "0 0 24 24", fill: "none", xmlns: "http://www.w3.org/2000/svg" },
      React.createElement("path", { d: "M12 3.5L4.5 7V13C4.5 17.5 12 21 12 21C12 21 19.5 17.5 19.5 13V7L12 3.5Z", stroke: "currentColor", strokeWidth: 1.8, strokeLinejoin: "round" }),
      React.createElement("path", { d: "M9.5 12L11 13.5L15 9.5", stroke: "currentColor", strokeWidth: 1.8, strokeLinecap: "round", strokeLinejoin: "round" })
    )
  },
  {
    id: "svg-brand-coins",
    name: "Brand Financial Currency",
    category: "Custom Brand SVGs",
    keywords: ["finance", "pound", "money", "commercial", "cost"],
    svgPathOrNode: React.createElement("svg", { width: 24, height: 24, viewBox: "0 0 24 24", fill: "none", xmlns: "http://www.w3.org/2000/svg" },
      React.createElement("circle", { cx: 12, cy: 12, r: 8.5, stroke: "currentColor", strokeWidth: 1.8 }),
      React.createElement("path", { d: "M13.5 8.5C12.5 8.5 11.5 9 11 10C10.5 11 10.5 12 10.5 13.5H14.5M9.5 12H13.5", stroke: "currentColor", strokeWidth: 1.8, strokeLinecap: "round" }),
      React.createElement("path", { d: "M9.5 15.5H14.5", stroke: "currentColor", strokeWidth: 1.8, strokeLinecap: "round" })
    )
  }
];

/**
 * Merged registry: original brand SVGs + 224 TT corporate SVG icons.
 * TT icons use svgMarkup (raw string) and are rendered via dangerouslySetInnerHTML.
 */
export const CUSTOM_SVG_ICONS: ICustomSvgIcon[] = [
  ...BRAND_SVG_ICONS,
  ...TT_SVG_ICONS.map((entry) => ({
    id: entry.id,
    name: entry.name,
    category: entry.category,
    keywords: entry.keywords,
    svgMarkup: entry.svgMarkup
  }))
];

/** Registers additional custom SVG icons into the runtime registry. */
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
 * Unified Icon Renderer.
 * Supports JSX svgPathOrNode (brand icons), raw svgMarkup string (TT corporate icons),
 * and standard Fluent UI 2 icon keys. Dynamic colour applied via CSS color -> currentColor.
 */
export function renderUnifiedIcon(
  iconKey?: string,
  iconColor?: string,
  fontSize: string = "20px"
): React.ReactElement {
  const customIcon = CUSTOM_SVG_ICONS.find((i) => i.id === iconKey || i.name === iconKey);

  if (customIcon) {
    const containerStyle: React.CSSProperties = {
      display: "inline-flex",
      alignItems: "center",
      justifyContent: "center",
      color: iconColor || "currentColor",
      width: fontSize,
      height: fontSize,
      lineHeight: 1
    };

    if (customIcon.svgMarkup) {
      return React.createElement("span", {
        style: containerStyle,
        // SVG content is generated at build time from TT_SVG_ICONS (no user input path)
        // eslint-disable-next-line react/no-danger
        dangerouslySetInnerHTML: {
          __html: customIcon.svgMarkup.replace("<svg ", `<svg width="${fontSize}" height="${fontSize}" `)
        }
      });
    }

    return React.createElement("span", { style: containerStyle }, customIcon.svgPathOrNode);
  }

  const iconStyle: React.CSSProperties = { color: iconColor || "currentColor", fontSize };

  switch (iconKey) {
    case "Document": return React.createElement(DocumentRegular, { style: iconStyle });
    case "Folder": return React.createElement(FolderRegular, { style: iconStyle });
    case "Financial": return React.createElement(MoneyRegular, { style: iconStyle });
    case "ReceiptMoney": return React.createElement(ReceiptMoneyRegular, { style: iconStyle });
    case "TimelineProgress": return React.createElement(ArrowTrendingLinesRegular, { style: iconStyle });
    case "ChartMultiple": return React.createElement(ChartMultipleRegular, { style: iconStyle });
    case "ComplianceAudit": return React.createElement(ShieldCheckmarkRegular, { style: iconStyle });
    case "CheckList": return React.createElement(CheckmarkCircleRegular, { style: iconStyle });
    case "Lock": return React.createElement(LockClosedRegular, { style: iconStyle });
    case "Globe": return React.createElement(GlobeRegular, { style: iconStyle });
    case "Wrench": return React.createElement(WrenchRegular, { style: iconStyle });
    case "People": return React.createElement(PeopleRegular, { style: iconStyle });
    case "Building": return React.createElement(BuildingRegular, { style: iconStyle });
    case "Megaphone": return React.createElement(MegaphoneRegular, { style: iconStyle });
    case "Sparkle": return React.createElement(SparkleRegular, { style: iconStyle });
    case "Star": return React.createElement(StarRegular, { style: iconStyle });
    case "AppIconDefault": return React.createElement(AppsRegular, { style: iconStyle });
    case "Tag": return React.createElement(TagRegular, { style: iconStyle });
    case "Search": return React.createElement(SearchRegular, { style: iconStyle });
    case "Calendar": return React.createElement(CalendarLtrRegular, { style: iconStyle });
    case "Lightbulb": return React.createElement(LightbulbRegular, { style: iconStyle });
    case "Target": return React.createElement(TargetRegular, { style: iconStyle });
    case "Trophy": return React.createElement(TrophyRegular, { style: iconStyle });
    case "Gauge": return React.createElement(GaugeRegular, { style: iconStyle });
    case "Link": return React.createElement(LinkRegular, { style: iconStyle });
    case "Mail": return React.createElement(MailRegular, { style: iconStyle });
    case "Phone": return React.createElement(PhoneRegular, { style: iconStyle });
    case "Timer": return React.createElement(TimerRegular, { style: iconStyle });
    case "Warning": return React.createElement(WarningRegular, { style: iconStyle });
    case "Info": return React.createElement(InfoRegular, { style: iconStyle });
    case "Key": return React.createElement(KeyRegular, { style: iconStyle });
    case "Bookmark": return React.createElement(BookmarkRegular, { style: iconStyle });
    case "Heart": return React.createElement(HeartRegular, { style: iconStyle });
    case "Share": return React.createElement(ShareRegular, { style: iconStyle });
    case "BookAnswers":
    default:
      return React.createElement(BookOpenRegular, { style: iconStyle });
  }
}
