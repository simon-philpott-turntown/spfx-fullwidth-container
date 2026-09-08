/**
 * @file ComposableContentSection.tsx
 * @description Composable Content Area for top-of-webpart and top-of-section content authoring.
 * Supports all 15 inner item types (buttons, filter buttons, process models, text, etc.),
 * alignment (Left, Centre, Right), insertion bars, and inline editing.
 */

import * as React from 'react';
import {
  makeStyles,
  tokens,
  shorthands,
  Button,
  Subtitle2,
  Caption1,
  Divider,
  Title3,
  Body1
} from '@fluentui/react-components';
import {
  EditRegular,
  DismissRegular,
  CursorClickRegular,
  MegaphoneRegular,
  OpenRegular,
  ArrowUpRegular,
  ArrowDownRegular
} from '@fluentui/react-icons';
import { ICardItem, ICardItemType } from '../models/IContainerModels';
import { InsertionBar } from './InsertionBar';
import { CardItemPropertyEditor } from './CardItemPropertyEditor';
import { RichTextEditable } from './RichTextEditable';
import { LiveDataRenderer } from './LiveDataRenderer';
import { TermStorePicker } from './TermStorePicker';
import { FilterButtonsRenderer } from './FilterButtonsRenderer';
import { ProcessModelRenderer } from './ProcessModelRenderer';
import { IAssetPickerService } from '../services/IAssetPickerService';

const useStyles = makeStyles({
  container: {
    width: '100%',
    display: 'flex',
    flexDirection: 'column',
    boxSizing: 'border-box',
    marginBottom: '8px'
  },
  itemWrapper: {
    position: 'relative',
    width: '100%',
    boxSizing: 'border-box',
    padding: '4px 0'
  },
  itemEditControls: {
    position: 'absolute',
    right: 0,
    top: 0,
    zIndex: 10,
    display: 'flex',
    gap: '2px',
    backgroundColor: 'rgba(255, 255, 255, 0.9)',
    borderRadius: '4px',
    padding: '1px',
    boxShadow: tokens.shadow4
  },
  ctaBox: {
    padding: '16px',
    borderRadius: '8px',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    display: 'flex',
    flexDirection: 'column',
    gap: '6px'
  },
  editorialBox: {
    padding: '12px',
    borderRadius: '8px',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    display: 'flex',
    flexDirection: 'column',
    gap: '4px'
  },
  heroBox: {
    padding: '24px',
    borderRadius: '8px',
    backgroundColor: tokens.colorNeutralBackground3,
    display: 'flex',
    flexDirection: 'column',
    gap: '8px'
  },
  galleryGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fill, minmax(110px, 1fr))',
    gap: '8px',
    width: '100%'
  },
  galleryImg: {
    width: '100%',
    height: '90px',
    objectFit: 'cover',
    borderRadius: '6px'
  },
  quickLinksRow: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '8px',
    width: '100%'
  }
});

export interface IComposableContentSectionProps {
  items?: ICardItem[];
  isEditMode?: boolean;
  contextTitle?: string;
  assetPickerService?: IAssetPickerService;
  onUpdateItems: (items: ICardItem[]) => void;
}

export const ComposableContentSection: React.FC<IComposableContentSectionProps> = ({
  items = [],
  isEditMode = false,
  contextTitle = 'Add content section',
  assetPickerService,
  onUpdateItems
}) => {
  const styles = useStyles();
  const [editingItem, setEditingItem] = React.useState<{ item: ICardItem; index: number } | null>(null);

  const handleInsertItem = (itemType: ICardItemType, index: number): void => {
    const newItem: ICardItem = {
      id: `item-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`,
      type: itemType,
      alignment: 'left',
      text: itemType === 'text' ? 'New content block. Click here to edit description and information.' : undefined,
      buttonLabel: itemType === 'button' ? 'Action Button' : undefined,
      buttonUrl: itemType === 'button' ? '#' : undefined,
      buttonVariant: 'primary',
      ctaHeading: itemType === 'cta' ? 'Call to action banner' : undefined,
      ctaDescription: itemType === 'cta' ? 'Key guidance and directives for project delivery.' : undefined,
      ctaButtonText: itemType === 'cta' ? 'View guidance' : undefined,
      editorialTitle: itemType === 'editorial' ? 'Strategic Insight' : undefined,
      editorialBody: itemType === 'editorial' ? 'Highlighting commercial models, ESG compliance, and capital investment.' : undefined,
      heroTitle: itemType === 'hero' ? 'Strategic Programme Overview' : undefined,
      heroSubtitle: itemType === 'hero' ? 'Enterprise lifecycle management and capability alignment.' : undefined,
      linkText: itemType === 'link' ? 'Resource document link' : undefined,
      linkUrl: itemType === 'link' ? '#' : undefined,
      quickLinks: itemType === 'quickLinks' ? [
        { label: 'Project Guidance', url: '#' },
        { label: 'Governance Toolkit', url: '#' }
      ] : undefined,
      termStoreTags: itemType === 'termStoreTags' ? [
        { id: 'sec-infra', label: 'Infrastructure', termSetName: 'Our Sectors' },
        { id: 'seg-comm', label: 'Commercial advisory', termSetName: 'Our Segments' }
      ] : undefined,
      filterButtons: itemType === 'filterButtons' ? [
        { id: 'btn-1', label: 'Programme advisory' },
        { id: 'btn-2', label: 'Cost and commercial management' },
        { id: 'btn-3', label: 'Controls and performance' },
        { id: 'btn-4', label: 'Project management' },
        { id: 'btn-5', label: 'Procurement and supply chain' },
        { id: 'btn-6', label: 'Sustainability' },
        { id: 'btn-7', label: 'Digital' }
      ] : undefined,
      processSteps: itemType === 'processModel' ? [
        {
          id: 'stg-1',
          stageNumber: 'Stage 1',
          title: 'Shape',
          description: 'Define the problem and the case for acting',
          metricBadge: '4 capabilities',
          actionType: 'filter'
        },
        {
          id: 'stg-2',
          stageNumber: 'Stage 2',
          title: 'Plan',
          description: 'Establish the approach, baseline and controls',
          metricBadge: '4 capabilities',
          actionType: 'filter'
        },
        {
          id: 'stg-3',
          stageNumber: 'Stage 3',
          title: 'Source',
          description: 'Take it to market and contract for delivery',
          metricBadge: '2 capabilities',
          actionType: 'filter'
        },
        {
          id: 'stg-4',
          stageNumber: 'Stage 4',
          title: 'Deliver',
          description: 'Execute, control and assure',
          metricBadge: '3 capabilities',
          actionType: 'filter'
        },
        {
          id: 'stg-5',
          stageNumber: 'Stage 5',
          title: 'Realise',
          description: 'Hand over, close out and bank the learning',
          metricBadge: '3 capabilities',
          actionType: 'filter'
        }
      ] : undefined
    };

    const next = [...items];
    next.splice(index, 0, newItem);
    onUpdateItems(next);
  };

  const handleRemoveItem = (id: string): void => {
    onUpdateItems(items.filter((i) => i.id !== id));
  };

  const handleUpdateItem = (updated: ICardItem): void => {
    if (!editingItem) return;
    const next = [...items];
    next[editingItem.index] = updated;
    onUpdateItems(next);
  };

  const renderItemContent = (item: ICardItem, idx: number): React.ReactElement => {
    const alignVal = item.alignment || 'left';
    const textAlignVal = alignVal === 'center' ? 'center' : alignVal === 'right' ? 'right' : 'left';
    const justifyVal = alignVal === 'center' ? 'center' : alignVal === 'right' ? 'flex-end' : 'flex-start';

    return (
      <div
        key={item.id}
        className={styles.itemWrapper}
        style={{
          textAlign: textAlignVal,
          display: 'flex',
          flexDirection: 'column',
          alignItems: alignVal === 'center' ? 'center' : alignVal === 'right' ? 'flex-end' : 'stretch'
        }}
      >
        {isEditMode && (
          <div className={styles.itemEditControls}>
            <Button
              size="small"
              appearance="subtle"
              icon={<ArrowUpRegular />}
              disabled={idx === 0}
              onClick={() => {
                if (idx === 0) return;
                const next = [...items];
                const temp = next[idx - 1];
                next[idx - 1] = next[idx];
                next[idx] = temp;
                onUpdateItems(next);
              }}
              title="Move item up"
            />
            <Button
              size="small"
              appearance="subtle"
              icon={<ArrowDownRegular />}
              disabled={idx === items.length - 1}
              onClick={() => {
                if (idx === items.length - 1) return;
                const next = [...items];
                const temp = next[idx + 1];
                next[idx + 1] = next[idx];
                next[idx] = temp;
                onUpdateItems(next);
              }}
              title="Move item down"
            />
            <Button
              size="small"
              appearance="subtle"
              icon={<EditRegular />}
              onClick={() => setEditingItem({ item, index: idx })}
              title="Edit item properties & alignment"
            />
            <Button
              size="small"
              appearance="subtle"
              icon={<DismissRegular />}
              onClick={() => handleRemoveItem(item.id)}
              title="Remove content item"
            />
          </div>
        )}

        {item.type === 'text' && (
          <div style={{ width: '100%', textAlign: textAlignVal }}>
            <RichTextEditable
              html={item.text || ''}
              isEditMode={isEditMode}
              placeholder="Enter text content..."
              onChange={(newHtml) => {
                const next = [...items];
                next[idx] = { ...next[idx], text: newHtml };
                onUpdateItems(next);
              }}
              style={{
                color: tokens.colorNeutralForeground1,
                fontSize: '0.95rem',
                textAlign: textAlignVal
              }}
            />
          </div>
        )}

        {item.type === 'button' && (
          <div style={{ display: 'flex', width: '100%', justifyContent: justifyVal }}>
            <Button
              appearance={item.buttonVariant || 'primary'}
              as="a"
              href={item.buttonUrl || '#'}
              icon={<CursorClickRegular />}
            >
              {item.buttonLabel || 'Action Button'}
            </Button>
          </div>
        )}

        {item.type === 'cta' && (
          <div
            className={styles.ctaBox}
            style={{
              textAlign: textAlignVal,
              alignItems: alignVal === 'center' ? 'center' : alignVal === 'right' ? 'flex-end' : 'flex-start',
              ...(item.ctaBgUrl ? {
                backgroundImage: `linear-gradient(rgba(255,255,255,0.85), rgba(255,255,255,0.85)), url(${item.ctaBgUrl})`,
                backgroundSize: 'cover',
                backgroundPosition: 'center'
              } : {})
            }}
          >
            <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
              <MegaphoneRegular style={{ color: tokens.colorBrandForeground1 }} />
              <Subtitle2>{item.ctaHeading || 'Call to Action'}</Subtitle2>
            </div>
            <Caption1 style={{ color: tokens.colorNeutralForeground2, textAlign: textAlignVal }}>
              {item.ctaDescription || 'Guidance details here.'}
            </Caption1>
            {item.ctaButtonText && (
              <Button size="small" appearance="primary" as="a" href={item.ctaButtonUrl || '#'}>
                {item.ctaButtonText}
              </Button>
            )}
          </div>
        )}

        {item.type === 'editorial' && (
          <div className={styles.editorialBox} style={{ textAlign: textAlignVal }}>
            {item.editorialImageUrl && (
              <div style={{ width: '100%', marginBottom: '8px', borderRadius: '6px', overflow: 'hidden' }}>
                <img
                  src={item.editorialImageUrl}
                  alt={item.editorialTitle || 'Editorial image'}
                  style={{ width: '100%', maxHeight: '140px', objectFit: 'cover', display: 'block' }}
                />
              </div>
            )}
            {item.editorialKicker && (
              <Caption1 style={{ color: tokens.colorBrandForeground1, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.5px' }}>
                {item.editorialKicker}
              </Caption1>
            )}
            <Subtitle2 style={{ fontWeight: 600 }}>{item.editorialTitle || 'Editorial Summary'}</Subtitle2>
            {item.editorialBody && (
              <Body1 style={{ color: tokens.colorNeutralForeground2, fontSize: '0.9rem', textAlign: textAlignVal }}>
                {item.editorialBody}
              </Body1>
            )}
            {item.editorialUrl && (
              <Button
                size="small"
                appearance="subtle"
                as="a"
                href={item.editorialUrl}
                icon={<OpenRegular />}
                iconPosition="after"
                style={{
                  alignSelf: alignVal === 'center' ? 'center' : alignVal === 'right' ? 'flex-end' : 'flex-start',
                  padding: 0,
                  height: 'auto',
                  marginTop: '4px'
                }}
              >
                Read full article
              </Button>
            )}
          </div>
        )}

        {item.type === 'hero' && (
          <div
            className={styles.heroBox}
            style={{
              textAlign: textAlignVal,
              alignItems: alignVal === 'center' ? 'center' : alignVal === 'right' ? 'flex-end' : 'flex-start',
              ...(item.heroBgUrl ? {
                backgroundImage: `linear-gradient(rgba(0,0,0,0.5), rgba(0,0,0,0.7)), url(${item.heroBgUrl})`,
                backgroundSize: 'cover',
                backgroundPosition: 'center',
                color: '#FFFFFF'
              } : {})
            }}
          >
            <Title3 style={item.heroBgUrl ? { color: '#FFFFFF' } : undefined}>
              {item.heroTitle || 'Strategic Title'}
            </Title3>
            <Body1 style={item.heroBgUrl ? { color: 'rgba(255,255,255,0.9)' } : { color: tokens.colorNeutralForeground2 }}>
              {item.heroSubtitle || 'Highlighting key objectives and requirements.'}
            </Body1>
          </div>
        )}

        {item.type === 'link' && (
          <div style={{ padding: '4px 0', width: '100%', display: 'flex', justifyContent: justifyVal }}>
            <Button
              appearance="subtle"
              icon={<OpenRegular />}
              iconPosition="after"
              as="a"
              href={item.linkUrl || item.buttonUrl || '#'}
              style={{ fontWeight: 600, color: tokens.colorBrandForeground1 }}
            >
              {item.linkText || item.buttonLabel || 'Direct resource link'}
            </Button>
          </div>
        )}

        {item.type === 'divider' && (
          <Divider
            style={{
              margin: '8px 0',
              borderStyle: item.dividerStyle || 'solid'
            }}
          />
        )}

        {item.type === 'image' && item.imageUrl && (
          <div style={{ width: '100%', textAlign: textAlignVal }}>
            <img
              src={item.imageUrl}
              alt={item.imageAlt || item.imageCaption || 'Section image'}
              style={{ width: '100%', maxHeight: '240px', objectFit: 'cover', borderRadius: '6px' }}
            />
            {item.imageCaption && (
              <Caption1 style={{ color: tokens.colorNeutralForeground4, display: 'block', marginTop: '4px', textAlign: textAlignVal }}>
                {item.imageCaption}
              </Caption1>
            )}
          </div>
        )}

        {item.type === 'gallery' && item.galleryImages && (
          <div className={styles.galleryGrid} style={{ justifyContent: justifyVal }}>
            {item.galleryImages.map((img, i) => (
              <div key={i} style={{ position: 'relative' }}>
                <img src={img.url} alt={img.caption || 'Gallery image'} className={styles.galleryImg} />
                {img.caption && (
                  <Caption1 style={{ fontSize: '0.7rem', color: tokens.colorNeutralForeground3, display: 'block', marginTop: '2px', textAlign: textAlignVal }}>
                    {img.caption}
                  </Caption1>
                )}
              </div>
            ))}
          </div>
        )}

        {item.type === 'quickLinks' && item.quickLinks && (
          <div className={styles.quickLinksRow} style={{ justifyContent: justifyVal }}>
            {item.quickLinks.map((ql, i) => (
              <Button key={i} size="small" appearance="outline" as="a" href={ql.url} icon={<OpenRegular />}>
                {ql.label}
              </Button>
            ))}
          </div>
        )}

        {item.type === 'video' && item.videoUrl && (
          <div style={{ width: '100%', textAlign: textAlignVal }}>
            <video
              src={item.videoUrl}
              controls
              style={{ width: '100%', maxHeight: '240px', borderRadius: '6px' }}
            />
          </div>
        )}

        {item.type === 'liveData' && item.liveDataConfig && (
          <div style={{ width: '100%', textAlign: textAlignVal }}>
            <LiveDataRenderer config={item.liveDataConfig} isEditMode={isEditMode} />
          </div>
        )}

        {item.type === 'termStoreTags' && (
          <div style={{ width: '100%', display: 'flex', justifyContent: justifyVal }}>
            <TermStorePicker
              selectedTags={item.termStoreTags || []}
              onChange={(tags) => {
                const next = [...items];
                next[idx] = { ...next[idx], termStoreTags: tags };
                onUpdateItems(next);
              }}
              isEditMode={isEditMode}
            />
          </div>
        )}

        {item.type === 'filterButtons' && item.filterButtons && (
          <FilterButtonsRenderer
            buttons={item.filterButtons}
            activeFilterId={item.activeFilterId}
            alignment={item.alignment || 'left'}
            onSelectFilter={(selectedBtn) => {
              const next = [...items];
              next[idx] = { ...next[idx], activeFilterId: selectedBtn ? selectedBtn.id : undefined };
              onUpdateItems(next);

              window.dispatchEvent(
                new CustomEvent('dashboard:card-filter-apply', {
                  detail: {
                    sourceItemId: item.id,
                    filterValue: selectedBtn ? (selectedBtn.filterValue || selectedBtn.label) : ''
                  }
                })
              );
            }}
            isEditMode={isEditMode}
            onEdit={() => setEditingItem({ item, index: idx })}
          />
        )}

        {item.type === 'processModel' && item.processSteps && (
          <ProcessModelRenderer
            steps={item.processSteps}
            activeStepId={item.activeFilterId}
            alignment={item.alignment || 'left'}
            onSelectStep={(selectedStep) => {
              const next = [...items];
              next[idx] = { ...next[idx], activeFilterId: selectedStep ? selectedStep.id : undefined };
              onUpdateItems(next);

              window.dispatchEvent(
                new CustomEvent('dashboard:card-filter-apply', {
                  detail: {
                    sourceItemId: item.id,
                    filterValue: selectedStep ? (selectedStep.filterValue || selectedStep.title) : ''
                  }
                })
              );
            }}
            isEditMode={isEditMode}
            onEdit={() => setEditingItem({ item, index: idx })}
          />
        )}
      </div>
    );
  };

  return (
    <div className={styles.container}>
      {items.map((item, idx) => (
        <React.Fragment key={item.id}>
          {isEditMode && (
            <InsertionBar
              onInsert={(type) => handleInsertItem(type, idx)}
              contextTitle={contextTitle}
            />
          )}
          {renderItemContent(item, idx)}
        </React.Fragment>
      ))}

      {isEditMode && (
        <InsertionBar
          alwaysVisible={items.length === 0}
          onInsert={(type) => handleInsertItem(type, items.length)}
          contextTitle={contextTitle}
        />
      )}

      {/* Property Editor Dialog for Configurable Items */}
      {editingItem && (
        <CardItemPropertyEditor
          isOpen={true}
          item={editingItem.item}
          assetPickerService={assetPickerService}
          onDismiss={() => setEditingItem(null)}
          onSave={handleUpdateItem}
        />
      )}
    </div>
  );
};
