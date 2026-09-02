/**
 * @file TermStorePicker.tsx
 * @description SharePoint Global Term Store tag picker component with Fluent UI 2 tags.
 */

import * as React from 'react';
import {
  makeStyles,
  tokens,
  shorthands,
  Tag,
  TagGroup,
  Input,
  Button,
  Caption1
} from '@fluentui/react-components';
import { TagRegular, AddRegular, SearchRegular } from '@fluentui/react-icons';
import { ITermStoreTag } from '../models/IContainerModels';
import { TaxonomyService } from '../services/TaxonomyService';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
    width: '100%'
  },
  tagContainer: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '6px',
    alignItems: 'center'
  },
  pickerBox: {
    ...shorthands.padding('12px'),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorNeutralBackground2,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke1),
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
    marginTop: '6px'
  },
  suggestionsList: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '6px',
    maxHeight: '130px',
    overflowY: 'auto',
    ...shorthands.padding('4px')
  },
  suggestionTag: {
    cursor: 'pointer',
    '&:hover': {
      ...shorthands.borderColor(tokens.colorBrandStroke1)
    }
  },
  addNewRow: {
    display: 'flex',
    gap: '8px',
    alignItems: 'center',
    marginTop: '4px'
  }
});

export interface ITermStorePickerProps {
  selectedTags?: ITermStoreTag[];
  onChange: (tags: ITermStoreTag[]) => void;
  isEditMode?: boolean;
}

export const TermStorePicker: React.FC<ITermStorePickerProps> = ({
  selectedTags = [],
  onChange,
  isEditMode = false
}) => {
  const styles = useStyles();
  const [isPickerOpen, setIsPickerOpen] = React.useState<boolean>(false);
  const [searchQuery, setSearchQuery] = React.useState<string>('');
  const [allTerms, setAllTerms] = React.useState<ITermStoreTag[]>([]);
  const [newTermInput, setNewTermInput] = React.useState<string>('');

  React.useEffect(() => {
    void TaxonomyService.getTerms().then((terms) => setAllTerms(terms));
  }, []);

  const handleSelectTerm = (term: ITermStoreTag): void => {
    if (!selectedTags.some((t) => t.id === term.id || t.label.toLowerCase() === term.label.toLowerCase())) {
      onChange([...selectedTags, term]);
    }
  };

  const handleRemoveTag = (tagId: string): void => {
    onChange(selectedTags.filter((t) => t.id !== tagId));
  };

  const handleCreateTerm = (): void => {
    if (!newTermInput.trim()) return;
    const created = TaxonomyService.addTerm(newTermInput.trim());
    setAllTerms([...allTerms, created]);
    handleSelectTerm(created);
    setNewTermInput('');
  };

  const filteredSuggestions = allTerms.filter(
    (t) =>
      !selectedTags.some((st) => st.label.toLowerCase() === t.label.toLowerCase()) &&
      (!searchQuery || t.label.toLowerCase().includes(searchQuery.toLowerCase()))
  );

  return (
    <div className={styles.root}>
      <TagGroup onDismiss={(e, data) => handleRemoveTag(data.value)}>
        <div className={styles.tagContainer}>
          {selectedTags.map((tag) => (
            <Tag
              key={tag.id}
              value={tag.id}
              shape="rounded"
              appearance="outline"
              icon={<TagRegular />}
              dismissible={isEditMode}
            >
              {tag.label}
            </Tag>
          ))}

          {isEditMode && (
            <Button
              size="small"
              appearance="subtle"
              icon={<AddRegular />}
              onClick={() => setIsPickerOpen(!isPickerOpen)}
            >
              {isPickerOpen ? 'Done' : '+ Add Term Store Tag'}
            </Button>
          )}
        </div>
      </TagGroup>

      {isEditMode && isPickerOpen && (
        <div className={styles.pickerBox}>
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
            Select tags from SharePoint Global Term Store:
          </Caption1>
          <Input
            size="small"
            contentBefore={<SearchRegular />}
            placeholder="Search Term Store..."
            value={searchQuery}
            onChange={(_, data) => setSearchQuery(data.value)}
          />

          <div className={styles.suggestionsList}>
            {filteredSuggestions.map((term) => (
              <Tag
                key={term.id}
                shape="rounded"
                appearance="brand"
                className={styles.suggestionTag}
                onClick={() => handleSelectTerm(term)}
              >
                + {term.label}
              </Tag>
            ))}
            {filteredSuggestions.length === 0 && (
              <Caption1 style={{ color: tokens.colorNeutralForeground4 }}>
                No matching terms. Create a custom tag below.
              </Caption1>
            )}
          </div>

          <div className={styles.addNewRow}>
            <Input
              size="small"
              placeholder="New term label..."
              value={newTermInput}
              onChange={(_, data) => setNewTermInput(data.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  e.preventDefault();
                  handleCreateTerm();
                }
              }}
            />
            <Button size="small" appearance="primary" onClick={handleCreateTerm}>
              Add Term
            </Button>
          </div>
        </div>
      )}
    </div>
  );
};
