import * as React from 'react';
import type { BaseComponentContext } from '@microsoft/sp-component-base';
import { Stack, Label, Text, Callout, DirectionalHint } from '@fluentui/react';
import { SearchBox, Persona, PersonaSize, IconButton } from '@fluentui/react';
import { searchPeople, getRecentPeople, type PersonSuggestion } from '../../../utils/people';

export interface PeoplePickerFieldProps {
  context: BaseComponentContext;
  label: string;
  description?: string;
  required?: boolean;
  multiSelect?: boolean;
  value: PersonSuggestion[];
  onChange: (people: PersonSuggestion[]) => void;
  errorMessage?: string;
}

export const PeoplePickerField: React.FC<PeoplePickerFieldProps> = ({ context, label, description, required, multiSelect = false, value, onChange, errorMessage }) => {
  const [query, setQuery] = React.useState('');
  const [suggestions, setSuggestions] = React.useState<PersonSuggestion[]>([]);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState('');
  const [showSuggestions, setShowSuggestions] = React.useState(false);
  const debounceRef = React.useRef<number | undefined>(undefined);
  const searchBoxRef = React.useRef<HTMLDivElement>(null);

  const load = React.useCallback(async (q: string) => {
    setLoading(true);
    setError('');
    try {
      const items = await searchPeople(context, q);
      setSuggestions(items);
    } catch {
      setError('Directory search unavailable. Showing limited results.');
      try {
        setSuggestions(await getRecentPeople(context));
      } catch {}
    } finally {
      setLoading(false);
    }
  }, [context]);

  // Don't load initial suggestions - only when user starts typing

  React.useEffect(() => {
    // Debounce typeahead and control suggestion visibility
    if (debounceRef.current) {
      window.clearTimeout(debounceRef.current);
    }
    
    if (query.trim().length >= 2) {
      setShowSuggestions(true);
      debounceRef.current = window.setTimeout(() => {
        load(query);
      }, 300);
    } else {
      setShowSuggestions(false);
      setSuggestions([]);
    }
    
    return () => { if (debounceRef.current) window.clearTimeout(debounceRef.current); };
  }, [query, load]);

  const addPerson = (p: PersonSuggestion) => {
    if (value.some((v: PersonSuggestion) => v.key === p.key)) return;
    onChange(multiSelect ? [...value, p] : [p]);
    setQuery('');
    setShowSuggestions(false);
  };

  const removePerson = (k: string) => {
    onChange(value.filter(v => v.key !== k));
  };

  return (
    <Stack tokens={{ childrenGap: 4 }}>
      {label && <Label required={required}>{label}</Label>}
      {description && <Text variant="small">{description}</Text>}
      <div ref={searchBoxRef} style={{ position: 'relative' }}>
        <SearchBox
          placeholder="Search people..."
          value={query}
          onChange={(_, v) => setQuery(v || '')}
          onSearch={(v) => load(v || '')}
          onEscape={() => { setQuery(''); setShowSuggestions(false); }}
          onClear={() => { setQuery(''); setShowSuggestions(false); }}
          underlined
          aria-label={label || 'Search people'}
        />
        {showSuggestions && searchBoxRef.current && (
          <Callout
            target={searchBoxRef.current}
            directionalHint={DirectionalHint.bottomLeftEdge}
            onDismiss={() => setShowSuggestions(false)}
            gapSpace={0}
            calloutMaxHeight={200}
            styles={{
              root: { zIndex: 1000 },
              calloutMain: { 
                padding: '8px', 
                maxHeight: '200px', 
                overflowY: 'auto',
                border: '1px solid #e1dfdd',
                borderRadius: '2px',
                boxShadow: '0 4px 8px rgba(0,0,0,0.1)'
              }
            }}
          >
            <Stack tokens={{ childrenGap: 4 }} role="listbox" aria-label={`${label} suggestions`}>
              {loading ? (
                <Text style={{ padding: '8px' }}>Loading...</Text>
              ) : suggestions.length > 0 ? (
                suggestions.slice(0, 8).map(s => (
                  <div
                    key={s.key}
                    role="option"
                    aria-selected={false}
                    style={{
                      display: 'flex',
                      alignItems: 'center',
                      padding: '4px 8px',
                      cursor: 'pointer',
                      borderRadius: '2px'
                    }}
                    onMouseEnter={(e) => { e.currentTarget.style.backgroundColor = '#f3f2f1'; }}
                    onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = 'transparent'; }}
                    onClick={() => addPerson(s)}
                  >
                    <Persona text={s.name} secondaryText={s.email} size={PersonaSize.size24} />
                  </div>
                ))
              ) : (
                <Text style={{ padding: '8px', fontStyle: 'italic' }}>No people found</Text>
              )}
            </Stack>
          </Callout>
        )}
      </div>
      {(error || errorMessage) && <Text role="alert" variant="small" styles={{ root: { color: 'var(--errorText, #a4262c)' } }}>{errorMessage || error}</Text>}
      <Stack tokens={{ childrenGap: 8 }}>
        {value.map(v => (
          <Stack key={v.key} horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
            <Persona text={v.name} secondaryText={v.email} size={PersonaSize.size28} />
            <IconButton iconProps={{ iconName: 'Cancel' }} ariaLabel={`Remove ${v.name}`} onClick={() => removePerson(v.key)} />
          </Stack>
        ))}
      </Stack>
    </Stack>
  );
};


