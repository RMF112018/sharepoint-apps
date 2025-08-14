import * as React from 'react';
import type { BaseComponentContext } from '@microsoft/sp-component-base';
import { Stack, Text, TextField, IconButton, TooltipHost } from '@fluentui/react';
import { PeoplePickerField } from './PeoplePickerField';
import type { FieldDef } from './FormRenderer';

export type FieldValue = string | number | boolean | unknown[] | null;

export interface KeyDatesMilestonesProps {
  context: BaseComponentContext;
  tasks: FieldDef[];
  catalog: FieldDef[];
  values: Record<string, FieldValue>;
  setValue: (key: string, v: FieldValue) => void;
}

function keys(base: string): { owner: string; deadline: string; notes: string } {
  return {
    owner: `${base}Owner`,
    deadline: `${base}Deadline`,
    notes: `${base}Notes`
  };
}

function toLocalDateTimeValue(d: Date): string {
  const pad = (n: number) => (n < 10 ? `0${n}` : `${n}`);
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

export const KeyDatesMilestones: React.FC<KeyDatesMilestonesProps> = ({ context, tasks, catalog, values, setValue }) => {
  // Locate proposal due field
  let dueField: FieldDef | undefined;
  for (const f of catalog) {
    if (/proposal\s*due/i.test(f.label)) { dueField = f; break; }
  }
  const dueVal = dueField ? String(values[dueField.field_key] ?? '') : '';
  const dueDate = dueVal ? new Date(dueVal) : null;

  const setRelative = (taskKey: string, daysBefore: number) => {
    if (!dueDate || String(dueDate) === 'Invalid Date') return;
    const d = new Date(dueDate.getTime());
    d.setDate(d.getDate() - daysBefore);
    setValue(keys(taskKey).deadline, toLocalDateTimeValue(d));
  };

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      {tasks.map((t) => {
        const k = keys(t.field_key);
        const owner = (values[k.owner] as any[]) || [];
        const deadlineStr = String(values[k.deadline] ?? '');
        const notes = String(values[k.notes] ?? '');
        const deadline = deadlineStr ? new Date(deadlineStr) : null;
        const conflict = !!(dueDate && String(dueDate) !== 'Invalid Date' && deadline && String(deadline) !== 'Invalid Date' && deadline.getTime() > dueDate.getTime());

        return (
          <Stack key={t.field_key} horizontal wrap verticalAlign="end" tokens={{ childrenGap: 12 }}>
            <Stack.Item grow={1} styles={{ root: { minWidth: 220 } }}>
              <Text variant="mediumPlus">{t.label}</Text>
            </Stack.Item>
            <Stack.Item grow={1} styles={{ root: { minWidth: 260 } }}>
              <PeoplePickerField
                context={context}
                label="Owner"
                value={owner}
                onChange={(p) => setValue(k.owner, p)}
              />
            </Stack.Item>
            <Stack.Item grow={1} styles={{ root: { minWidth: 220 } }}>
              <TextField
                label="Deadline"
                type="datetime-local"
                value={deadlineStr}
                onChange={(_, v) => setValue(k.deadline, v ?? '')}
                onGetErrorMessage={() => (conflict ? 'Occurs after Proposal Due date' : '')}
                validateOnLoad={false}
              />
              <Stack horizontal tokens={{ childrenGap: 4 }}>
                <TooltipHost content="Set to 1 day before Proposal Due">
                  <IconButton iconProps={{ iconName: 'Back' }} text="-1d" disabled={!dueDate} onClick={() => setRelative(t.field_key, 1)} />
                </TooltipHost>
                <TooltipHost content="Set to 3 days before Proposal Due">
                  <IconButton iconProps={{ iconName: 'Back' }} text="-3d" disabled={!dueDate} onClick={() => setRelative(t.field_key, 3)} />
                </TooltipHost>
                <TooltipHost content="Set to 7 days before Proposal Due">
                  <IconButton iconProps={{ iconName: 'Back' }} text="-7d" disabled={!dueDate} onClick={() => setRelative(t.field_key, 7)} />
                </TooltipHost>
              </Stack>
            </Stack.Item>
            <Stack.Item grow={2} styles={{ root: { minWidth: 260 } }}>
              <TextField
                label="Notes"
                multiline
                value={notes}
                onChange={(_, v) => setValue(k.notes, v ?? '')}
              />
            </Stack.Item>
          </Stack>
        );
      })}
    </Stack>
  );
};


