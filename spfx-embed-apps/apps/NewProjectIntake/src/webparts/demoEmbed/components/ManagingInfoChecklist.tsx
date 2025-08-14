import * as React from 'react';
import type { BaseComponentContext } from '@microsoft/sp-component-base';
import { Stack, Text, Toggle, TextField, IconButton, TooltipHost, Separator } from '@fluentui/react';
import { PeoplePickerField } from './PeoplePickerField';
import type { FieldDef } from './FormRenderer';

export type FieldValue = string | number | boolean | unknown[] | null;

export interface ManagingInfoChecklistProps {
  context: BaseComponentContext;
  tasks: FieldDef[];
  values: Record<string, FieldValue>;
  setValue: (key: string, v: FieldValue) => void;
}

function rowKeys(baseKey: string): { done: string; responsible: string; deadline: string; notes: string; updatedBy: string; updatedAt: string } {
  return {
    done: `${baseKey}Done`,
    responsible: `${baseKey}Responsible`,
    deadline: `${baseKey}Deadline`,
    notes: `${baseKey}Notes`,
    updatedBy: `${baseKey}UpdatedBy`,
    updatedAt: `${baseKey}UpdatedAt`
  };
}

export const ManagingInfoChecklist: React.FC<ManagingInfoChecklistProps> = ({ context, tasks, values, setValue }) => {
  const onRowChange = (baseKey: string): void => {
    const { updatedBy, updatedAt } = rowKeys(baseKey);
    setValue(updatedBy, context.pageContext?.user?.displayName || '');
    setValue(updatedAt, new Date().toISOString());
  };

  const markAll = (complete: boolean): void => {
    for (const t of tasks) {
      const { done } = rowKeys(t.field_key);
      setValue(done, complete);
      onRowChange(t.field_key);
    }
  };

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
        <Text variant="large">Managing Information</Text>
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <IconButton iconProps={{ iconName: 'CheckMark' }} text="Mark all complete" onClick={() => markAll(true)} />
          <IconButton iconProps={{ iconName: 'Clear' }} text="Clear all" onClick={() => markAll(false)} />
        </Stack>
      </Stack>
      <Separator />
      <Stack tokens={{ childrenGap: 12 }}>
        {tasks.map((t) => {
          const k = rowKeys(t.field_key);
          const isDone = Boolean(values[k.done]);
          const responsible = (values[k.responsible] as any[]) || [];
          const deadline = String(values[k.deadline] ?? '');
          const notes = String(values[k.notes] ?? '');
          const updatedBy = String(values[k.updatedBy] ?? '');
          const updatedAt = String(values[k.updatedAt] ?? '');
          const hover = updatedBy && updatedAt ? `${updatedBy} • ${new Date(updatedAt).toLocaleString()}` : '';

          return (
            <TooltipHost key={t.field_key} content={hover || undefined}>
              <Stack horizontal wrap verticalAlign="center" tokens={{ childrenGap: 12 }}>
                <Stack.Item grow={1} styles={{ root: { minWidth: 220 } }}>
                  <Text>{t.label}</Text>
                </Stack.Item>
                <Stack.Item styles={{ root: { minWidth: 120 } }}>
                  <Toggle
                    label=""
                    checked={isDone}
                    onChange={(_, val) => { setValue(k.done, !!val); onRowChange(t.field_key); }}
                  />
                </Stack.Item>
                <Stack.Item grow={1} styles={{ root: { minWidth: 260 } }}>
                  <PeoplePickerField
                    context={context}
                    label="Responsible"
                    multiSelect
                    value={responsible}
                    onChange={(people) => { setValue(k.responsible, people); onRowChange(t.field_key); }}
                  />
                </Stack.Item>
                <Stack.Item grow={1} styles={{ root: { minWidth: 180 } }}>
                  <TextField
                    label="Deadline/Frequency"
                    value={deadline}
                    onChange={(_, v) => { setValue(k.deadline, v ?? ''); onRowChange(t.field_key); }}
                  />
                </Stack.Item>
                <Stack.Item grow={2} styles={{ root: { minWidth: 220 } }}>
                  <TextField
                    label="Notes"
                    value={notes}
                    onChange={(_, v) => { setValue(k.notes, v ?? ''); onRowChange(t.field_key); }}
                  />
                </Stack.Item>
                <Stack.Item>
                  <TooltipHost content={notes ? notes : 'Add a quick note'}>
                    <IconButton iconProps={{ iconName: 'Comment' }} ariaLabel="Row note" />
                  </TooltipHost>
                </Stack.Item>
              </Stack>
            </TooltipHost>
          );
        })}
      </Stack>
    </Stack>
  );
};


