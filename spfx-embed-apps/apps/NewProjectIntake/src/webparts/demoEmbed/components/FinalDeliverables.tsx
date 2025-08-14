import * as React from 'react';
import type { BaseComponentContext } from '@microsoft/sp-component-base';
import { Stack, Text, Toggle, TextField, Separator } from '@fluentui/react';
import { PeoplePickerField } from './PeoplePickerField';
import type { FieldDef } from './FormRenderer';

export type FieldValue = string | number | boolean | unknown[] | null;

export interface FinalDeliverablesProps {
  context: BaseComponentContext;
  standard: FieldDef[];
  nonStandard: FieldDef[];
  values: Record<string, FieldValue>;
  setValue: (key: string, v: FieldValue) => void;
}

function k(base: string) {
  return {
    required: `${base}TabRequired`,
    responsible: `${base}Responsible`,
    deadline: `${base}Deadline`,
    notes: `${base}Notes`,
    title: `${base}SectionTitle`
  } as const;
}

function SectionList(props: { title: string; rows: FieldDef[]; context: BaseComponentContext; values: Record<string, FieldValue>; setValue: (key: string, v: FieldValue) => void; }) {
  const { title, rows, context, values, setValue } = props;
  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <Text variant="large">{title}</Text>
      {rows.map((t) => {
        const keys = k(t.field_key);
        const req = Boolean(values[keys.required]);
        const resp = (values[keys.responsible] as any[]) || [];
        const deadline = String(values[keys.deadline] ?? '');
        const notes = String(values[keys.notes] ?? '');
        const isOther = /other/i.test(t.label);
        const titleVal = String(values[keys.title] ?? '');
        return (
          <Stack key={t.field_key} horizontal wrap verticalAlign="end" tokens={{ childrenGap: 12 }}>
            <Stack.Item grow={1} styles={{ root: { minWidth: 220 } }}>
              <Text>{t.label}</Text>
              {isOther && (
                <TextField
                  label="Section Title"
                  value={titleVal}
                  onChange={(_, v) => setValue(keys.title, v ?? '')}
                />
              )}
            </Stack.Item>
            <Stack.Item styles={{ root: { minWidth: 120 } }}>
              <Toggle
                label="Tab required?"
                checked={req}
                onChange={(_, checked) => setValue(keys.required, !!checked)}
              />
            </Stack.Item>
            <Stack.Item grow={1} styles={{ root: { minWidth: 260 } }}>
              <PeoplePickerField
                context={context}
                label="Responsible"
                multiSelect
                value={resp}
                onChange={(people) => setValue(keys.responsible, people)}
              />
            </Stack.Item>
            <Stack.Item grow={1} styles={{ root: { minWidth: 200 } }}>
              <TextField
                label="Deadline"
                type="datetime-local"
                value={deadline}
                onChange={(_, v) => setValue(keys.deadline, v ?? '')}
              />
            </Stack.Item>
            <Stack.Item grow={2} styles={{ root: { minWidth: 220 } }}>
              <TextField
                label="Notes"
                value={notes}
                onChange={(_, v) => setValue(keys.notes, v ?? '')}
              />
            </Stack.Item>
          </Stack>
        );
      })}
    </Stack>
  );
}

export const FinalDeliverables: React.FC<FinalDeliverablesProps> = ({ context, standard, nonStandard, values, setValue }) => {
  const total = standard.length + nonStandard.length;
  let ready = 0;
  for (const t of [...standard, ...nonStandard]) {
    const keys = k(t.field_key);
    const req = Boolean(values[keys.required]);
    const resp = (values[keys.responsible] as any[]) || [];
    if (req && resp.length > 0) ready++;
  }

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
        <Text variant="xLarge">Final Deliverables</Text>
        <Text variant="medium">{ready} of {total} tabs ready</Text>
      </Stack>
      <Separator />
      <SectionList title="Standard" rows={standard} context={context} values={values} setValue={setValue} />
      <Separator />
      <SectionList title="Non-standard" rows={nonStandard} context={context} values={values} setValue={setValue} />
    </Stack>
  );
};


