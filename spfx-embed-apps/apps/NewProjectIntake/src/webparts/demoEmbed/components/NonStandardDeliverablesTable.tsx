import * as React from 'react';
import { Toggle, TextField, Text, Stack, IconButton, PrimaryButton } from '@fluentui/react';
import { PeoplePickerField } from './PeoplePickerField';
import { DateTimePickerField } from './DateTimePickerField';
import type { BaseComponentContext } from '@microsoft/sp-component-base';

export interface NonStandardDeliverablesTableProps {
  context: BaseComponentContext;
  values: Record<string, any>;
  setValue: (key: string, value: any) => void;
}

interface NonStandardDeliverablesRow {
  key: string;
  label: string;
  yesNoKey: string;
  responsibleKey: string;
  deadlineKey: string;
  notesKey: string;
  isCustom?: boolean;
  isEditable?: boolean;
}

interface CustomField {
  id: string;
  label: string;
}

export const NonStandardDeliverablesTable: React.FC<NonStandardDeliverablesTableProps> = ({ context, values, setValue }) => {
  const [customFields, setCustomFields] = React.useState<CustomField[]>([]);
  const [nextCustomId, setNextCustomId] = React.useState(1);

  const baseRows: NonStandardDeliverablesRow[] = [
    {
      key: 'financials',
      label: 'Financials',
      yesNoKey: 'financials',
      responsibleKey: 'financialsResponsible',
      deadlineKey: 'financialsDeadline',
      notesKey: 'financialsNotes'
    },
    {
      key: 'gcLicense',
      label: 'GC License',
      yesNoKey: 'gcLicense',
      responsibleKey: 'gcLicenseResponsible',
      deadlineKey: 'gcLicenseDeadline',
      notesKey: 'gcLicenseNotes'
    },
    {
      key: 'bim',
      label: 'BIM',
      yesNoKey: 'bim',
      responsibleKey: 'bimResponsible',
      deadlineKey: 'bimDeadline',
      notesKey: 'bimNotes'
    },
    {
      key: 'contract',
      label: 'Contract',
      yesNoKey: 'contract',
      responsibleKey: 'contractResponsible',
      deadlineKey: 'contractDeadline',
      notesKey: 'contractNotes'
    },
    {
      key: 'bidBond',
      label: 'Bid Bond',
      yesNoKey: 'bidBond',
      responsibleKey: 'bidBondResponsible',
      deadlineKey: 'bidBondDeadline',
      notesKey: 'bidBondNotes'
    },
    {
      key: 'businessTerms',
      label: 'Business Terms',
      yesNoKey: 'businessTerms',
      responsibleKey: 'businessTermsResponsible',
      deadlineKey: 'businessTermsDeadline',
      notesKey: 'businessTermsNotes'
    }
  ];

  // Add placeholders for "other" fields from the image
  const otherRows: NonStandardDeliverablesRow[] = [
    {
      key: 'other1',
      label: 'other',
      yesNoKey: 'other1',
      responsibleKey: 'other1Responsible',
      deadlineKey: 'other1Deadline',
      notesKey: 'other1Notes',
      isCustom: true,
      isEditable: true
    },
    {
      key: 'other2',
      label: 'other',
      yesNoKey: 'other2',
      responsibleKey: 'other2Responsible',
      deadlineKey: 'other2Deadline',
      notesKey: 'other2Notes',
      isCustom: true,
      isEditable: true
    },
    {
      key: 'other3',
      label: 'other',
      yesNoKey: 'other3',
      responsibleKey: 'other3Responsible',
      deadlineKey: 'other3Deadline',
      notesKey: 'other3Notes',
      isCustom: true,
      isEditable: true
    }
  ];

  const allRows = [...baseRows, ...otherRows];

  // Add custom fields dynamically
  const customRows: NonStandardDeliverablesRow[] = customFields.map(field => ({
    key: `custom_${field.id}`,
    label: field.label,
    yesNoKey: `custom_${field.id}`,
    responsibleKey: `custom_${field.id}_responsible`,
    deadlineKey: `custom_${field.id}_deadline`,
    notesKey: `custom_${field.id}_notes`,
    isCustom: true,
    isEditable: true
  }));

  const finalRows = [...allRows, ...customRows];

  const addCustomField = () => {
    const newField: CustomField = {
      id: `${nextCustomId}`,
      label: 'Custom Field'
    };
    setCustomFields([...customFields, newField]);
    setNextCustomId(nextCustomId + 1);
  };

  const updateCustomFieldLabel = (fieldId: string, newLabel: string) => {
    setCustomFields(customFields.map((field: CustomField) => 
      field.id === fieldId ? { ...field, label: newLabel } : field
    ));
  };

  const removeCustomField = (fieldId: string) => {
    setCustomFields(customFields.filter((field: CustomField) => field.id !== fieldId));
    // Clear related values
    const prefix = `custom_${fieldId}`;
    setValue(prefix, undefined);
    setValue(`${prefix}_responsible`, undefined);
    setValue(`${prefix}_deadline`, undefined);
    setValue(`${prefix}_notes`, undefined);
  };

  const updateOtherLabel = (rowKey: string, newLabel: string) => {
    // This would ideally be stored in form state, but for simplicity we'll use a direct approach
    const labelKey = `${rowKey}_label`;
    setValue(labelKey, newLabel);
  };

  const getOtherLabel = (rowKey: string): string => {
    const labelKey = `${rowKey}_label`;
    return values[labelKey] || 'other';
  };

  // Using CSS classes instead of inline styles for better responsiveness

  return (
    <div style={{ width: '100%' }}>
      <Text variant="large" style={{ fontWeight: 600, marginBottom: '16px', display: 'block' }}>
        Final Deliverables (NON-STANDARD SECTIONS)
      </Text>
      <div style={{ overflowX: 'auto' }}>
        <table className="clean-table">
        <thead>
          <tr>
            <th>Task Name</th>
            <th>Yes/No</th>
            <th>Responsible</th>
            <th>Deadline</th>
            <th>Notes</th>
            <th></th>
          </tr>
        </thead>
        <tbody>
          {finalRows.map((row) => (
            <tr key={row.key}>
              <td>
                {row.isEditable ? (
                  <TextField
                    value={row.key.indexOf('custom_') === 0 ? 
                      customFields.filter((f: CustomField) => f.id === row.key.replace('custom_', ''))[0]?.label || '' :
                      getOtherLabel(row.key)
                    }
                    onChange={(_, newValue) => {
                      if (row.key.indexOf('custom_') === 0) {
                        const fieldId = row.key.replace('custom_', '');
                        updateCustomFieldLabel(fieldId, newValue || '');
                      } else {
                        updateOtherLabel(row.key, newValue || '');
                      }
                    }}
                    placeholder="Enter field name"
                    borderless
                    styles={{ root: { width: '100%' } }}
                  />
                ) : (
                  <Text variant="medium">{row.label}</Text>
                )}
              </td>
              <td>
                <Toggle
                  checked={Boolean(values[row.yesNoKey])}
                  onChange={(_, checked) => setValue(row.yesNoKey, !!checked)}
                  onText="Yes"
                  offText="No"
                  styles={{
                    root: { margin: 0 },
                    pill: { width: '40px', height: '20px' },
                    thumb: { width: '16px', height: '16px' }
                  }}
                />
              </td>
              <td>
                <PeoplePickerField
                  context={context}
                  label=""
                  multiSelect={true}
                  value={(values[row.responsibleKey] as any[]) || []}
                  onChange={(people) => setValue(row.responsibleKey, people)}
                  required={false}
                />
              </td>
              <td>
                <DateTimePickerField
                  value={values[row.deadlineKey]}
                  onChange={(date) => setValue(row.deadlineKey, date)}
                  placeholder="Select deadline"
                />
              </td>
              <td>
                <TextField
                  value={String(values[row.notesKey] || '')}
                  onChange={(_, v) => setValue(row.notesKey, v || '')}
                  multiline
                  rows={2}
                  borderless
                />
              </td>
              <td style={{ textAlign: 'center' }}>
                {row.key.indexOf('custom_') === 0 && (
                  <IconButton
                    iconProps={{ iconName: 'Delete' }}
                    ariaLabel="Remove custom field"
                    onClick={() => removeCustomField(row.key.replace('custom_', ''))}
                    styles={{ root: { minWidth: '24px', width: '24px', height: '24px' } }}
                  />
                )}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      </div>
      
      <Stack horizontal tokens={{ childrenGap: 12 }} style={{ marginTop: '16px' }}>
        <PrimaryButton
          text="Add Custom Field"
          iconProps={{ iconName: 'Add' }}
          onClick={addCustomField}
          styles={{ root: { minWidth: '140px' } }}
        />
      </Stack>
    </div>
  );
};
