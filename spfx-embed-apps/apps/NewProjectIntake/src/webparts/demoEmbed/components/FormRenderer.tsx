import * as React from 'react';
import type { BaseComponentContext } from '@microsoft/sp-component-base';
import { Stack, Text, Separator, TextField, Dropdown, type IDropdownOption, Toggle } from '@fluentui/react';
import { PeoplePickerField } from './PeoplePickerField';
import { ManagingInfoChecklist } from './ManagingInfoChecklist';
import { KeyDatesMilestones } from './KeyDatesMilestones';
import { FinalDeliverables } from './FinalDeliverables';
import { ManagingInfoTable } from './ManagingInfoTable';
import { KeyDatesTable } from './KeyDatesTable';
import { FinalDeliverablesTable } from './FinalDeliverablesTable';
import { NonStandardDeliverablesTable } from './NonStandardDeliverablesTable';

type FieldType = 'Text' | 'Number' | 'DateTime' | 'Choice' | 'Boolean' | 'Multiline' | 'Person';

export interface FieldDef {
  field_key: string;
  label: string;
  group: string;
  suggested_type: FieldType;
  required?: boolean;
  help_text?: string;
  choice_options?: string[];
  default?: any;
  multi_select?: boolean;
}

export interface FormRendererProps {
  context: BaseComponentContext;
  catalog: FieldDef[];
  initialValues?: Record<string, FieldValue>;
  onValuesChange?: (values: Record<string, FieldValue>) => void;
  currentStep?: number;
  onProjectNameChange?: (projectName: string) => void;
}

type FieldValue = string | number | boolean | unknown[] | null;

export const FormRenderer: React.FC<FormRendererProps> = ({ context, catalog, initialValues, onValuesChange, currentStep = 1, onProjectNameChange }) => {
  const [values, setValues] = React.useState<Record<string, FieldValue>>(initialValues || {});
  const [errors, setErrors] = React.useState<Record<string, string>>({});

  const groups: Array<{ name: string; fields: FieldDef[] }> = React.useMemo(() => {
    const map = new Map<string, FieldDef[]>();
    for (const f of catalog) {
      const arr = map.get(f.group) || [];
      arr.push(f);
      map.set(f.group, arr);
    }
    const list: Array<{ name: string; fields: FieldDef[] }> = [];
    map.forEach((v, k) => list.push({ name: k, fields: v }));
    return list;
  }, [catalog]);

  const setValue = (key: string, v: FieldValue): void => {
    setValues(s => ({ ...s, [key]: v }));
    
    // If this is the project name field, notify parent
    if (key === 'projectName' && onProjectNameChange) {
      onProjectNameChange(String(v || ''));
    }
  };

  React.useEffect(() => {
    if (initialValues) {
      setValues(initialValues);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [initialValues]);

  React.useEffect(() => {
    if (onValuesChange) onValuesChange(values);
  }, [values, onValuesChange]);

  const getValueString = (key: string): string => String(values[key] ?? '');

  const isEmpty = (v: FieldValue): boolean => v === undefined || v === null || v === '' || (Array.isArray(v) && v.length === 0);

  const findFieldByLabel = (pattern: RegExp): FieldDef | undefined => {
    for (const f of catalog) {
      if (pattern.test(f.label.toLowerCase())) return f;
    }
    return undefined;
  };

  const validateField = (f: FieldDef, allValues: Record<string, FieldValue>): string => {
    const v = allValues[f.field_key];
    // Required
    if (f.required && isEmpty(v)) return 'This field is required.';
    
    // Job Number format validation (##-###-##)
    if (f.field_key === 'jobNumber' && !isEmpty(v)) {
      const jobNumberPattern = /^\d{2}-\d{3}-\d{2}$/;
      if (!jobNumberPattern.test(String(v))) {
        return 'Job Number must be in format ##-###-## (e.g., 12-345-67)';
      }
    }
    
    // Conditional: if Delivered Via = Hand Delivery -> require Copies
    const deliveredVia = findFieldByLabel(/delivered\s+via/);
    let copiesField: FieldDef | undefined;
    for (const cf of catalog) {
      if (/copies/.test(cf.label.toLowerCase())) { copiesField = cf; break; }
    }
    if (copiesField && f.field_key === copiesField.field_key) {
      const viaVal = (deliveredVia && (allValues[deliveredVia.field_key] as string)) || '';
      if (typeof viaVal === 'string' && viaVal.toLowerCase() === 'hand delivery' && isEmpty(v)) return 'Required when delivered via Hand Delivery.';
    }
    // Numbers non-negative
    if (f.suggested_type === 'Number') {
      const num = typeof v === 'number' ? v : parseFloat(String(v ?? ''));
      if (isNaN(num as any)) return '';
      if (num < 0) return 'Value must be zero or greater.';
    }
    // Date/time validity
    if (f.suggested_type === 'DateTime') {
      const dt = new Date(String(v ?? ''));
      if (String(v ?? '') !== '' && String(dt) === 'Invalid Date') return 'Enter a valid date/time.';
    }
    return '';
  };

  React.useEffect(() => {
    // Apply business-required flags based on labels - only for Panel 1 (Project Info)
    const reqPatterns = [
      /proposal\s*due/i,
      /type\s*of\s*proposal/i,
      /rfi\s*format/i
    ];
    for (const f of catalog) {
      // Only apply required flags to Project Info fields
      if (f.group === 'Project Info' && reqPatterns.some(r => r.test(f.label))) {
        f.required = true;
      }
      // Ensure fields in panels 2-5 are not required, but keep Project Details (panel 6) as defined in catalog
      if (f.group !== 'Project Info' && f.group !== 'Project Details') {
        f.required = false;
      }
      if (/job\s*number/i.test(f.label)) f.suggested_type = 'Text';
    }
  }, [catalog]);

  React.useEffect(() => {
    // Recompute errors when values change
    const e: Record<string, string> = {};
    for (const f of catalog) {
      const err = validateField(f, values);
      if (err) e[f.field_key] = err;
    }
    setErrors(e);
  }, [values, catalog]);

  // Step-based rendering
  if (currentStep === 1) {
    // Show Project Info fields only (excluding project name which is handled in parent)
    const projectInfoFields = catalog.filter(f => f.group === 'Project Info');
    return (
      <div className="responsive-form-grid">
        {projectInfoFields.map(f => {
          if (f.suggested_type === 'Boolean') {
            return (
              <Toggle
                key={f.field_key}
                label={f.label}
                checked={Boolean(values[f.field_key])}
                onChange={(_, checked) => setValue(f.field_key, !!checked)}
              />
            );
          }
          if (f.suggested_type === 'Choice') {
            const options: IDropdownOption[] = (f.choice_options || []).map(o => ({ key: o, text: o }));
            const sel = String(values[f.field_key] ?? f.default ?? '');
            return (
              <Dropdown
                key={f.field_key}
                label={f.label}
                options={options}
                selectedKey={sel || undefined}
                onChange={(_, opt) => setValue(f.field_key, opt ? String(opt.key) : '')}
                required={!!f.required}
                aria-describedby={errors[f.field_key] ? `${f.field_key}-error` : undefined}
                errorMessage={errors[f.field_key] || undefined}
              />
            );
          }
          if (f.suggested_type === 'Number') {
            const val = getValueString(f.field_key);
            return (
              <TextField
                key={f.field_key}
                label={f.label}
                type="number"
                min={0}
                value={val}
                onChange={(_, v) => setValue(f.field_key, v ?? '')}
                required={!!f.required}
                aria-describedby={errors[f.field_key] ? `${f.field_key}-error` : undefined}
                onGetErrorMessage={() => errors[f.field_key] || ''}
                validateOnLoad={false}
              />
            );
          }
          if (f.suggested_type === 'DateTime') {
            const val = getValueString(f.field_key);
            return (
              <TextField
                key={f.field_key}
                label={f.label}
                type="datetime-local"
                value={val}
                onChange={(_, v) => setValue(f.field_key, v ?? '')}
                required={!!f.required}
                aria-describedby={errors[f.field_key] ? `${f.field_key}-error` : undefined}
                onGetErrorMessage={() => errors[f.field_key] || ''}
                validateOnLoad={false}
              />
            );
          }
          if (f.suggested_type === 'Person') {
            const multi = f.multi_select === true || /s$/.test(f.label.trim().toLowerCase());
            return (
              <PeoplePickerField
                key={f.field_key}
                context={context}
                label={f.label}
                description={f.help_text}
                required={!!f.required}
                multiSelect={multi}
                value={(values[f.field_key] as unknown[] as any[]) || []}
                onChange={(people) => setValue(f.field_key, people)}
                errorMessage={errors[f.field_key]}
              />
            );
          }
          // Default: Text
          const val = getValueString(f.field_key);
          return (
            <TextField
              key={f.field_key}
              label={f.label}
              value={val}
              onChange={(_, v) => setValue(f.field_key, v ?? '')}
              required={!!f.required}
              description={f.help_text}
              placeholder={f.field_key === 'jobNumber' ? '12-345-67' : undefined}
              aria-describedby={errors[f.field_key] ? `${f.field_key}-error` : undefined}
              onGetErrorMessage={() => errors[f.field_key] || ''}
              validateOnLoad={false}
            />
          );
        })}
      </div>
    );
  }

  // Step 2: Managing Information Table
  if (currentStep === 2) {
    return (
      <Stack tokens={{ childrenGap: 16 }}>
        <ManagingInfoTable 
          context={context}
          values={values}
          setValue={setValue}
        />
      </Stack>
    );
  }

  // Step 3: Key Dates Table
  if (currentStep === 3) {
    return (
      <Stack tokens={{ childrenGap: 16 }}>
        <KeyDatesTable 
          context={context}
          values={values}
          setValue={setValue}
        />
      </Stack>
    );
  }

  // Step 4: Final Deliverables Table
  if (currentStep === 4) {
    return (
      <Stack tokens={{ childrenGap: 16 }}>
        <FinalDeliverablesTable 
          context={context}
          values={values}
          setValue={setValue}
        />
      </Stack>
    );
  }

  // Step 5: Non-Standard Deliverables Table
  if (currentStep === 5) {
    return (
      <Stack tokens={{ childrenGap: 16 }}>
        <NonStandardDeliverablesTable 
          context={context}
          values={values}
          setValue={setValue}
        />
      </Stack>
    );
  }

  // Step 6: Project Details - Render as standard form fields in responsive grid
  if (currentStep === 6) {
    const projectDetailsFields = catalog.filter(f => f.group === 'Project Details');
    return (
      <div className="responsive-form-grid">
        {projectDetailsFields.map(f => {
          if (f.suggested_type === 'Boolean') {
            return (
              <Toggle
                key={f.field_key}
                label={f.label}
                checked={Boolean(values[f.field_key])}
                onChange={(_, checked) => setValue(f.field_key, !!checked)}
                onText="Yes"
                offText="No"
              />
            );
          }
          if (f.suggested_type === 'Choice') {
            const options: IDropdownOption[] = (f.choice_options || []).map(o => ({ key: o, text: o }));
            const sel = String(values[f.field_key] ?? f.default ?? '');
            return (
              <Dropdown
                key={f.field_key}
                label={f.label}
                options={options}
                selectedKey={sel || undefined}
                onChange={(_, opt) => setValue(f.field_key, opt ? String(opt.key) : '')}
                required={!!f.required}
                aria-describedby={errors[f.field_key] ? `${f.field_key}-error` : undefined}
                errorMessage={errors[f.field_key] || undefined}
              />
            );
          }
          if (f.suggested_type === 'Person') {
            const multi = f.multi_select === true || /s$/.test(f.label.trim().toLowerCase());
            return (
              <PeoplePickerField
                key={f.field_key}
                context={context}
                label={f.label}
                description={f.help_text}
                required={!!f.required}
                multiSelect={multi}
                value={(values[f.field_key] as unknown[] as any[]) || []}
                onChange={(people) => setValue(f.field_key, people)}
                errorMessage={errors[f.field_key]}
              />
            );
          }
          // Default: Text
          const val = getValueString(f.field_key);
          return (
            <TextField
              key={f.field_key}
              label={f.label}
              value={val}
              onChange={(_, v) => setValue(f.field_key, v ?? '')}
              required={!!f.required}
              description={f.help_text}
              placeholder={f.field_key === 'jobNumber' ? '12-345-67' : undefined}
              aria-describedby={errors[f.field_key] ? `${f.field_key}-error` : undefined}
              onGetErrorMessage={() => errors[f.field_key] || ''}
              validateOnLoad={false}
            />
          );
        })}
      </div>
    );
  }

  // Original rendering for other steps
  return (
    <Stack tokens={{ childrenGap: 24 }}>
      {groups.map(({ name: groupName, fields }) => (
        <Stack key={groupName} tokens={{ childrenGap: 12 }}>
          <Separator alignContent="start"><Text variant="large">{groupName}</Text></Separator>
          <Stack tokens={{ childrenGap: 16 }}>
            {groupName === 'Managing Information' ? (
              <ManagingInfoChecklist
                context={context}
                tasks={fields}
                values={values}
                setValue={setValue}
              />
            ) : groupName === 'Key Dates' ? (
              <KeyDatesMilestones
                context={context}
                tasks={fields}
                catalog={catalog}
                values={values}
                setValue={setValue}
              />
            ) : groupName.indexOf('Final Deliverables') === 0 ? (
              <FinalDeliverables
                context={context}
                standard={fields.filter(f => /standard/i.test(f.group))}
                nonStandard={fields.filter(f => /non-\s*standard/i.test(f.group))}
                values={values}
                setValue={setValue}
              />
            ) : (
              fields.map((f: FieldDef) => {
              if (f.suggested_type === 'Person') {
                const multi = f.multi_select === true || /s$/.test(f.label.trim().toLowerCase());
                return (
                  <PeoplePickerField
                    key={f.field_key}
                    context={context}
                    label={f.label}
                    description={f.help_text}
                    required={!!f.required}
                    multiSelect={multi}
                    value={(values[f.field_key] as unknown[] as any[]) || []}
                    onChange={(people) => setValue(f.field_key, people)}
                    errorMessage={errors[f.field_key]}
                  />
                );
              }
              if (f.suggested_type === 'Boolean') {
                const boolVal = Boolean(values[f.field_key]);
                return (
                  <Toggle
                    key={f.field_key}
                    label={f.label}
                    checked={boolVal}
                    onChange={(_, checked) => setValue(f.field_key, !!checked)}
                  />
                );
              }
              if (f.suggested_type === 'Choice') {
                const options: IDropdownOption[] = (f.choice_options || []).map(o => ({ key: o, text: o }));
                const sel = String(values[f.field_key] ?? f.default ?? '');
                return (
                  <Dropdown
                    key={f.field_key}
                    label={f.label}
                    options={options}
                    selectedKey={sel || undefined}
                    onChange={(_, opt) => setValue(f.field_key, opt ? String(opt.key) : '')}
                    required={!!f.required}
                    aria-describedby={errors[f.field_key] ? `${f.field_key}-error` : undefined}
                    errorMessage={errors[f.field_key] || undefined}
                  />
                );
              }
              if (f.suggested_type === 'Number') {
                const val = getValueString(f.field_key);
                return (
                  <TextField
                    key={f.field_key}
                    label={f.label}
                    type="number"
                    min={0}
                    value={val}
                    onChange={(_, v) => setValue(f.field_key, v ?? '')}
                    required={!!f.required}
                    aria-describedby={errors[f.field_key] ? `${f.field_key}-error` : undefined}
                    onGetErrorMessage={() => errors[f.field_key] || ''}
                    validateOnLoad={false}
                  />
                );
              }
              if (f.suggested_type === 'DateTime') {
                const val = getValueString(f.field_key);
                return (
                  <TextField
                    key={f.field_key}
                    label={f.label}
                    type="datetime-local"
                    value={val}
                    onChange={(_, v) => setValue(f.field_key, v ?? '')}
                    required={!!f.required}
                    aria-describedby={errors[f.field_key] ? `${f.field_key}-error` : undefined}
                    onGetErrorMessage={() => errors[f.field_key] || ''}
                    validateOnLoad={false}
                  />
                );
              }
              if (f.suggested_type === 'Multiline') {
                const val = getValueString(f.field_key);
                const needsCounter = /clarifications|value\s*analysis|other/i.test(f.label);
                const maxLen = needsCounter ? 1000 : undefined;
                const desc = needsCounter ? `${val.length}/${maxLen}` : (f.help_text || undefined);
                const over = maxLen ? val.length > maxLen : false;
                return (
                  <TextField
                    key={f.field_key}
                    label={f.label}
                    multiline
                    value={val}
                    onChange={(_, v) => setValue(f.field_key, v ?? '')}
                    required={!!f.required}
                    description={desc}
                    aria-describedby={errors[f.field_key] ? `${f.field_key}-error` : undefined}
                    onGetErrorMessage={() => over ? `Please reduce to ${maxLen} characters or fewer.` : (errors[f.field_key] || '')}
                    validateOnLoad={false}
                  />
                );
              }
              // Default: Text
              const val = getValueString(f.field_key);
              return (
                <TextField
                  key={f.field_key}
                  label={f.label}
                  value={val}
                  onChange={(_, v) => setValue(f.field_key, v ?? '')}
                  required={!!f.required}
                  aria-describedby={errors[f.field_key] ? `${f.field_key}-error` : undefined}
                  onGetErrorMessage={() => errors[f.field_key] || ''}
                  validateOnLoad={false}
                />
              );
              })
            )}
          </Stack>
        </Stack>
      ))}
    </Stack>
  );
};


