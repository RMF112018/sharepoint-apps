import * as React from 'react';
import { Toggle, TextField, Text } from '@fluentui/react';
import { PeoplePickerField } from './PeoplePickerField';
import { DateTimePickerField } from './DateTimePickerField';
import type { BaseComponentContext } from '@microsoft/sp-component-base';

export interface FinalDeliverablesTableProps {
  context: BaseComponentContext;
  values: Record<string, any>;
  setValue: (key: string, value: any) => void;
}

interface FinalDeliverablesRow {
  key: string;
  label: string;
  yesNoKey: string;
  responsibleKey: string;
  deadlineKey: string;
  notesKey: string;
}

export const FinalDeliverablesTable: React.FC<FinalDeliverablesTableProps> = ({ context, values, setValue }) => {
  const rows: FinalDeliverablesRow[] = [
    {
      key: 'frontCover',
      label: 'Front Cover',
      yesNoKey: 'frontCover',
      responsibleKey: 'frontCoverResponsible',
      deadlineKey: 'frontCoverDeadline',
      notesKey: 'frontCoverNotes'
    },
    {
      key: 'executiveSummary',
      label: 'Executive Summary',
      yesNoKey: 'executiveSummary',
      responsibleKey: 'executiveSummaryResponsible',
      deadlineKey: 'executiveSummaryDeadline',
      notesKey: 'executiveSummaryNotes'
    },
    {
      key: 'costSummary',
      label: 'Cost Summary',
      yesNoKey: 'costSummary',
      responsibleKey: 'costSummaryResponsible',
      deadlineKey: 'costSummaryDeadline',
      notesKey: 'costSummaryNotes'
    },
    {
      key: 'detailedGcGfcBreakdown',
      label: 'Detailed GC/GFC Breakdown (optional)',
      yesNoKey: 'detailedGcGfcBreakdown',
      responsibleKey: 'detailedGcGfcBreakdownResponsible',
      deadlineKey: 'detailedGcGfcBreakdownDeadline',
      notesKey: 'detailedGcGfcBreakdownNotes'
    },
    {
      key: 'detailedCowBreakdown',
      label: 'Detailed COW Breakdown (optional)',
      yesNoKey: 'detailedCowBreakdown',
      responsibleKey: 'detailedCowBreakdownResponsible',
      deadlineKey: 'detailedCowBreakdownDeadline',
      notesKey: 'detailedCowBreakdownNotes'
    },
    {
      key: 'listOfAllowances',
      label: 'List of Allowances',
      yesNoKey: 'listOfAllowances',
      responsibleKey: 'listOfAllowancesResponsible',
      deadlineKey: 'listOfAllowancesDeadline',
      notesKey: 'listOfAllowancesNotes'
    },
    {
      key: 'clarificationsAndAssumptions',
      label: 'Clarifications and Assumptions',
      yesNoKey: 'clarificationsAndAssumptions',
      responsibleKey: 'clarificationsAndAssumptionsResponsible',
      deadlineKey: 'clarificationsAndAssumptionsDeadline',
      notesKey: 'clarificationsAndAssumptionsNotes'
    },
    {
      key: 'valueAnalysisLog',
      label: 'Value Analysis log',
      yesNoKey: 'valueAnalysisLog',
      responsibleKey: 'valueAnalysisLogResponsible',
      deadlineKey: 'valueAnalysisLogDeadline',
      notesKey: 'valueAnalysisLogNotes'
    },
    {
      key: 'schedule',
      label: 'Schedule',
      yesNoKey: 'schedule',
      responsibleKey: 'scheduleResponsible',
      deadlineKey: 'scheduleDeadline',
      notesKey: 'scheduleNotes'
    },
    {
      key: 'logisticsPlan',
      label: 'Logistics Plan',
      yesNoKey: 'logisticsPlan',
      responsibleKey: 'logisticsPlanResponsible',
      deadlineKey: 'logisticsPlanDeadline',
      notesKey: 'logisticsPlanNotes'
    },
    {
      key: 'listOfDocuments',
      label: 'List of Documents',
      yesNoKey: 'listOfDocuments',
      responsibleKey: 'listOfDocumentsResponsible',
      deadlineKey: 'listOfDocumentsDeadline',
      notesKey: 'listOfDocumentsNotes'
    },
    {
      key: 'teamOrganizationChartAndResumes',
      label: 'Team Organization Chart and Resumes',
      yesNoKey: 'teamOrganizationChartAndResumes',
      responsibleKey: 'teamOrganizationChartAndResumesResponsible',
      deadlineKey: 'teamOrganizationChartAndResumesDeadline',
      notesKey: 'teamOrganizationChartAndResumesNotes'
    },
    {
      key: 'previousExperience',
      label: 'Previous Experience',
      yesNoKey: 'previousExperience',
      responsibleKey: 'previousExperienceResponsible',
      deadlineKey: 'previousExperienceDeadline',
      notesKey: 'previousExperienceNotes'
    },
    {
      key: 'rmpProposalRequired',
      label: 'RMP Proposal Required',
      yesNoKey: 'rmpProposalRequired',
      responsibleKey: 'rmpProposalRequiredResponsible',
      deadlineKey: 'rmpProposalRequiredDeadline',
      notesKey: 'rmpProposalRequiredNotes'
    },
    {
      key: 'byWhoList',
      label: 'By Who List',
      yesNoKey: 'byWhoList',
      responsibleKey: 'byWhoListResponsible',
      deadlineKey: 'byWhoListDeadline',
      notesKey: 'byWhoListNotes'
    },
    {
      key: 'backCover',
      label: 'Back Cover',
      yesNoKey: 'backCover',
      responsibleKey: 'backCoverResponsible',
      deadlineKey: 'backCoverDeadline',
      notesKey: 'backCoverNotes'
    }
  ];

  // Using CSS classes instead of inline styles for better responsiveness

  return (
    <div style={{ width: '100%' }}>
      <Text variant="large" style={{ fontWeight: 600, marginBottom: '16px', display: 'block' }}>
        Final Deliverables (STANDARD SECTIONS)
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
          </tr>
        </thead>
        <tbody>
          {rows.map((row) => (
            <tr key={row.key}>
              <td>
                <Text variant="medium">{row.label}</Text>
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
            </tr>
          ))}
        </tbody>
      </table>
      </div>
    </div>
  );
};
