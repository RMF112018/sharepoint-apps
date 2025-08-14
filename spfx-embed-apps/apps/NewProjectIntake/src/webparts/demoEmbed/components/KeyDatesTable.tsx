import * as React from 'react';
import { TextField, Text } from '@fluentui/react';
import { PeoplePickerField } from './PeoplePickerField';
import { DateTimePickerField } from './DateTimePickerField';
import type { BaseComponentContext } from '@microsoft/sp-component-base';

export interface KeyDatesTableProps {
  context: BaseComponentContext;
  values: Record<string, any>;
  setValue: (key: string, value: any) => void;
}

interface KeyDatesRow {
  key: string;
  label: string;
  responsibleKey: string;
  deadlineKey: string;
  notesKey: string;
}

export const KeyDatesTable: React.FC<KeyDatesTableProps> = ({ context, values, setValue }) => {
  const rows: KeyDatesRow[] = [
    {
      key: 'itbProposalDue',
      label: 'ITB\'s Proposal Due',
      responsibleKey: 'itbProposalDueResponsible',
      deadlineKey: 'itbProposalDueDeadline',
      notesKey: 'itbProposalDueNotes'
    },
    {
      key: 'subcontractorProposalsDue',
      label: 'Subcontractor Proposals Due',
      responsibleKey: 'subcontractorProposalsDueResponsible',
      deadlineKey: 'subcontractorProposalsDueDeadline',
      notesKey: 'subcontractorProposalsDueNotes'
    },
    {
      key: 'schedulePreSubmissionEstimateReview',
      label: 'Schedule Pre-Submission Estimate Review',
      responsibleKey: 'schedulePreSubmissionEstimateReviewResponsible',
      deadlineKey: 'schedulePreSubmissionEstimateReviewDeadline',
      notesKey: 'schedulePreSubmissionEstimateReviewNotes'
    },
    {
      key: 'scheduleWinStrategyMeeting',
      label: 'Schedule Win Strategy Meeting',
      responsibleKey: 'scheduleWinStrategyMeetingResponsible',
      deadlineKey: 'scheduleWinStrategyMeetingDeadline',
      notesKey: 'scheduleWinStrategyMeetingNotes'
    },
    {
      key: 'scheduleSubcontractorSiteWalkThru',
      label: 'Schedule Subcontractor Site Walk-Thru',
      responsibleKey: 'scheduleSubcontractorSiteWalkThruResponsible',
      deadlineKey: 'scheduleSubcontractorSiteWalkThruDeadline',
      notesKey: 'scheduleSubcontractorSiteWalkThruNotes'
    },
    {
      key: 'scheduleOwnerEstimateReview',
      label: 'Schedule Owner Estimate Review',
      responsibleKey: 'scheduleOwnerEstimateReviewResponsible',
      deadlineKey: 'scheduleOwnerEstimateReviewDeadline',
      notesKey: 'scheduleOwnerEstimateReviewNotes'
    }
  ];

  // Using CSS classes instead of inline styles for better responsiveness

  return (
    <div style={{ width: '100%' }}>
      <Text variant="large" style={{ fontWeight: 600, marginBottom: '16px', display: 'block' }}>
        Estimating Preparation - Key Dates
      </Text>
      <div style={{ overflowX: 'auto' }}>
        <table className="clean-table">
        <thead>
          <tr>
            <th>Task Name</th>
            <th>Deadline</th>
            <th>Responsible</th>
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
                <DateTimePickerField
                  value={values[row.deadlineKey]}
                  onChange={(date) => setValue(row.deadlineKey, date)}
                  placeholder="Select deadline"
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
