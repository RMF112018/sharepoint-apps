import * as React from 'react';
import { Toggle, TextField, Text } from '@fluentui/react';
import { PeoplePickerField } from './PeoplePickerField';
import { DateTimePickerField } from './DateTimePickerField';
import type { BaseComponentContext } from '@microsoft/sp-component-base';

export interface ManagingInfoTableProps {
  context: BaseComponentContext;
  values: Record<string, any>;
  setValue: (key: string, value: any) => void;
}

interface ManagingInfoRow {
  key: string;
  label: string;
  yesNoKey: string;
  responsibleKey: string;
  deadlineKey: string;
  notesKey: string;
}

export const ManagingInfoTable: React.FC<ManagingInfoTableProps> = ({ context, values, setValue }) => {
  const rows: ManagingInfoRow[] = [
    {
      key: 'enableSubcontractorBidList',
      label: 'Enable Subcontractor Bid List in BC',
      yesNoKey: 'enableSubcontractorBidListInBc',
      responsibleKey: 'enableSubcontractorBidListInBcResponsible',
      deadlineKey: 'enableSubcontractorBidListInBcDeadline',
      notesKey: 'enableSubcontractorBidListInBcNotes'
    },
    {
      key: 'sendItbToSubcontractors',
      label: 'Send ITB to Subcontractors',
      yesNoKey: 'sendItbToSubcontractors',
      responsibleKey: 'sendItbToSubcontractorsResponsible',
      deadlineKey: 'sendItbToSubcontractorsDeadline',
      notesKey: 'sendItbToSubcontractorsNotes'
    },
    {
      key: 'phoneCallsToImproveSubCoverage',
      label: 'Phone Calls to Improve Sub Coverage',
      yesNoKey: 'phoneCallsToImproveSubCoverage',
      responsibleKey: 'phoneCallsToImproveSubCoverageResponsible',
      deadlineKey: 'phoneCallsToImproveSubCoverageDeadline',
      notesKey: 'phoneCallsToImproveSubCoverageNotes'
    },
    {
      key: 'sendMassMessages',
      label: 'Send Mass Messages in Building Connected to improve Sub coverage',
      yesNoKey: 'sendMassMessagesInBuildingConnected',
      responsibleKey: 'sendMassMessagesInBuildingConnectedResponsible',
      deadlineKey: 'sendMassMessagesInBuildingConnectedDeadline',
      notesKey: 'sendMassMessagesInBuildingConnectedNotes'
    },
    {
      key: 'completeBidPackages',
      label: 'Complete Bid Packages',
      yesNoKey: 'completeBidPackages',
      responsibleKey: 'completeBidPackagesResponsible',
      deadlineKey: 'completeBidPackagesDeadline',
      notesKey: 'completeBidPackagesNotes'
    },
    {
      key: 'rfiManagement',
      label: 'RFI Management (Who is Point Person?)',
      yesNoKey: 'rfiManagement',
      responsibleKey: 'rfiManagementResponsible',
      deadlineKey: 'rfiManagementDeadline',
      notesKey: 'rfiManagementNotes'
    },
    {
      key: 'inviteProjectTeam',
      label: 'Invite Project Team to Procore',
      yesNoKey: 'inviteProjectTeamToProcore',
      responsibleKey: 'inviteProjectTeamToProcoreResponsible',
      deadlineKey: 'inviteProjectTeamToProcoreDeadline',
      notesKey: 'inviteProjectTeamToProcoreNotes'
    }
  ];

  // Using CSS classes instead of inline styles for better responsiveness

  return (
    <div style={{ width: '100%' }}>
      <Text variant="large" style={{ fontWeight: 600, marginBottom: '16px', display: 'block' }}>
        Managing Information
      </Text>
      <div style={{ overflowX: 'auto' }}>
        <table className="clean-table">
        <thead>
          <tr>
            <th>Task Name</th>
            <th>Yes/No</th>
            <th>Responsible</th>
            <th>Deadline/Frequency</th>
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
