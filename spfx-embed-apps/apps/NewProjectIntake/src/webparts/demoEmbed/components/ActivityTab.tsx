import * as React from 'react';
import type { BaseComponentContext } from '@microsoft/sp-component-base';
import { Stack, Text } from '@fluentui/react';
import { getSyncPercentiles, getRecentFailures } from '../../../utils/telemetry';
import { createSharePointClient } from '@hbi/sp-client';
import mappingRegistry from '../../../config/fieldMappingRegistry.json';

export interface ActivityTabProps {
  context: BaseComponentContext;
  projectId: string;
}

export const ActivityTab: React.FC<ActivityTabProps> = ({ context, projectId }) => {
  const [rows, setRows] = React.useState<any[]>([]);
  const [p50, setP50] = React.useState<number | undefined>();
  const [p95, setP95] = React.useState<number | undefined>();
  const [failures, setFailures] = React.useState<any[]>([]);
  React.useEffect(() => {
    void (async () => {
      try {
        const sp = createSharePointClient({ spfxContext: context });
        const auditTitle = (mappingRegistry as any).lists?.auditListTitle || 'Project Audit';
        const items = await sp.web.lists.getByTitle(auditTitle).items.filter(`ProjectId eq '${projectId}'`).orderBy('Id', false).top(50)();
        setRows(items);
      } catch {}
    })();
    const p = getSyncPercentiles();
    setP50(p.p50);
    setP95(p.p95);
    setFailures(getRecentFailures());
  }, [context, projectId]);

  return (
    <Stack tokens={{ childrenGap: 8 }}>
      {(p50 || p95) && (
        <Text>Sync latency p50 {Math.round(p50 || 0)}ms • p95 {Math.round(p95 || 0)}ms</Text>
      )}
      {failures.length > 0 && (
        <Stack tokens={{ childrenGap: 4 }}>
          <Text>Recent failures (client):</Text>
          {failures.slice(-5).map((f, idx) => (
            <Text key={idx} variant="small">{f.at} • {f.status || ''} • {f.reason || ''}</Text>
          ))}
        </Stack>
      )}
      {rows.map((r) => (
        <Stack key={r.Id} tokens={{ childrenGap: 4 }}>
          <Text>{r.Timestamp || r.Created} • {r.FieldKeys}</Text>
          <Text variant="small">By {r.SourceUser || 'system'} SyncId {r.SyncId || ''}</Text>
        </Stack>
      ))}
      {rows.length === 0 && <Text>No recent activity.</Text>}
    </Stack>
  );
};


