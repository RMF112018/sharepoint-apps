import type { SPFI } from "@pnp/sp";
import "@pnp/sp/batching";

export interface BatchOperation<T> {
  key: string;
  invoke: (sp: SPFI) => Promise<T>;
}

export interface BatchResult<T> {
  key: string;
  status: "fulfilled" | "rejected";
  value?: T;
  reason?: unknown;
}

export async function executeBatch<T>(sp: SPFI, operations: BatchOperation<T>[]): Promise<BatchResult<T>[]> {
  const [batchedSP, execute] = sp.batched();

  const promises = operations.map((op) =>
    op
      .invoke(batchedSP)
      .then((value) => ({ key: op.key, status: "fulfilled" as const, value }))
      .catch((reason) => ({ key: op.key, status: "rejected" as const, reason }))
  );

  await execute();
  return Promise.all(promises);
}

