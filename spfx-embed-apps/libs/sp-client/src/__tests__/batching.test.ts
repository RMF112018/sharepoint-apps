import { executeBatch } from "../batching";
import type { SPFI } from "@pnp/sp";

describe("batching helper", () => {
  test("executes operations and returns results keyed", async () => {
    const results: Record<string, number> = {};
    const batchedSp: Partial<SPFI> = {} as any;
    const sp: Partial<SPFI> = {
      batched: () => [batchedSp as SPFI, async () => void 0],
    } as any;

    const ops = [
      { key: "a", invoke: async () => (results["a"] = 1) },
      { key: "b", invoke: async () => (results["b"] = 2) },
    ];

    const out = await executeBatch(sp as SPFI, ops);
    expect(out).toHaveLength(2);
    expect(out.map((r) => r.key).sort()).toEqual(["a", "b"]);
    expect(results).toEqual({ a: 1, b: 2 });
  });
});


