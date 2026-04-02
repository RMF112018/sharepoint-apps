import { getCurrentUser, getUserByEmail } from "../users";
import type { SPFI } from "@pnp/sp";

function createUserSpMock() {
  const current = { Id: 1, LoginName: "i:0#.f|membership|john@example.com", Email: "john@example.com", Title: "John" };
  const byId = { Id: 2, LoginName: "i:0#.f|membership|jane@example.com", Email: "jane@example.com", Title: "Jane" };

  const sp: Partial<SPFI> = {
    web: {
      currentUser: { select: jest.fn(() => async () => current) } as any,
      ensureUser: jest.fn(async () => ({ data: { Id: byId.Id } })),
      siteUsers: { getById: jest.fn(() => ({ select: jest.fn(() => async () => byId) })) } as any,
    },
  } as any;
  return { sp: sp as SPFI, current, byId };
}

describe("users helpers", () => {
  test("getCurrentUser maps fields", async () => {
    const { sp, current } = createUserSpMock();
    await expect(getCurrentUser(sp)).resolves.toEqual({
      id: current.Id,
      loginName: current.LoginName,
      email: current.Email,
      displayName: current.Title,
    });
  });

  test("getUserByEmail resolves and maps", async () => {
    const { sp, byId } = createUserSpMock();
    await expect(getUserByEmail(sp, "jane@example.com")).resolves.toEqual({
      id: byId.Id,
      loginName: byId.LoginName,
      email: byId.Email,
      displayName: byId.Title,
    });
  });
});


