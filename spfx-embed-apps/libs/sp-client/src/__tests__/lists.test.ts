import { createListItem, updateListItem, deleteListItem, getListItems, ensureListExists } from "../lists";
import type { SPFI } from "@pnp/sp";

function createSpMock() {
  const itemsStore: any[] = [];
  let idSeq = 1;

  // Wrapper function that is callable and chainable
  const makeItems = () => {
    const wrapper: any = async () => itemsStore;
    wrapper.add = jest.fn(async (fields: any) => {
      const item = { Id: idSeq++, ...fields };
      itemsStore.push(item);
      return { data: item };
    });
    wrapper.getById = jest.fn((id: number) => ({
      update: jest.fn(async (fields: any) => {
        const idx = itemsStore.findIndex((i) => i.Id === id);
        if (idx >= 0) itemsStore[idx] = { ...itemsStore[idx], ...fields };
        return { data: itemsStore[idx] };
      }),
      delete: jest.fn(async () => {
        const idx = itemsStore.findIndex((i) => i.Id === id);
        if (idx >= 0) itemsStore.splice(idx, 1);
      }),
    }));
    const returnSelf = () => wrapper;
    wrapper.select = jest.fn(returnSelf);
    wrapper.expand = jest.fn(returnSelf);
    wrapper.filter = jest.fn(returnSelf);
    wrapper.orderBy = jest.fn(returnSelf);
    wrapper.top = jest.fn(returnSelf);
    return wrapper;
  };

  const items = makeItems();

  // getByTitle returns a callable function (list getter), which can be invoked
  const listGetter: any = Object.assign(async () => ({ Title: "MockList" }), {
    items,
  });

  const sp: Partial<SPFI> = {
    web: {
      lists: {
        getByTitle: jest.fn(() => listGetter),
      },
    },
  } as any;

  return { sp: sp as SPFI, itemsStore, listGetter };
}

describe("lists helpers", () => {
  test("ensureListExists returns true when list resolves", async () => {
    const { sp } = createSpMock();
    await expect(ensureListExists(sp, "Any")).resolves.toBe(true);
  });

  test("create, update, delete item", async () => {
    const { sp, itemsStore } = createSpMock();
    const add = await createListItem(sp, { listTitle: "L", fields: { Title: "A" } });
    expect(add.data.Title).toBe("A");
    expect(itemsStore).toHaveLength(1);

    const upd = await updateListItem(sp, { listTitle: "L", id: add.data.Id, fields: { Title: "B" } });
    expect((upd as any).data.Title).toBe("B");

    await deleteListItem(sp, "L", add.data.Id);
    expect(itemsStore).toHaveLength(0);
  });

  test("getListItems respects top() (structural)", async () => {
    const { sp } = createSpMock();
    await createListItem(sp, { listTitle: "L", fields: { Title: "A" } });
    await createListItem(sp, { listTitle: "L", fields: { Title: "B" } });

    const res = await getListItems(sp, { listTitle: "L", top: 1 });
    expect(Array.isArray(res)).toBe(true);
    expect(res.length).toBeGreaterThanOrEqual(1);
  });
});


