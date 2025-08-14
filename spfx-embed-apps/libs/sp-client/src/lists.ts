import type { IItems, IItemAddResult, IItemUpdateResult } from "@pnp/sp/items";
import type { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export interface CreateListItemParams<TFields extends object> {
  listTitle: string;
  fields: TFields;
}

export interface UpdateListItemParams<TFields extends object> {
  listTitle: string;
  id: number;
  fields: Partial<TFields>;
}

export interface GetListItemsParams {
  listTitle: string;
  select?: string[];
  expand?: string[];
  filter?: string;
  orderBy?: { field: string; ascending?: boolean }[];
  top?: number;
}

export async function ensureListExists(sp: SPFI, listTitle: string): Promise<boolean> {
  try {
    await sp.web.lists.getByTitle(listTitle)();
    return true;
  } catch (err) {
    return false;
  }
}

export async function createListItem<TFields extends object>(
  sp: SPFI,
  params: CreateListItemParams<TFields>
): Promise<IItemAddResult> {
  const { listTitle, fields } = params;
  const list = sp.web.lists.getByTitle(listTitle);
  return list.items.add(fields as any);
}

export async function updateListItem<TFields extends object>(
  sp: SPFI,
  params: UpdateListItemParams<TFields>
): Promise<IItemUpdateResult> {
  const { listTitle, id, fields } = params;
  const list = sp.web.lists.getByTitle(listTitle);
  return list.items.getById(id).update(fields as any);
}

export async function deleteListItem(sp: SPFI, listTitle: string, id: number): Promise<void> {
  const list = sp.web.lists.getByTitle(listTitle);
  await list.items.getById(id).delete();
}

export async function getListItems(sp: SPFI, params: GetListItemsParams) {
  const { listTitle, select, expand, filter, orderBy, top } = params;
  let items: IItems = sp.web.lists.getByTitle(listTitle).items;

  if (select?.length) items = items.select(...select);
  if (expand?.length) items = items.expand(...expand);
  if (filter) items = items.filter(filter);
  if (orderBy?.length) {
    for (const ob of orderBy) {
      items = items.orderBy(ob.field, ob.ascending !== false);
    }
  }
  if (top) items = items.top(top);

  return items();
}

