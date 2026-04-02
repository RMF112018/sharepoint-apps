import { SPFI, spfi, SPFx } from "@pnp/sp";
import type { BaseComponentContext } from "@microsoft/sp-component-base";
// Augment SPFI with web, lists, items, batching, and site-users APIs
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import "@pnp/sp/site-users/web";

export interface SharePointClientFactoryOptions {
  spfxContext: BaseComponentContext;
  baseUrl?: string;
}

export function createSharePointClient(options: SharePointClientFactoryOptions): SPFI {
  const { spfxContext, baseUrl } = options;
  const sp = spfi(baseUrl).using(SPFx(spfxContext));
  return sp;
}

