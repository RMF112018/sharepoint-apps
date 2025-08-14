// A file is required to be in the root of the /src directory by the TypeScript compiler
// Ensure PnPjs side-effect modules are included for SPFI augmentation
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/batching';
import '@pnp/sp/site-users/web';
