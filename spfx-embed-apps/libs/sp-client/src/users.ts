import type { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";

export interface UserInfo {
  id: number;
  loginName: string;
  email?: string;
  displayName?: string;
}

export async function getCurrentUser(sp: SPFI): Promise<UserInfo> {
  const u = await sp.web.currentUser.select("Id", "LoginName", "Email", "Title")();
  return {
    id: u.Id,
    loginName: u.LoginName,
    email: u.Email,
    displayName: u.Title,
  };
}

export async function getUserByEmail(sp: SPFI, email: string): Promise<UserInfo | null> {
  try {
    const user = await sp.web.ensureUser(email);
    const u = await sp.web.siteUsers.getById(user.data.Id).select("Id", "LoginName", "Email", "Title")();
    return {
      id: u.Id,
      loginName: u.LoginName,
      email: u.Email,
      displayName: u.Title,
    };
  } catch {
    return null;
  }
}

