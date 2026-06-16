import { GraphRequestClient, type TokenProvider } from "./graph-request";
import type { PermissionRequestStore } from "./permission-request-store";
import { SharePointListPermissionRequestStore } from "./sharepoint-list-permission-request-store";

export function createPermissionRequestStore(getToken: TokenProvider): PermissionRequestStore {
  return new SharePointListPermissionRequestStore(new GraphRequestClient(getToken));
}
