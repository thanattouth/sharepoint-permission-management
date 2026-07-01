import { managedSiteListEnabled } from "../../app-config";
import { GraphRequestClient, type TokenProvider } from "../../graph-request";
import { LocalManagedSiteStore, type ManagedSiteStore } from "./managed-site-store";
import { SharePointListManagedSiteStore } from "./sharepoint-list-managed-site-store";

export function createManagedSiteStore(getToken: TokenProvider): ManagedSiteStore {
  if (!managedSiteListEnabled) {
    return new LocalManagedSiteStore();
  }

  return new SharePointListManagedSiteStore(new GraphRequestClient(getToken));
}
