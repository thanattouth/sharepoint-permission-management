import { auditLogEnabled } from "../../app-config";
import { DisabledAuditStore, type AuditStore } from "./audit-store";
import { GraphRequestClient, type TokenProvider } from "../../graph-request";
import { SharePointListAuditStore } from "./sharepoint-list-audit-store";

export function createAuditStore(getToken: TokenProvider): AuditStore {
  if (!auditLogEnabled) {
    return new DisabledAuditStore();
  }

  return new SharePointListAuditStore(new GraphRequestClient(getToken));
}
