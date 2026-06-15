import { auditListName, auditSite } from "./app-config";
import type { GraphCollection, GraphRequestClient } from "./graph-request";
import type { AuditLogDraft } from "./types";
import type { AuditStore } from "./audit-store";

type GraphSite = {
  id: string;
  name?: string;
  displayName?: string;
  webUrl?: string;
};

type GraphList = {
  id: string;
  displayName?: string;
};

type GraphListItem = {
  id: string;
};

export class SharePointListAuditStore implements AuditStore {
  private auditListPromises = new Map<string, Promise<{ siteId: string; listId: string }>>();
  private defaultAuditSiteIdPromise?: Promise<string>;

  constructor(private readonly graph: GraphRequestClient) {}

  async write(entry: AuditLogDraft) {
    const auditTarget = await this.getAuditList(entry.siteId);

    await this.graph.request<GraphListItem>(
      `/sites/${encodeURIComponent(auditTarget.siteId)}/lists/${encodeURIComponent(auditTarget.listId)}/items`,
      {
        method: "POST",
        body: JSON.stringify({
          fields: {
            Title: formatAuditTitle(entry),
            Action: entry.action,
            ActorEmail: entry.actorEmail,
            ActorName: entry.actorName,
            ActorRole: entry.actorRole,
            TargetEmail: entry.targetEmail ?? "",
            TargetName: entry.targetName ?? "",
            PermissionRole: entry.permissionRole ?? "",
            PreviousRole: entry.previousRole ?? "",
            SiteName: entry.siteName ?? "",
            LibraryName: entry.libraryName ?? "",
            ItemId: entry.itemId ?? "",
            Source: entry.source ?? "",
            TenantType: entry.tenantType ?? "",
            Status: entry.status,
            ErrorMessage: entry.errorMessage ?? "",
            GraphRequestId: entry.graphRequestId ?? "",
            CreatedAt: new Date().toISOString(),
          },
        }),
      },
    );
  }

  private async getAuditList(siteId?: string) {
    const auditSiteId = siteId || await this.getDefaultAuditSiteId();
    const existing = this.auditListPromises.get(auditSiteId);
    if (existing) return existing;

    const promise = this.resolveAuditList(auditSiteId);
    this.auditListPromises.set(auditSiteId, promise);
    return promise;
  }

  private async getDefaultAuditSiteId() {
    this.defaultAuditSiteIdPromise ??= this.resolveDefaultAuditSiteId();
    return this.defaultAuditSiteIdPromise;
  }

  private async resolveDefaultAuditSiteId() {
    const site = await this.graph.request<GraphSite>(
      `/sites/${auditSite.hostname}:${encodeURI(auditSite.path)}?$select=id,name,displayName,webUrl`,
    );
    return site.id;
  }

  private async resolveAuditList(siteId: string) {
    const lists = await this.graph.request<GraphCollection<GraphList>>(
      `/sites/${encodeURIComponent(siteId)}/lists?$select=id,displayName`,
    );
    const existing = (lists.value ?? []).find((list) => list.displayName === auditListName);
    if (existing) return { siteId, listId: existing.id };

    const created = await this.graph.request<GraphList>(`/sites/${encodeURIComponent(siteId)}/lists`, {
      method: "POST",
      body: JSON.stringify({
        displayName: auditListName,
        columns: auditColumns.map((name) => ({
          name,
          text: {},
        })),
        list: {
          template: "genericList",
        },
      }),
    });

    return { siteId, listId: created.id };
  }
}

const auditColumns = [
  "Action",
  "ActorEmail",
  "ActorName",
  "ActorRole",
  "TargetEmail",
  "TargetName",
  "PermissionRole",
  "PreviousRole",
  "SiteName",
  "LibraryName",
  "ItemId",
  "Source",
  "TenantType",
  "Status",
  "ErrorMessage",
  "GraphRequestId",
  "CreatedAt",
];

function formatAuditTitle(entry: AuditLogDraft) {
  return [entry.action, entry.targetEmail || entry.targetName || entry.libraryName, entry.status]
    .filter(Boolean)
    .join(" - ");
}
