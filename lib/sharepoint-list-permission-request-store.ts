import { auditSite, permissionRequestListName } from "./app-config";
import type { GraphCollection, GraphRequestClient } from "./graph-request";
import type { PermissionRequestStore } from "./permission-request-store";
import type { PermissionRequestDraft } from "./types";

type GraphSite = {
  id: string;
};

type GraphList = {
  id: string;
  displayName?: string;
};

type GraphColumn = {
  name?: string;
};

type GraphListItem = {
  id: string;
};

export class SharePointListPermissionRequestStore implements PermissionRequestStore {
  private requestListPromises = new Map<string, Promise<{ siteId: string; listId: string }>>();
  private defaultSiteIdPromise?: Promise<string>;

  constructor(private readonly graph: GraphRequestClient) {}

  async submit(request: PermissionRequestDraft) {
    const requestTarget = await this.getRequestList(request.siteId);

    await this.graph.request<GraphListItem>(
      `/sites/${encodeURIComponent(requestTarget.siteId)}/lists/${encodeURIComponent(requestTarget.listId)}/items`,
      {
        method: "POST",
        body: JSON.stringify({
          fields: {
            Title: formatPermissionRequestTitle(request),
            Action: request.action,
            Status: request.status,
            ActorEmail: request.actorEmail,
            ActorName: request.actorName,
            ActorRole: request.actorRole,
            TargetEmail: request.targetEmail,
            TargetName: request.targetName,
            RequestedRole: request.requestedRole,
            PreviousRole: request.previousRole ?? "",
            PermissionId: request.permissionId ?? "",
            LibraryId: request.libraryId ?? "",
            DriveId: request.driveId ?? "",
            SiteName: request.siteName ?? "",
            LibraryName: request.libraryName ?? "",
            ItemId: request.itemId ?? "",
            ItemName: request.itemName ?? "",
            Source: request.source ?? "",
            TenantType: request.tenantType,
            RequestedAt: new Date().toISOString(),
          },
        }),
      },
    );
  }

  private async getRequestList(siteId?: string) {
    const targetSiteId = siteId || await this.getDefaultSiteId();
    const existing = this.requestListPromises.get(targetSiteId);
    if (existing) return existing;

    const promise = this.resolveRequestList(targetSiteId);
    this.requestListPromises.set(targetSiteId, promise);
    return promise;
  }

  private async getDefaultSiteId() {
    this.defaultSiteIdPromise ??= this.resolveDefaultSiteId();
    return this.defaultSiteIdPromise;
  }

  private async resolveDefaultSiteId() {
    const site = await this.graph.request<GraphSite>(
      `/sites/${auditSite.hostname}:${encodeURI(auditSite.path)}?$select=id`,
    );
    return site.id;
  }

  private async resolveRequestList(siteId: string) {
    const lists = await this.graph.request<GraphCollection<GraphList>>(
      `/sites/${encodeURIComponent(siteId)}/lists?$select=id,displayName`,
    );
    const existing = (lists.value ?? []).find((list) => list.displayName === permissionRequestListName);
    if (existing) {
      await this.ensureRequestColumns(siteId, existing.id);
      return { siteId, listId: existing.id };
    }

    const created = await this.graph.request<GraphList>(`/sites/${encodeURIComponent(siteId)}/lists`, {
      method: "POST",
      body: JSON.stringify({
        displayName: permissionRequestListName,
        columns: permissionRequestColumns.map((name) => ({
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

  private async ensureRequestColumns(siteId: string, listId: string) {
    const columns = await this.graph.request<GraphCollection<GraphColumn>>(
      `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/columns?$select=name`,
    );
    const existingColumnNames = new Set((columns.value ?? []).map((column) => column.name).filter(Boolean));
    const missingColumns = permissionRequestColumns.filter((name) => !existingColumnNames.has(name));

    await Promise.all(
      missingColumns.map((name) =>
        this.graph.request<GraphColumn>(
          `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/columns`,
          {
            method: "POST",
            body: JSON.stringify({
              name,
              text: {},
            }),
          },
        ),
      ),
    );
  }
}

const permissionRequestColumns = [
  "Action",
  "Status",
  "ActorEmail",
  "ActorName",
  "ActorRole",
  "TargetEmail",
  "TargetName",
  "RequestedRole",
  "PreviousRole",
  "PermissionId",
  "LibraryId",
  "DriveId",
  "SiteName",
  "LibraryName",
  "ItemId",
  "ItemName",
  "Source",
  "TenantType",
  "RequestedAt",
];

function formatPermissionRequestTitle(request: PermissionRequestDraft) {
  return [
    request.action === "GrantAccess" ? "Grant access" : "Update role",
    request.targetEmail || request.targetName,
    request.requestedRole,
  ]
    .filter(Boolean)
    .join(" - ");
}
