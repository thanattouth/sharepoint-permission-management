import { auditListName, auditSite } from "../../app-config";
import type { GraphCollection, GraphRequestClient } from "../../graph-request";
import type { AccessRole, AuditLogAction, AuditLogDraft, AuditLogRecord, AuditLogStatus, PermissionEntry } from "../../types";
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

type GraphColumn = {
  name?: string;
};

type GraphListItem = {
  id: string;
  createdDateTime?: string;
  fields?: Record<string, unknown>;
};

export class SharePointListAuditStore implements AuditStore {
  private auditListPromises = new Map<string, Promise<{ siteId: string; listId: string }>>();
  private defaultAuditSiteIdPromise?: Promise<string>;

  constructor(private readonly graph: GraphRequestClient) {}

  async write(entry: AuditLogDraft) {
    const auditTarget = await this.getAuditList(undefined, { ensureColumns: true });

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
            ApprovalRequestNo: entry.approvalRequestNo ?? "",
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
            InviteDeliveryStatus: entry.inviteDeliveryStatus ?? "",
            InviteDiagnostics: entry.inviteDiagnostics ?? "",
            ShareLink: entry.shareLink ?? "",
            CreatedAt: new Date().toISOString(),
          },
        }),
      },
    );
  }

  async list(limit = 100): Promise<AuditLogRecord[]> {
    const auditTarget = await this.getAuditList(undefined, { ensureColumns: false });
    const top = Math.min(Math.max(limit, 1), 500);
    const response = await this.graph.request<GraphCollection<GraphListItem>>(
      `/sites/${encodeURIComponent(auditTarget.siteId)}/lists/${encodeURIComponent(auditTarget.listId)}/items?$top=${top}&$orderby=createdDateTime desc&$expand=fields`,
    );

    return (response.value ?? [])
      .map(mapAuditListItem)
      .sort((left, right) => getAuditTime(right) - getAuditTime(left));
  }

  private async getAuditList(siteId?: string, options: { ensureColumns: boolean } = { ensureColumns: true }) {
    const auditSiteId = siteId || await this.getDefaultAuditSiteId();
    const cacheKey = `${auditSiteId}:${options.ensureColumns ? "write" : "read"}`;
    const existing = this.auditListPromises.get(cacheKey);
    if (existing) return existing;

    const promise = this.resolveAuditList(auditSiteId, options);
    this.auditListPromises.set(cacheKey, promise);
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

  private async resolveAuditList(siteId: string, options: { ensureColumns: boolean }) {
    const lists = await this.graph.request<GraphCollection<GraphList>>(
      `/sites/${encodeURIComponent(siteId)}/lists?$select=id,displayName`,
    );
    const existing = (lists.value ?? []).find((list) => list.displayName === auditListName);
    if (existing) {
      if (options.ensureColumns) {
        await this.ensureAuditColumns(siteId, existing.id);
      }
      return { siteId, listId: existing.id };
    }

    if (!options.ensureColumns) {
      throw new Error(`Audit list "${auditListName}" was not found.`);
    }

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

  private async ensureAuditColumns(siteId: string, listId: string) {
    const columns = await this.graph.request<GraphCollection<GraphColumn>>(
      `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/columns?$select=name`,
    );
    const existingColumnNames = new Set((columns.value ?? []).map((column) => column.name).filter(Boolean));
    const missingColumns = auditColumns.filter((name) => !existingColumnNames.has(name));

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

const auditColumns = [
  "Action",
  "ActorEmail",
  "ActorName",
  "ActorRole",
  "ApprovalRequestNo",
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
  "InviteDeliveryStatus",
  "InviteDiagnostics",
  "ShareLink",
  "CreatedAt",
];

function formatAuditTitle(entry: AuditLogDraft) {
  return [entry.action, entry.targetEmail || entry.targetName || entry.libraryName, entry.status]
    .filter(Boolean)
    .join(" - ");
}

function mapAuditListItem(item: GraphListItem): AuditLogRecord {
  const fields = item.fields ?? {};
  const status = readStatus(fields.Status);

  return {
    id: item.id,
    title: readText(fields.Title) || "Audit event",
    action: readAction(fields.Action),
    actorEmail: readText(fields.ActorEmail),
    actorName: readText(fields.ActorName),
    actorRole: readText(fields.ActorRole),
    approvalRequestNo: readText(fields.ApprovalRequestNo),
    targetEmail: readText(fields.TargetEmail),
    targetName: readText(fields.TargetName),
    permissionRole: readRole(fields.PermissionRole),
    previousRole: readRole(fields.PreviousRole),
    siteName: readText(fields.SiteName),
    libraryName: readText(fields.LibraryName),
    itemId: readText(fields.ItemId),
    source: readSource(fields.Source),
    tenantType: readTenant(fields.TenantType),
    status,
    errorMessage: readText(fields.ErrorMessage),
    graphRequestId: readText(fields.GraphRequestId),
    inviteDeliveryStatus: readInviteDeliveryStatus(fields.InviteDeliveryStatus),
    inviteDiagnostics: readText(fields.InviteDiagnostics),
    shareLink: readText(fields.ShareLink),
    createdAt: readText(fields.CreatedAt) || item.createdDateTime || "",
  };
}

function readText(value: unknown) {
  return typeof value === "string" ? value : "";
}

function readAction(value: unknown): AuditLogAction {
  if (value === "GrantAccess" || value === "UpdateRole" || value === "RemoveAccess" || value === "RefreshReport" || value === "Login") {
    return value;
  }
  return "Login";
}

function readStatus(value: unknown): AuditLogStatus {
  return value === "Failed" ? "Failed" : "Success";
}

function readInviteDeliveryStatus(value: unknown): AuditLogRecord["inviteDeliveryStatus"] {
  if (value === "Accepted" || value === "Partial" || value === "Failed" || value === "Unknown") return value;
  return undefined;
}

function readRole(value: unknown): AccessRole | undefined {
  if (value === "viewer" || value === "editor" || value === "owner") return value;
  return undefined;
}

function readSource(value: unknown): PermissionEntry["source"] | undefined {
  if (value === "direct" || value === "group" || value === "link" || value === "inherited") return value;
  return undefined;
}

function readTenant(value: unknown): PermissionEntry["tenant"] | undefined {
  if (value === "internal" || value === "external") return value;
  return undefined;
}

function getAuditTime(record: AuditLogRecord) {
  const time = new Date(record.createdAt).getTime();
  return Number.isNaN(time) ? 0 : time;
}
