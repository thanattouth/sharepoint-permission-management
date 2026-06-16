import { NextRequest, NextResponse } from "next/server";
import { GraphSharePointPermissionClient } from "@/lib/graph";
import { GraphRequestClient } from "@/lib/graph-request";
import { SharePointListAuditStore } from "@/lib/sharepoint-list-audit-store";
import { createClientCredentialTokenProvider } from "@/lib/server-graph-token";
import type { AccessRole, PermissionEntry, PermissionRequestAction } from "@/lib/types";

export const runtime = "nodejs";

type ApplyPermissionRequestPayload = {
  Action?: string;
  ActorEmail?: string;
  ActorName?: string;
  ActorRole?: string;
  TargetEmail?: string;
  TargetName?: string;
  RequestedRole?: string;
  PreviousRole?: string;
  PermissionId?: string;
  LibraryId?: string;
  DriveId?: string;
  SiteId?: string;
  SiteName?: string;
  LibraryName?: string;
  ItemId?: string;
  ItemName?: string;
  Source?: string;
  TenantType?: string;
  fields?: ApplyPermissionRequestPayload;
};

export async function POST(request: NextRequest) {
  const configuredSecret = process.env.POWER_AUTOMATE_SHARED_SECRET;
  const submittedSecret = request.headers.get("x-flow-secret");

  if (!configuredSecret) {
    return NextResponse.json({ error: "POWER_AUTOMATE_SHARED_SECRET is not configured." }, { status: 500 });
  }

  if (!submittedSecret || submittedSecret !== configuredSecret) {
    return NextResponse.json({ error: "Forbidden." }, { status: 403 });
  }

  let payload: ApplyPermissionRequestPayload;
  try {
    payload = await request.json();
  } catch {
    return NextResponse.json({ error: "Invalid JSON body." }, { status: 400 });
  }

  const approval = normalizeApplyPayload(payload);
  if ("error" in approval) {
    return NextResponse.json({ error: approval.error }, { status: 400 });
  }

  const tokenProvider = createClientCredentialTokenProvider();
  const permissionClient = new GraphSharePointPermissionClient(tokenProvider);
  const auditStore = new SharePointListAuditStore(new GraphRequestClient(tokenProvider));

  try {
    if (approval.action === "GrantAccess") {
      await permissionClient.grantPermission(
        {
          id: approval.libraryId,
          siteId: approval.siteId,
          name: approval.itemName,
          type: "library",
          driveId: approval.driveId,
          itemId: approval.itemId,
          protected: false,
          rightsPolicy: "Standard",
        },
        {
          displayName: approval.targetName,
          email: approval.targetEmail,
          role: approval.requestedRole,
        },
      );
    } else {
      await permissionClient.updatePermissionRole(
        {
          id: approval.permissionId ?? "",
          libraryId: approval.libraryId,
          driveId: approval.driveId,
          itemId: approval.itemId,
          displayName: approval.targetName,
          email: approval.targetEmail,
          type: "user",
          role: approval.previousRole ?? "viewer",
          source: approval.source,
          tenant: approval.tenantType as PermissionEntry["tenant"],
          lastActivity: "Approval request",
        },
        approval.requestedRole,
      );
    }

    await auditStore.write({
      action: approval.action,
      status: "Success",
      actorEmail: approval.actorEmail,
      actorName: approval.actorName,
      actorRole: approval.actorRole,
      targetEmail: approval.targetEmail,
      targetName: approval.targetName,
      permissionRole: approval.requestedRole,
      previousRole: approval.previousRole,
      siteId: approval.siteId,
      siteName: approval.siteName,
      libraryName: approval.libraryName,
      itemId: approval.itemId,
      source: approval.source,
      tenantType: approval.tenantType as PermissionEntry["tenant"],
    });

    return NextResponse.json({ status: "Applied" });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Unable to apply approved permission request.";

    await auditStore.write({
      action: approval.action,
      status: "Failed",
      actorEmail: approval.actorEmail,
      actorName: approval.actorName,
      actorRole: approval.actorRole,
      targetEmail: approval.targetEmail,
      targetName: approval.targetName,
      permissionRole: approval.requestedRole,
      previousRole: approval.previousRole,
      siteId: approval.siteId,
      siteName: approval.siteName,
      libraryName: approval.libraryName,
      itemId: approval.itemId,
      source: approval.source,
      tenantType: approval.tenantType as PermissionEntry["tenant"],
      errorMessage: message,
      graphRequestId: extractGraphRequestId(message),
    });

    return NextResponse.json({ error: message, status: "Failed" }, { status: 500 });
  }
}

function normalizeApplyPayload(payload: ApplyPermissionRequestPayload) {
  const fields = payload.fields ?? payload;
  const action = readAction(fields.Action);
  const requestedRole = readMutableRole(fields.RequestedRole);
  const previousRole = readRole(fields.PreviousRole);
  const tenantType = fields.TenantType === "external" ? "external" : "internal";
  const source = readSource(fields.Source);

  const approval = {
    action,
    actorEmail: readRequired(fields.ActorEmail, "ActorEmail"),
    actorName: readRequired(fields.ActorName, "ActorName"),
    actorRole: readRequired(fields.ActorRole, "ActorRole"),
    targetEmail: readRequired(fields.TargetEmail, "TargetEmail"),
    targetName: readRequired(fields.TargetName, "TargetName"),
    requestedRole,
    previousRole,
    permissionId: fields.PermissionId?.trim(),
    libraryId: readRequired(fields.LibraryId, "LibraryId"),
    driveId: readRequired(fields.DriveId, "DriveId"),
    siteId: readRequired(fields.SiteId, "SiteId"),
    siteName: fields.SiteName?.trim(),
    libraryName: fields.LibraryName?.trim(),
    itemId: readRequired(fields.ItemId, "ItemId"),
    itemName: fields.ItemName?.trim() || fields.LibraryName?.trim() || "Approved item",
    source,
    tenantType,
  };

  const missing = Object.entries(approval)
    .filter(([, value]) => typeof value === "string" && value.startsWith(missingRequiredValue))
    .map(([key]) => key);

  if (!action) return { error: "Action must be GrantAccess or UpdateRole." };
  if (!requestedRole) return { error: "RequestedRole must be viewer or editor." };
  if (action === "UpdateRole" && !approval.permissionId) return { error: "PermissionId is required for UpdateRole." };
  if (missing.length) return { error: `Missing required field: ${missing.join(", ")}.` };

  return {
    ...approval,
    action,
    requestedRole,
  };
}

const missingRequiredValue = "__MISSING_REQUIRED_VALUE__";

function readRequired(value: string | undefined, fieldName: string) {
  return value?.trim() || `${missingRequiredValue}:${fieldName}`;
}

function readAction(value: string | undefined): PermissionRequestAction | undefined {
  if (value === "GrantAccess" || value === "UpdateRole") return value;
  return undefined;
}

function readMutableRole(value: string | undefined): Exclude<AccessRole, "owner"> | undefined {
  if (value === "viewer" || value === "editor") return value;
  return undefined;
}

function readRole(value: string | undefined): AccessRole | undefined {
  if (value === "viewer" || value === "editor" || value === "owner") return value;
  return undefined;
}

function readSource(value: string | undefined): PermissionEntry["source"] {
  if (value === "direct" || value === "group" || value === "link" || value === "inherited") return value;
  return "direct";
}

function extractGraphRequestId(message: string) {
  return message.match(/Request ID:\s*([0-9a-f-]+)/i)?.[1];
}
