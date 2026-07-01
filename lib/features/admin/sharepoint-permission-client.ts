import { isInternalEmail, protectedLibraryNames, targetSites } from "../../app-config";
import { GraphRequestClient, type GraphCollection, type TokenProvider } from "../../graph-request";
import type {
  AccessRole,
  AccessReadinessResult,
  ContentItem,
  LibrarySummary,
  PermissionEntry,
  SiteSummary,
  UserSuggestion,
  InviteDeliveryStatus,
} from "../../types";

export type PermissionDraft = {
  displayName: string;
  email: string;
  role: AccessRole;
};

export type InviteDiagnostic = {
  email: string;
  permissionId?: string;
  code?: string;
  innerCode?: string;
  message?: string;
  guidance?: string;
};

export type GrantPermissionResult = {
  permissions: PermissionEntry[];
  inviteDeliveryStatus: InviteDeliveryStatus;
  inviteDiagnostics: InviteDiagnostic[];
  accessReadiness: AccessReadinessResult;
  shareLink?: string;
};

export type RemovePermissionResult = {
  removedPermissionIds: string[];
  remainingPermissions: PermissionEntry[];
  blockedPermissions: PermissionEntry[];
  status: "revoked" | "residual-access-risk" | "not-removed";
  verificationDetails: string[];
};

export class GraphInviteFailureError extends Error {
  constructor(
    message: string,
    readonly inviteDeliveryStatus: InviteDeliveryStatus,
    readonly inviteDiagnostics: InviteDiagnostic[],
  ) {
    super(message);
    this.name = "GraphInviteFailureError";
  }
}

export interface SharePointInventoryReader {
  listSites(): Promise<SiteSummary[]>;
  resolveSite(hostname: string, path: string): Promise<SiteSummary>;
  listLibraries(siteId: string): Promise<LibrarySummary[]>;
  listChildren(item: ContentItem): Promise<ContentItem[]>;
  listPermissions(itemId: string): Promise<PermissionEntry[]>;
}

export interface SharePointPermissionClient extends SharePointInventoryReader {
  listContentItems(siteId: string): Promise<ContentItem[]>;
  grantPermission(target: ContentItem, draft: PermissionDraft): Promise<GrantPermissionResult>;
  createShareLink(target: ContentItem, role: AccessRole): Promise<string | undefined>;
  updatePermissionRole(permission: PermissionEntry, role: AccessRole): Promise<PermissionEntry>;
  removePermission(permission: PermissionEntry): Promise<RemovePermissionResult>;
  searchUsers(query: string): Promise<UserSuggestion[]>;
}

export const graphReadScopes = [
  "User.Read",
  "User.ReadBasic.All",
  "Sites.Read.All",
];

export const graphWriteScopes = [
  ...graphReadScopes,
  "Sites.ReadWrite.All",
  "Files.ReadWrite.All",
];

type GraphSite = {
  id: string;
  name?: string;
  displayName?: string;
  webUrl?: string;
};

type GraphDrive = {
  id: string;
  name?: string;
  webUrl?: string;
};

type GraphDriveItem = {
  id: string;
  name?: string;
  webUrl?: string;
  size?: number;
  lastModifiedDateTime?: string;
  file?: unknown;
  folder?: {
    childCount?: number;
  };
};

type GraphIdentity = {
  id?: string;
  displayName?: string;
  email?: string;
  loginName?: string;
  userPrincipalName?: string;
};

type GraphUser = {
  id: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
  jobTitle?: string;
};

type GraphIdentitySet = {
  user?: GraphIdentity;
  group?: GraphIdentity;
  application?: GraphIdentity;
  siteUser?: GraphIdentity;
  siteGroup?: GraphIdentity;
  sharePointGroup?: GraphIdentity;
};

type GraphPermission = {
  id: string;
  roles?: string[];
  inheritedFrom?: unknown;
  link?: {
    type?: string;
    scope?: string;
    webUrl?: string;
  };
  invitation?: {
    email?: string;
  };
  grantedTo?: GraphIdentitySet;
  grantedToIdentities?: GraphIdentitySet[];
  grantedToV2?: GraphIdentitySet;
  grantedToIdentitiesV2?: GraphIdentitySet[];
  error?: GraphInviteError;
};

type GraphInviteError = {
  code?: string;
  message?: string;
  localizedMessage?: string;
  innererror?: {
    code?: string;
  };
};

export class GraphSharePointAdminClient implements SharePointPermissionClient {
  private readonly graph: GraphRequestClient;

  constructor(getToken: TokenProvider) {
    this.graph = new GraphRequestClient(getToken);
  }

  async listSites() {
    const sites = await Promise.all(
      targetSites.map((target) => this.resolveSite(target.hostname, target.path)),
    );
    return sites;
  }

  async resolveSite(hostname: string, path: string) {
    const normalizedHostname = hostname.trim().toLowerCase();
    const normalizedPath = normalizeSitePath(path);
    const site = await this.graph.request<GraphSite>(
      `/sites/${normalizedHostname}:${encodeURI(normalizedPath)}?$select=id,name,displayName,webUrl`,
    );
    return mapSite(site, normalizedHostname, normalizedPath);
  }

  async listLibraries(siteId: string) {
    const drives = await this.graph.request<GraphCollection<GraphDrive>>(
      `/sites/${encodeURIComponent(siteId)}/drives?$select=id,name,webUrl`,
    );

    const libraries = await Promise.all(
      (drives.value ?? []).map(async (drive) => {
        const root = await this.graph.request<GraphDriveItem>(
          `/drives/${encodeURIComponent(drive.id)}/root?$select=id,name,folder`,
        );
        return mapLibrary(siteId, drive, root);
      }),
    );

    return libraries.sort((left, right) => {
      const leftProtected = protectedLibraryNames.has(left.name) ? 0 : 1;
      const rightProtected = protectedLibraryNames.has(right.name) ? 0 : 1;
      return leftProtected - rightProtected || left.name.localeCompare(right.name);
    });
  }

  async listContentItems(siteId: string) {
    const libraries = await this.listLibraries(siteId);
    return libraries.map((library) => libraryToContentItem(library));
  }

  async listChildren(item: ContentItem) {
    if (!item.driveId || !item.itemId) {
      throw new Error("Selected item is missing Graph drive metadata.");
    }

    const children: GraphDriveItem[] = [];
    let nextPath = `/drives/${encodeURIComponent(item.driveId)}/items/${encodeURIComponent(item.itemId)}/children?$top=200&$select=id,name,webUrl,size,lastModifiedDateTime,folder,file`;

    while (nextPath) {
      const response = await this.graph.request<GraphCollection<GraphDriveItem>>(nextPath);
      children.push(...(response.value ?? []));
      nextPath = response["@odata.nextLink"] ?? "";
    }

    return children.map((child) => mapContentChild(item, child));
  }

  async listPermissions(itemId: string) {
    const [driveId, graphItemId] = parseDriveItemId(itemId);
    const permissions = await this.graph.request<GraphCollection<GraphPermission>>(
      `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(graphItemId)}/permissions`,
    );
    return (permissions.value ?? []).map((permission) => mapPermission(itemId, driveId, graphItemId, permission));
  }

  async grantPermission(target: ContentItem, draft: PermissionDraft) {
    if (!target.driveId || !target.itemId) {
      throw new Error("Selected item is missing Graph drive metadata.");
    }
    const recipientEmail = normalizeRecipientEmail(draft.email);
    if (!recipientEmail) {
      throw new Error("Enter a valid recipient email address before granting SharePoint access.");
    }
    const { driveId, itemId } = target;

    const body = {
      recipients: [{ email: recipientEmail }],
      requireSignIn: true,
      sendInvitation: true,
      roles: [toGraphRole(draft.role)],
      message: "You have been granted access from SharePoint Permission Management.",
    };

    const inviteResponse = await this.graph.request<GraphCollection<GraphPermission>>(
      `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/invite`,
      {
        method: "POST",
        body: JSON.stringify(body),
      },
    );

    const permissions = inviteResponse.value ?? [];
    const failures = permissions.filter((permission) => permission.error);
    const successfulPermissions = permissions.filter((permission) => !permission.error);
    const inviteDiagnostics = permissions.map((permission) => mapInviteDiagnostic(permission, recipientEmail));
    const inviteDeliveryStatus = getInviteDeliveryStatus(successfulPermissions.length, failures.length);

    if (failures.length > 0 && successfulPermissions.length === 0) {
      throw new GraphInviteFailureError(formatInviteErrors(failures), inviteDeliveryStatus, inviteDiagnostics);
    }

    const shareLink = target.webUrl ?? getPermissionShareLink(successfulPermissions);

    return {
      permissions: successfulPermissions.map((permission) => mapPermission(target.id, driveId, itemId, permission)),
      inviteDeliveryStatus,
      inviteDiagnostics,
      accessReadiness: assessGrantReadiness(target, { ...draft, email: recipientEmail }, inviteDeliveryStatus, successfulPermissions.length),
      shareLink,
    };
  }

  private async createItemShareLink(driveId: string, itemId: string, role: AccessRole) {
    const permission = await this.graph.request<GraphPermission>(
      `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/createLink`,
      {
        method: "POST",
        body: JSON.stringify({
          type: toGraphLinkType(role),
          scope: "users",
          retainInheritedPermissions: true,
        }),
      },
    );

    return permission.link?.webUrl;
  }

  private async createRecipientShareLink(driveId: string, itemId: string, recipientEmail: string, role: AccessRole) {
    const shareLink = await this.createItemShareLink(driveId, itemId, role);
    if (!shareLink) return undefined;

    const grantResponse = await this.graph.request<GraphCollection<GraphPermission>>(
      `/shares/${encodeSharingUrl(shareLink)}/permission/grant`,
      {
        method: "POST",
        body: JSON.stringify({
          recipients: [{ email: recipientEmail }],
          roles: [toGraphRole(role)],
        }),
      },
    );

    return getPermissionShareLink(grantResponse.value ?? []) ?? shareLink;
  }

  async createShareLink(target: ContentItem, role: AccessRole) {
    if (!target.driveId || !target.itemId) {
      throw new Error("Selected item is missing Graph drive metadata.");
    }

    return this.createItemShareLink(target.driveId, target.itemId, role);
  }

  async updatePermissionRole(permission: PermissionEntry, role: AccessRole) {
    if (!permission.driveId || !permission.itemId) {
      throw new Error("Selected permission is missing Graph drive metadata.");
    }

    const updated = await this.graph.request<GraphPermission>(
      `/drives/${encodeURIComponent(permission.driveId)}/items/${encodeURIComponent(permission.itemId)}/permissions/${encodeURIComponent(permission.id)}`,
      {
        method: "PATCH",
        body: JSON.stringify({ roles: [toGraphRole(role)] }),
      },
    );

    return mapPermission(permission.libraryId, permission.driveId, permission.itemId, updated);
  }

  async removePermission(permission: PermissionEntry) {
    if (!permission.driveId || !permission.itemId) {
      throw new Error("Selected permission is missing Graph drive metadata.");
    }

    const permissions = await this.graph.request<GraphCollection<GraphPermission>>(
      `/drives/${encodeURIComponent(permission.driveId)}/items/${encodeURIComponent(permission.itemId)}/permissions`,
    );
    const removablePermissions = (permissions.value ?? []).filter((candidate) =>
      shouldRemovePermission(candidate, permission),
    );

    for (const candidate of removablePermissions) {
      await this.deletePermission(permission.driveId, permission.itemId, candidate.id);
    }

    const remainingPermissions = await this.listPermissions(`${permission.driveId}:${permission.itemId}`);
    const blockedPermissions = getResidualAccessRisks(remainingPermissions, permission);
    const verificationDetails = getRemovalVerificationDetails(removablePermissions, blockedPermissions);

    return {
      removedPermissionIds: removablePermissions.map((candidate) => candidate.id),
      remainingPermissions,
      blockedPermissions,
      status: getRemovalStatus(removablePermissions, blockedPermissions),
      verificationDetails,
    };
  }

  private async deletePermission(driveId: string, itemId: string, permissionId: string) {
    await this.graph.request<void>(
      `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/permissions/${encodeURIComponent(permissionId)}`,
      {
        method: "DELETE",
      },
    );
  }

  async searchUsers(query: string) {
    const normalizedQuery = normalizeDirectorySearchQuery(query);
    if (!normalizedQuery) return [];

    const search = encodeURIComponent(
      `"displayName:${normalizedQuery}" OR "mail:${normalizedQuery}" OR "userPrincipalName:${normalizedQuery}"`,
    );
    const users = await this.graph.request<GraphCollection<GraphUser>>(
      `/users?$search=${search}&$select=id,displayName,mail,userPrincipalName,jobTitle&$top=8&$count=true`,
      {
        headers: {
          ConsistencyLevel: "eventual",
        },
      },
    );

    return (users.value ?? [])
      .map(mapUserSuggestion)
      .filter((user): user is UserSuggestion => Boolean(user));
  }
}

export function libraryToContentItem(library: LibrarySummary): ContentItem {
  return {
    id: library.id,
    siteId: library.siteId,
    name: library.name,
    type: "library",
    driveId: library.driveId,
    itemId: library.rootItemId,
    webUrl: library.webUrl,
    childCount: library.itemCount,
    modified: library.lastScan,
    protected: library.protected,
    rightsPolicy: library.rightsPolicy,
  };
}

function normalizeDirectorySearchQuery(value: string) {
  return value.trim().replace(/["\\]/g, "").slice(0, 64);
}

function mapUserSuggestion(user: GraphUser): UserSuggestion | undefined {
  const email = user.mail || user.userPrincipalName;
  if (!email) return undefined;

  return {
    id: user.id,
    displayName: user.displayName || email,
    email,
    jobTitle: user.jobTitle,
  };
}

function formatInviteErrors(permissions: GraphPermission[]) {
  const details = permissions
    .map((permission) => {
      const email = permission.invitation?.email;
      const error = permission.error;
      const code = error?.innererror?.code ?? error?.code;
      const message = error?.localizedMessage ?? error?.message;
      const guidance = getInviteGuidance(code, message);
      return [email, code, message, guidance].filter(Boolean).join(": ");
    })
    .filter(Boolean);

  return [
    "Graph invitation failed.",
    ...details,
    "Check SharePoint organization sharing, site sharing, domain restrictions, Microsoft Entra external collaboration settings, and whether this email domain is allowed.",
  ].join(" ");
}

function getInviteDeliveryStatus(successCount: number, failureCount: number): InviteDeliveryStatus {
  if (failureCount > 0 && successCount > 0) return "Partial";
  if (failureCount > 0) return "Failed";
  if (successCount > 0) return "Accepted";
  return "Unknown";
}

function getPermissionShareLink(permissions: GraphPermission[]) {
  const links = permissions.map((permission) => permission.link?.webUrl).filter((url): url is string => Boolean(url));
  return links.find((url) => !url.toLowerCase().includes("/guestaccess.aspx"));
}

function shouldRemovePermission(candidate: GraphPermission, selected: PermissionEntry) {
  if (candidate.id === selected.id) return true;
  if (candidate.inheritedFrom) return false;

  const subject = extractSubject(candidate);
  if (!isSameGuestIdentity(subject.email, selected.email) && !isSameGuestIdentity(subject.displayName, selected.email)) {
    return false;
  }

  return !candidate.link || candidate.invitation?.email || subject.type === "user";
}

function isSameGuestPermission(candidate: PermissionEntry, selected: PermissionEntry) {
  return isSameGuestIdentity(candidate.email, selected.email) || isSameGuestIdentity(candidate.displayName, selected.email);
}

function getResidualAccessRisks(remainingPermissions: PermissionEntry[], selected: PermissionEntry) {
  return remainingPermissions.filter((candidate) => {
    if (isSameGuestPermission(candidate, selected)) return true;

    if (selected.tenant === "external" && candidate.source === "link") {
      return true;
    }

    if (selected.tenant === "external" && candidate.source === "inherited") {
      return candidate.tenant === "external" || candidate.type === "sharing-link";
    }

    return false;
  });
}

function getRemovalStatus(removablePermissions: GraphPermission[], blockedPermissions: PermissionEntry[]): RemovePermissionResult["status"] {
  if (removablePermissions.length === 0) return "not-removed";
  if (blockedPermissions.length > 0) return "residual-access-risk";
  return "revoked";
}

function getRemovalVerificationDetails(removablePermissions: GraphPermission[], blockedPermissions: PermissionEntry[]) {
  if (removablePermissions.length === 0) {
    return [
      "No removable direct permission matching this principal was found on the selected item.",
      "The access may come from a parent folder, site permission, group membership, or a sharing link.",
    ];
  }

  if (blockedPermissions.length === 0) {
    return [
      `Removed ${removablePermissions.length} permission entr${removablePermissions.length === 1 ? "y" : "ies"} from the selected item.`,
      "Verification scan did not find the same principal or external sharing-link risk on the refreshed permission list.",
    ];
  }

  const groupedRisks = summarizeResidualRisks(blockedPermissions);
  return [
    `Removed ${removablePermissions.length} permission entr${removablePermissions.length === 1 ? "y" : "ies"} from the selected item.`,
    ...groupedRisks,
    "To fully revoke access, remove the remaining permission at its source: parent folder/site, SharePoint group, or sharing link.",
  ];
}

function summarizeResidualRisks(permissions: PermissionEntry[]) {
  const counts = permissions.reduce<Record<PermissionEntry["source"], number>>(
    (summary, permission) => {
      summary[permission.source] += 1;
      return summary;
    },
    { direct: 0, group: 0, link: 0, inherited: 0 },
  );

  return [
    counts.direct > 0 ? `${counts.direct} direct permission entr${counts.direct === 1 ? "y" : "ies"} still match this principal.` : "",
    counts.link > 0 ? `${counts.link} sharing link entr${counts.link === 1 ? "y" : "ies"} may still allow access to anyone who has the link.` : "",
    counts.inherited > 0 ? `${counts.inherited} inherited permission entr${counts.inherited === 1 ? "y" : "ies"} still flow from a parent folder or site.` : "",
    counts.group > 0 ? `${counts.group} group permission entr${counts.group === 1 ? "y" : "ies"} remain; verify the guest is not a member of those groups.` : "",
  ].filter(Boolean);
}

function isSameGuestIdentity(left: string | undefined, right: string | undefined) {
  const normalizedLeft = normalizeGuestIdentity(left);
  const normalizedRight = normalizeGuestIdentity(right);
  return Boolean(normalizedLeft && normalizedRight && normalizedLeft === normalizedRight);
}

function normalizeGuestIdentity(value: string | undefined) {
  const normalized = value?.trim().toLowerCase();
  if (!normalized || normalized === "invitation pending" || normalized.includes("principal details unavailable")) return "";
  return normalized;
}

function mapInviteDiagnostic(permission: GraphPermission, fallbackEmail: string): InviteDiagnostic {
  const error = permission.error;
  const message = error?.localizedMessage ?? error?.message;
  const code = error?.innererror?.code ?? error?.code;
  return {
    email: permission.invitation?.email || fallbackEmail,
    permissionId: permission.id,
    code: error?.code,
    innerCode: error?.innererror?.code,
    message,
    guidance: getInviteGuidance(code, message),
  };
}

function assessGrantReadiness(
  target: ContentItem,
  draft: PermissionDraft,
  inviteDeliveryStatus: InviteDeliveryStatus,
  successfulPermissionCount: number,
): AccessReadinessResult {
  const details: string[] = [];
  const externalRecipient = !isInternalEmail(draft.email);

  if (successfulPermissionCount === 0 || inviteDeliveryStatus === "Failed") {
    return {
      status: "blocked",
      title: "Permission was not granted",
      details: ["SharePoint rejected the invitation or did not return a usable permission. Review the Graph diagnostics before sending a link."],
    };
  }

  if (inviteDeliveryStatus === "Partial") {
    details.push("SharePoint accepted only part of the invitation response. Confirm the recipient appears in Manage Access before sending the link.");
  }

  if (externalRecipient) {
    details.push("External guests must redeem the invitation and open the link with the exact same email address that was granted access.");
    details.push("If Microsoft asks which organization to use, the guest must choose the resource organization that owns this SharePoint site.");
    details.push("If the recipient still sees You need access, verify SharePoint org/site external sharing, Entra B2B collaboration, cross-tenant access, Conditional Access, and domain allow/block lists.");
  }

  if (target.protected) {
    details.push("This item is in a protected library. SharePoint permission does not guarantee that Rights Management or a sensitivity label will let the recipient open the file content.");
  }

  if (target.type === "file") {
    details.push("Use the direct file link for validation. Parent folders may not be browsable unless the recipient also has folder or site access.");
  }

  if (details.length === 0) {
    return {
      status: "ready",
      title: `Access granted for ${draft.email}`,
      details: ["SharePoint accepted the permission change. Send the link if the invitation email is not received."],
    };
  }

  return {
    status: "caution",
    title: `Permission granted for ${draft.email}; recipient validation required`,
    details,
  };
}

function normalizeRecipientEmail(value: string) {
  const normalized = value.trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalized)) return "";
  return normalized;
}

function normalizeSitePath(value: string) {
  const trimmed = value.trim();
  return trimmed.startsWith("/") ? trimmed : `/${trimmed}`;
}

function getInviteGuidance(code?: string, message = "") {
  const normalized = `${code ?? ""} ${message}`.toLowerCase();

  if (!normalized.trim()) return undefined;

  if (normalized.includes("domain") || normalized.includes("blocked") || normalized.includes("not allowed")) {
    return "Recipient domain may be blocked by SharePoint or Entra external collaboration settings.";
  }

  if (normalized.includes("external") || normalized.includes("guest") || normalized.includes("invitation")) {
    return "External guest invitation may be blocked by SharePoint site sharing, B2B collaboration, or cross-tenant access policy.";
  }

  if (normalized.includes("conditional access") || normalized.includes("mfa")) {
    return "Conditional Access or MFA requirements may be blocking guest sign-in.";
  }

  if (normalized.includes("sensitivity") || normalized.includes("label") || normalized.includes("irm") || normalized.includes("rights")) {
    return "File protection may require additional Purview sensitivity label or Rights Management access.";
  }

  return undefined;
}

function mapSite(site: GraphSite, hostname: string, path: string): SiteSummary {
  return {
    id: site.id,
    name: site.displayName || site.name || path.split("/").at(-1) || path,
    hostname,
    path,
    webUrl: site.webUrl,
    status: "connected",
  };
}

function mapLibrary(siteId: string, drive: GraphDrive, root: GraphDriveItem): LibrarySummary {
  const name = drive.name || root.name || "Document Library";
  const protectedLibrary = protectedLibraryNames.has(name);
  return {
    id: `${drive.id}:${root.id}`,
    siteId,
    name,
    driveId: drive.id,
    rootItemId: root.id,
    itemCount: root.folder?.childCount ?? 0,
    protected: protectedLibrary,
    rightsPolicy: protectedLibrary ? "Library RMS" : "Standard",
    lastScan: "Live Graph",
    webUrl: drive.webUrl,
  };
}

function mapContentChild(parent: ContentItem, child: GraphDriveItem): ContentItem {
  const type = child.folder ? "folder" : "file";
  return {
    id: `${parent.driveId}:${child.id}`,
    siteId: parent.siteId,
    name: child.name || "Untitled",
    type,
    driveId: parent.driveId,
    itemId: child.id,
    parentItemId: parent.itemId,
    webUrl: child.webUrl,
    size: child.size,
    childCount: child.folder?.childCount,
    modified: child.lastModifiedDateTime ? new Date(child.lastModifiedDateTime).toLocaleDateString() : "Live Graph",
    protected: parent.protected,
    rightsPolicy: parent.rightsPolicy,
  };
}

function mapPermission(
  libraryId: string,
  driveId: string,
  itemId: string,
  permission: GraphPermission,
): PermissionEntry {
  const subject = extractSubject(permission);
  const source = permission.inheritedFrom
    ? "inherited"
    : permission.link
      ? "link"
      : subject.type === "group"
        ? "group"
        : "direct";
  const role = fromGraphRole(permission.roles?.[0]);
  return {
    id: permission.id,
    libraryId,
    driveId,
    itemId,
    displayName: subject.displayName,
    email: subject.email,
    type: permission.link ? "sharing-link" : subject.type,
    role,
    source,
    tenant: isInternalPrincipal(subject) ? "internal" : "external",
    lastActivity: source === "inherited" ? "Inherited from parent" : "Live Graph",
    canEditRole: !permission.inheritedFrom && !permission.link,
    canDelete: !permission.inheritedFrom,
  };
}

function isInternalPrincipal(subject: { displayName: string; email: string }) {
  return isInternalEmail(subject.email) || subject.email === subject.displayName;
}

function extractSubject(permission: GraphPermission): {
  displayName: string;
  email: string;
  type: "user" | "group";
} {
  const identities = [
    permission.grantedToV2,
    ...(permission.grantedToIdentitiesV2 ?? []),
    permission.grantedTo,
    ...(permission.grantedToIdentities ?? []),
  ].filter(Boolean);

  for (const identity of identities) {
    const sharePointGroup = identity?.sharePointGroup ?? identity?.siteGroup;
    if (sharePointGroup) return mapIdentity(sharePointGroup, "group");

    const group = identity?.group;
    if (group) return mapIdentity(group, "group");

    const siteUser = identity?.siteUser;
    if (siteUser) return mapIdentity(siteUser, "user");

    const user = identity?.user;
    if (user) return mapIdentity(user, "user");

    const application = identity?.application;
    if (application) return mapIdentity(application, "group");
  }

  if (permission.invitation?.email) {
    return {
      displayName: permission.invitation.email,
      email: "Invitation pending",
      type: "user",
    };
  }

  if (permission.link) {
    const label = `${permission.link.scope ?? "sharing"} ${permission.link.type ?? "link"}`;
    return {
      displayName: label,
      email: permission.link.webUrl ?? label,
      type: "user",
    };
  }

  return {
    displayName: `Permission ${permission.id}`,
    email: "Principal details unavailable from Graph",
    type: "user",
  };
}

function mapIdentity(identity: GraphIdentity, type: "user" | "group") {
  const displayName =
    identity.displayName ||
    identity.email ||
    identity.userPrincipalName ||
    identity.loginName ||
    identity.id ||
    (type === "group" ? "SharePoint group" : "User");
  const email =
    identity.email ||
    identity.userPrincipalName ||
    identity.loginName ||
    identity.displayName ||
    displayName;

  return {
    displayName,
    email,
    type,
  };
}

function parseDriveItemId(itemId: string) {
  const [driveId, graphItemId] = itemId.split(":");
  if (!driveId || !graphItemId) {
    throw new Error("Invalid Graph item id.");
  }
  return [driveId, graphItemId] as const;
}

function toGraphRole(role: AccessRole) {
  if (role === "editor") return "write";
  if (role === "owner") return "owner";
  return "read";
}

function toGraphLinkType(role: AccessRole) {
  return role === "editor" || role === "owner" ? "edit" : "view";
}

function encodeSharingUrl(url: string) {
  return `u!${globalThis.btoa(url).replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-")}`;
}

function fromGraphRole(role = "read"): AccessRole {
  if (role === "write") return "editor";
  if (role === "owner") return "owner";
  return "viewer";
}
