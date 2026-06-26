import { isInternalEmail, protectedLibraryNames, targetSites } from "../../app-config";
import { GraphRequestClient, type GraphCollection, type TokenProvider } from "../../graph-request";
import type {
  AccessRole,
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
};

export type GrantPermissionResult = {
  permissions: PermissionEntry[];
  inviteDeliveryStatus: InviteDeliveryStatus;
  inviteDiagnostics: InviteDiagnostic[];
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
  listLibraries(siteId: string): Promise<LibrarySummary[]>;
  listChildren(item: ContentItem): Promise<ContentItem[]>;
  listPermissions(itemId: string): Promise<PermissionEntry[]>;
}

export interface SharePointPermissionClient extends SharePointInventoryReader {
  listContentItems(siteId: string): Promise<ContentItem[]>;
  grantPermission(target: ContentItem, draft: PermissionDraft): Promise<GrantPermissionResult>;
  updatePermissionRole(permission: PermissionEntry, role: AccessRole): Promise<PermissionEntry>;
  removePermission(permission: PermissionEntry): Promise<void>;
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
      targetSites.map(async (target) => {
        const site = await this.graph.request<GraphSite>(
          `/sites/${target.hostname}:${encodeURI(target.path)}?$select=id,name,displayName,webUrl`,
        );
        return mapSite(site, target.hostname, target.path);
      }),
    );
    return sites;
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
    const { driveId, itemId } = target;

    const body = {
      recipients: [{ email: draft.email }],
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
    const inviteDiagnostics = permissions.map((permission) => mapInviteDiagnostic(permission, draft.email));
    const inviteDeliveryStatus = getInviteDeliveryStatus(successfulPermissions.length, failures.length);

    if (failures.length > 0 && successfulPermissions.length === 0) {
      throw new GraphInviteFailureError(formatInviteErrors(failures), inviteDeliveryStatus, inviteDiagnostics);
    }

    return {
      permissions: successfulPermissions.map((permission) => mapPermission(target.id, driveId, itemId, permission)),
      inviteDeliveryStatus,
      inviteDiagnostics,
    };
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

    await this.graph.request<void>(
      `/drives/${encodeURIComponent(permission.driveId)}/items/${encodeURIComponent(permission.itemId)}/permissions/${encodeURIComponent(permission.id)}`,
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
      return [email, code, message].filter(Boolean).join(": ");
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

function mapInviteDiagnostic(permission: GraphPermission, fallbackEmail: string): InviteDiagnostic {
  const error = permission.error;
  return {
    email: permission.invitation?.email || fallbackEmail,
    permissionId: permission.id,
    code: error?.code,
    innerCode: error?.innererror?.code,
    message: error?.localizedMessage ?? error?.message,
  };
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

function fromGraphRole(role = "read"): AccessRole {
  if (role === "write") return "editor";
  if (role === "owner") return "owner";
  return "viewer";
}
