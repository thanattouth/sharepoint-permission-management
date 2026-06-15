import type {
  AccessRole,
  ContentItem,
  LibrarySummary,
  PermissionEntry,
  ReportPermissionRow,
  ReportSiteSummary,
  ReportSummary,
  SiteSummary,
  UserSuggestion,
} from "./types";
import { isInternalEmail, protectedLibraryNames, targetSites } from "./app-config";
import { normalizeSharePointPrincipals } from "./permission-normalization";

export type PermissionDraft = {
  displayName: string;
  email: string;
  role: AccessRole;
};

export interface SharePointPermissionClient {
  listSites(): Promise<SiteSummary[]>;
  listLibraries(siteId: string): Promise<LibrarySummary[]>;
  listContentItems(siteId: string): Promise<ContentItem[]>;
  listChildren(item: ContentItem): Promise<ContentItem[]>;
  listPermissions(libraryId: string): Promise<PermissionEntry[]>;
  grantPermission(target: ContentItem, draft: PermissionDraft): Promise<PermissionEntry[]>;
  updatePermissionRole(permission: PermissionEntry, role: AccessRole): Promise<PermissionEntry>;
  removePermission(permission: PermissionEntry): Promise<void>;
  searchUsers(query: string): Promise<UserSuggestion[]>;
  getReportSummary(): Promise<ReportSummary>;
}

export const graphScopes = [
  "User.Read",
  "User.ReadBasic.All",
  "Sites.Read.All",
  "Sites.ReadWrite.All",
  "Files.ReadWrite.All",
];

type TokenProvider = () => Promise<string>;

type GraphCollection<T> = {
  value?: T[];
};

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

export class GraphSharePointPermissionClient implements SharePointPermissionClient {
  constructor(private readonly getToken: TokenProvider) {}

  async listSites() {
    const sites = await Promise.all(
      targetSites.map(async (target) => {
        const site = await this.request<GraphSite>(
          `/sites/${target.hostname}:${encodeURI(target.path)}?$select=id,name,displayName,webUrl`,
        );
        return mapSite(site, target.hostname, target.path);
      }),
    );
    return sites;
  }

  async listLibraries(siteId: string) {
    const drives = await this.request<GraphCollection<GraphDrive>>(
      `/sites/${encodeURIComponent(siteId)}/drives?$select=id,name,webUrl`,
    );

    const libraries = await Promise.all(
      (drives.value ?? []).map(async (drive) => {
        const root = await this.request<GraphDriveItem>(
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

    const children = await this.request<GraphCollection<GraphDriveItem>>(
      `/drives/${encodeURIComponent(item.driveId)}/items/${encodeURIComponent(item.itemId)}/children?$select=id,name,webUrl,size,lastModifiedDateTime,folder,file`,
    );

    return (children.value ?? []).map((child) => mapContentChild(item, child));
  }

  async listPermissions(libraryId: string) {
    const [driveId, itemId] = parseLibraryId(libraryId);
    const permissions = await this.request<GraphCollection<GraphPermission>>(
      `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/permissions`,
    );
    return (permissions.value ?? []).map((permission) => mapPermission(libraryId, driveId, itemId, permission));
  }

  async grantPermission(target: ContentItem, draft: PermissionDraft) {
    if (!target.driveId || !target.itemId) {
      throw new Error("Selected item is missing Graph drive metadata.");
    }
    const { driveId, itemId } = target;

    const body = {
      recipients: [{ email: draft.email }],
      requireSignIn: true,
      sendInvitation: false,
      roles: [toGraphRole(draft.role)],
      message: "You have been granted access from SharePoint Permission Management.",
    };

    const inviteResponse = await this.request<GraphCollection<GraphPermission>>(
      `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/invite`,
      {
        method: "POST",
        body: JSON.stringify(body),
      },
    );

    const permissions = inviteResponse.value ?? [];
    const failures = permissions.filter((permission) => permission.error);
    const successfulPermissions = permissions.filter((permission) => !permission.error);

    if (failures.length > 0 && successfulPermissions.length === 0) {
      throw new Error(formatInviteErrors(failures));
    }

    return successfulPermissions.map((permission) => mapPermission(target.id, driveId, itemId, permission));
  }

  async updatePermissionRole(permission: PermissionEntry, role: AccessRole) {
    if (!permission.driveId || !permission.itemId) {
      throw new Error("Selected permission is missing Graph drive metadata.");
    }

    const updated = await this.request<GraphPermission>(
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

    await this.request<void>(
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
    const users = await this.request<GraphCollection<GraphUser>>(
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

  async getReportSummary() {
    const sites = await this.listSites();
    const reportSites = await Promise.all(
      sites.map(async (site) => {
        const libraries = await this.listLibraries(site.id);
        const permissionsByLibrary = await Promise.all(
          libraries.map(async (library) => ({
            library,
            permissions: normalizeSharePointPrincipals(
              await this.listPermissions(library.id).catch(() => []),
              site.name,
            ),
          })),
        );
        const permissions = permissionsByLibrary.flatMap((entry) => entry.permissions);
        return {
          summary: mapReportSite(site, libraries, permissions),
          permissions: permissionsByLibrary.flatMap((entry) =>
            entry.permissions.map((permission) => mapReportPermission(site, entry.library, permission)),
          ),
        };
      }),
    );
    const siteSummaries = reportSites.map((site) => site.summary);
    const permissionRows = reportSites.flatMap((site) => site.permissions);

    return {
      generatedAt: new Date().toISOString(),
      siteCount: siteSummaries.length,
      libraryCount: sumBy(siteSummaries, "libraryCount"),
      protectedLibraryCount: sumBy(siteSummaries, "protectedLibraryCount"),
      standardLibraryCount: sumBy(siteSummaries, "standardLibraryCount"),
      directPermissionCount: sumBy(siteSummaries, "directPermissionCount"),
      externalPermissionCount: sumBy(siteSummaries, "externalPermissionCount"),
      inheritedPermissionCount: sumBy(siteSummaries, "inheritedPermissionCount"),
      sites: siteSummaries,
      permissions: permissionRows,
    };
  }

  private async request<T>(path: string, init: RequestInit = {}) {
    const token = await this.getToken();
    const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      ...init,
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        ...init.headers,
      },
    });

    if (!response.ok) {
      const body = await response.text();
      throw new Error(formatGraphError(response.status, body || response.statusText));
    }

    if (response.status === 204) {
      return undefined as T;
    }

    return (await response.json()) as T;
  }
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

function mapReportSite(
  site: SiteSummary,
  libraries: LibrarySummary[],
  permissions: PermissionEntry[],
): ReportSiteSummary {
  return {
    siteId: site.id,
    siteName: site.name,
    hostname: site.hostname,
    libraryCount: libraries.length,
    protectedLibraryCount: libraries.filter((library) => library.protected).length,
    standardLibraryCount: libraries.filter((library) => !library.protected).length,
    directPermissionCount: permissions.filter(isDirectlyManageablePermission).length,
    externalPermissionCount: permissions.filter((permission) => permission.tenant === "external").length,
    inheritedPermissionCount: permissions.filter((permission) => permission.source === "inherited").length,
  };
}

function isDirectlyManageablePermission(permission: PermissionEntry) {
  return permission.canEditRole !== false || permission.canDelete !== false;
}

function mapReportPermission(
  site: SiteSummary,
  library: LibrarySummary,
  permission: PermissionEntry,
): ReportPermissionRow {
  return {
    id: `${site.id}:${library.id}:${permission.id}`,
    siteName: site.name,
    libraryName: library.name,
    principalName: permission.displayName,
    email: permission.email,
    role: permission.role,
    source: permission.source,
    tenant: permission.tenant,
  };
}

function sumBy(items: ReportSiteSummary[], key: keyof Pick<
  ReportSiteSummary,
  | "libraryCount"
  | "protectedLibraryCount"
  | "standardLibraryCount"
  | "directPermissionCount"
  | "externalPermissionCount"
  | "inheritedPermissionCount"
>) {
  return items.reduce((total, item) => total + item[key], 0);
}

function formatGraphError(status: number, body: string) {
  try {
    const parsed = JSON.parse(body) as {
      error?: {
        code?: string;
        message?: string;
        innerError?: {
          "request-id"?: string;
          date?: string;
        };
      };
    };
    const code = parsed.error?.code;
    const message = parsed.error?.message;
    const requestId = parsed.error?.innerError?.["request-id"];

    if (code === "sharingFailed") {
      return [
        "Graph sharing failed. SharePoint rejected this invitation.",
        message ? `Graph message: ${message}` : "",
        "Check SharePoint organization sharing, site sharing, domain restrictions, Microsoft Entra external collaboration settings, and whether this email domain is allowed.",
        requestId ? `Request ID: ${requestId}` : "",
      ]
        .filter(Boolean)
        .join(" ");
    }

    return [`Graph ${status}${code ? ` ${code}` : ""}: ${message || body}`, requestId ? `Request ID: ${requestId}` : ""]
      .filter(Boolean)
      .join(" ");
  } catch {
    return `Graph ${status}: ${body}`;
  }
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
    "Check external sharing policy, allowed domains, guest invitation settings, and whether the target email can be invited to this tenant.",
  ].join(" ");
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

function libraryToContentItem(library: LibrarySummary): ContentItem {
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

function parseLibraryId(libraryId: string) {
  const [driveId, itemId] = libraryId.split(":");
  if (!driveId || !itemId) {
    throw new Error("Invalid Graph library id.");
  }
  return [driveId, itemId] as const;
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
