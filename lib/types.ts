export type AccessRole = "viewer" | "editor" | "owner";

export type SiteSummary = {
  id: string;
  name: string;
  hostname: string;
  path: string;
  webUrl?: string;
  status: "connected" | "ready";
};

export type LibrarySummary = {
  id: string;
  siteId: string;
  name: string;
  driveId?: string;
  rootItemId?: string;
  itemCount: number;
  protected: boolean;
  rightsPolicy: "Library RMS" | "Standard";
  lastScan: string;
  webUrl?: string;
};

export type PermissionSubjectType = "user" | "group" | "sharing-link";

export type ContentItemType = "library" | "folder" | "file";

export type ContentItem = {
  id: string;
  siteId: string;
  name: string;
  type: ContentItemType;
  driveId?: string;
  itemId?: string;
  parentItemId?: string;
  webUrl?: string;
  size?: number;
  childCount?: number;
  modified?: string;
  protected: boolean;
  rightsPolicy: "Library RMS" | "Standard";
};

export type PermissionEntry = {
  id: string;
  libraryId: string;
  driveId?: string;
  itemId?: string;
  displayName: string;
  email: string;
  type: PermissionSubjectType;
  role: AccessRole;
  source: "direct" | "group" | "link" | "inherited";
  tenant: "baht.net" | "external";
  lastActivity: string;
  canEditRole?: boolean;
  canDelete?: boolean;
};

export type AuditEntry = {
  id: string;
  actor: string;
  action: string;
  target: string;
  time: string;
};
