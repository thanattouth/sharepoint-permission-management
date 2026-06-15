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
  tenant: "internal" | "external";
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
  status?: "Success" | "Failed";
};

export type AuditLogAction =
  | "GrantAccess"
  | "UpdateRole"
  | "RemoveAccess"
  | "RefreshReport"
  | "Login";

export type AuditLogStatus = "Success" | "Failed";

export type AuditLogDraft = {
  action: AuditLogAction;
  actorEmail: string;
  actorName: string;
  actorRole: string;
  targetEmail?: string;
  targetName?: string;
  permissionRole?: AccessRole;
  previousRole?: AccessRole;
  siteName?: string;
  libraryName?: string;
  itemId?: string;
  source?: PermissionEntry["source"];
  tenantType?: PermissionEntry["tenant"];
  status: AuditLogStatus;
  errorMessage?: string;
  graphRequestId?: string;
};

export type UserSuggestion = {
  id: string;
  displayName: string;
  email: string;
  jobTitle?: string;
};

export type ReportSiteSummary = {
  siteId: string;
  siteName: string;
  hostname: string;
  libraryCount: number;
  protectedLibraryCount: number;
  standardLibraryCount: number;
  directPermissionCount: number;
  externalPermissionCount: number;
  inheritedPermissionCount: number;
};

export type ReportPermissionRow = {
  id: string;
  siteName: string;
  libraryName: string;
  principalName: string;
  email: string;
  role: AccessRole;
  source: PermissionEntry["source"];
  tenant: PermissionEntry["tenant"];
};

export type ReportSummary = {
  generatedAt: string;
  siteCount: number;
  libraryCount: number;
  protectedLibraryCount: number;
  standardLibraryCount: number;
  directPermissionCount: number;
  externalPermissionCount: number;
  inheritedPermissionCount: number;
  sites: ReportSiteSummary[];
  permissions: ReportPermissionRow[];
};
