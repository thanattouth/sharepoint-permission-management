import type { AccountInfo } from "@azure/msal-browser";
import type { ContentItem } from "./types";

export type AppRole = "Admin" | "InternalUser" | "GuestUser" | "ExecutiveUser";

export type UserCapabilities = {
  canManagePermissions: boolean;
  canViewAllSites: boolean;
  canViewReports: boolean;
  isReadOnly: boolean;
};

const roleLabels: Record<AppRole, string> = {
  Admin: "Admin",
  InternalUser: "Internal User",
  GuestUser: "Guest User",
  ExecutiveUser: "Executive User",
};

const defaultCapabilities: UserCapabilities = {
  canManagePermissions: false,
  canViewAllSites: false,
  canViewReports: false,
  isReadOnly: true,
};

export function getAccountRoles(account: AccountInfo | null): AppRole[] {
  const rawRoles = getRawRoles(account);
  return rawRoles.filter(isAppRole);
}

export function getPrimaryRole(roles: AppRole[]) {
  return roles[0];
}

export function getRoleLabel(role: AppRole | undefined) {
  return role ? roleLabels[role] : "No app role";
}

export function getCapabilities(roles: AppRole[]): UserCapabilities {
  if (roles.includes("Admin")) {
    return {
      canManagePermissions: true,
      canViewAllSites: true,
      canViewReports: true,
      isReadOnly: false,
    };
  }

  if (roles.includes("ExecutiveUser")) {
    return {
      canManagePermissions: false,
      canViewAllSites: true,
      canViewReports: true,
      isReadOnly: true,
    };
  }

  if (roles.includes("InternalUser")) {
    return {
      canManagePermissions: false,
      canViewAllSites: false,
      canViewReports: false,
      isReadOnly: true,
    };
  }

  if (roles.includes("GuestUser")) {
    return {
      canManagePermissions: false,
      canViewAllSites: false,
      canViewReports: false,
      isReadOnly: true,
    };
  }

  return defaultCapabilities;
}

export function filterContentItemsForRoles(items: ContentItem[], roles: AppRole[]) {
  if (roles.includes("InternalUser")) {
    return items.filter((item) => !item.protected && item.rightsPolicy === "Standard");
  }

  if (roles.includes("GuestUser")) {
    return [];
  }

  return items;
}

function getRawRoles(account: AccountInfo | null) {
  const claims = account?.idTokenClaims as { roles?: unknown } | undefined;
  return Array.isArray(claims?.roles) ? claims.roles : [];
}

function isAppRole(value: unknown): value is AppRole {
  return (
    value === "Admin" ||
    value === "InternalUser" ||
    value === "GuestUser" ||
    value === "ExecutiveUser"
  );
}
