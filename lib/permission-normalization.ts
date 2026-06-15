import type { AccessRole, PermissionEntry } from "./types";

export function normalizeSharePointPrincipals(permissions: PermissionEntry[], siteName: string) {
  return permissions.map((permission) => {
    if (!isOpaqueSharePointPrincipal(permission.displayName)) {
      if (isDefaultSharePointGroup(permission, siteName)) {
        return lockManagedPermission(permission, permission.source === "inherited"
          ? "Inherited from parent"
          : "System-managed SharePoint group");
      }

      return permission;
    }

    const decoded = decodeSharePointPrincipal(permission.displayName);
    const principalId = decoded?.split("_").at(-1);
    const groupLabel = getDefaultSharePointGroupName(permission.role, principalId);

    return {
      ...permission,
      displayName: `${siteName} ${groupLabel}`,
      email: principalId ? `Inherited SharePoint group - Principal ${principalId}` : "Inherited SharePoint group",
      type: "group" as const,
      source: "inherited" as const,
      lastActivity: "Inherited from parent",
      canEditRole: false,
      canDelete: false,
    };
  });
}

function lockManagedPermission(permission: PermissionEntry, lastActivity: string): PermissionEntry {
  return {
    ...permission,
    lastActivity,
    canEditRole: false,
    canDelete: false,
  };
}

function isDefaultSharePointGroup(permission: PermissionEntry, siteName: string) {
  if (permission.type !== "group") return false;
  const normalizedName = permission.displayName.toLowerCase();
  const normalizedSite = siteName.toLowerCase();
  return (
    normalizedName === `${normalizedSite} owners` ||
    normalizedName === `${normalizedSite} members` ||
    normalizedName === `${normalizedSite} visitors`
  );
}

function isOpaqueSharePointPrincipal(value: string) {
  if (value.length < 28 || /\s/.test(value)) return false;
  return /^[A-Za-z0-9_-]+={0,2}$/.test(value);
}

function decodeSharePointPrincipal(value: string) {
  try {
    const normalized = value.replace(/-/g, "+").replace(/_/g, "/");
    const padded = normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "=");
    const decoded = atob(padded).replace(/[^\x20-\x7E_:-]/g, "");
    return decoded.includes("_") ? decoded : undefined;
  } catch {
    return undefined;
  }
}

function getDefaultSharePointGroupName(role: AccessRole, principalId?: string) {
  if (principalId === "3") return "Owners";
  if (principalId === "4") return "Visitors";
  if (principalId === "5") return "Members";
  if (role === "owner") return "Owners";
  if (role === "editor") return "Members";
  return "Visitors";
}
