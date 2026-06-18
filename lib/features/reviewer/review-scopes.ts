export type ReviewOwnerRole = "OwnerRep" | "VP" | "EVP" | "Reviewer";

export type ReviewScope = {
  ownerEmail: string;
  ownerName: string;
  ownerRole?: ReviewOwnerRole;
  siteName?: string;
  hostname?: string;
  path?: string;
  libraryName?: string;
  sensitivityLabel?: "Confidential" | "Secret" | string;
  department?: string;
  section?: string;
  active?: boolean;
};

export type ReviewScopeOwner = {
  email: string;
  name: string;
  role: ReviewOwnerRole;
  scopeCount: number;
};

export const fallbackReviewScopes = parseReviewScopes(process.env.NEXT_PUBLIC_REVIEW_SCOPES);
export const fallbackReviewScopeOwners = getReviewScopeOwners(fallbackReviewScopes);

export function getReviewScopesForOwner(ownerEmail: string | undefined, scopes = fallbackReviewScopes) {
  const normalizedOwner = normalizeEmail(ownerEmail);
  if (!normalizedOwner) return [];

  return scopes.filter((scope) => scope.active !== false && normalizeEmail(scope.ownerEmail) === normalizedOwner);
}

export function hasReviewScopes(scopes = fallbackReviewScopes) {
  return scopes.some((scope) => scope.active !== false);
}

export function matchesReviewScopeSite(scope: ReviewScope, site: { name: string; hostname: string; path: string }) {
  return (
    matchesOptional(scope.siteName, site.name) &&
    matchesOptional(scope.hostname, site.hostname) &&
    matchesOptional(scope.path, site.path)
  );
}

export function matchesReviewScopeLibrary(scope: ReviewScope, library: { name: string }) {
  return matchesOptional(scope.libraryName, library.name);
}

export function getReviewScopeOwners(scopes: ReviewScope[]) {
  const owners = new Map<string, ReviewScopeOwner>();

  scopes
    .filter((scope) => scope.active !== false)
    .forEach((scope) => {
      const email = normalizeEmail(scope.ownerEmail);
      if (!email) return;

      const existing = owners.get(email);
      if (existing) {
        existing.scopeCount += 1;
        return;
      }

      owners.set(email, {
        email,
        name: scope.ownerName || email,
        role: scope.ownerRole ?? "Reviewer",
        scopeCount: 1,
      });
    });

  return Array.from(owners.values()).sort((left, right) =>
    left.name.localeCompare(right.name) || left.email.localeCompare(right.email),
  );
}

export function parseReviewScopes(value: string | undefined): ReviewScope[] {
  if (!value?.trim()) return [];

  try {
    const parsed = JSON.parse(value) as unknown;
    if (!Array.isArray(parsed)) return [];

    return parsed
      .map((item) => normalizeReviewScope(item))
      .filter((scope): scope is ReviewScope => Boolean(scope));
  } catch {
    return [];
  }
}

export function normalizeReviewScope(value: unknown): ReviewScope | undefined {
  if (!value || typeof value !== "object") return undefined;
  const item = value as Record<string, unknown>;
  const ownerEmail = readText(item.ownerEmail ?? item.OwnerEmail);
  const ownerName = readText(item.ownerName ?? item.OwnerName ?? item.Title) || ownerEmail;
  if (!ownerEmail || !ownerName) return undefined;

  return {
    ownerEmail,
    ownerName,
    ownerRole: readOwnerRole(item.ownerRole ?? item.OwnerRole),
    siteName: readOptionalText(item.siteName ?? item.SiteName),
    hostname: readOptionalText(item.hostname ?? item.Hostname),
    path: readOptionalText(item.path ?? item.Path),
    libraryName: readOptionalText(item.libraryName ?? item.LibraryName),
    sensitivityLabel: readOptionalText(item.sensitivityLabel ?? item.SensitivityLabel),
    department: readOptionalText(item.department ?? item.Department),
    section: readOptionalText(item.section ?? item.Section),
    active: readActive(item.active ?? item.Active),
  };
}

function readText(value: unknown) {
  return typeof value === "string" ? value.trim() : "";
}

function readOptionalText(value: unknown) {
  const text = readText(value);
  return text || undefined;
}

function readOwnerRole(value: unknown): ReviewOwnerRole | undefined {
  if (value === "OwnerRep" || value === "VP" || value === "EVP" || value === "Reviewer") return value;
  return undefined;
}

function readActive(value: unknown) {
  if (typeof value === "boolean") return value;
  if (typeof value !== "string") return true;

  const normalized = value.trim().toLowerCase();
  if (["false", "no", "0", "inactive"].includes(normalized)) return false;
  return true;
}

function normalizeEmail(value: string | undefined) {
  return value?.trim().toLowerCase() ?? "";
}

function normalizeMatch(value: string) {
  return value.trim().toLowerCase();
}

function matchesOptional(expected: string | undefined, actual: string) {
  return !expected || normalizeMatch(expected) === normalizeMatch(actual);
}
