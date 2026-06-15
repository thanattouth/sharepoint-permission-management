export type TargetSiteConfig = {
  hostname: string;
  path: string;
};

const defaultTenantDomain = "baht.net";
const defaultProtectedLibraries = ["Confidential", "Secret"];
const defaultTargetSites: TargetSiteConfig[] = [
  {
    hostname: "bahtnet.sharepoint.com",
    path: "/sites/DGCS",
  },
  {
    hostname: "bahtnet.sharepoint.com",
    path: "/sites/EngineerSite",
  },
];

export const tenantDomain = normalizeDomain(
  process.env.NEXT_PUBLIC_TENANT_DOMAIN ?? defaultTenantDomain,
);
export const internalDomains = new Set(
  parseCsv(process.env.NEXT_PUBLIC_INTERNAL_DOMAINS, [tenantDomain]).map(normalizeDomain),
);

export const protectedLibraryNames = new Set(
  parseCsv(process.env.NEXT_PUBLIC_PROTECTED_LIBRARY_NAMES, defaultProtectedLibraries),
);

export const targetSites = parseTargetSites(process.env.NEXT_PUBLIC_TARGET_SITES) ?? defaultTargetSites;

export function isInternalEmail(value: string) {
  const normalized = value.trim().toLowerCase();
  return Array.from(internalDomains).some((domain) => normalized.endsWith(`@${domain}`));
}

function normalizeDomain(value: string) {
  return value.trim().toLowerCase().replace(/^@/, "");
}

function parseCsv(value: string | undefined, fallback: string[]) {
  const parsed = value
    ?.split(",")
    .map((item) => item.trim())
    .filter(Boolean);

  return parsed?.length ? parsed : fallback;
}

function parseTargetSites(value: string | undefined): TargetSiteConfig[] | undefined {
  if (!value?.trim()) return undefined;

  const sites = value
    .split(",")
    .map((entry) => entry.trim())
    .filter(Boolean)
    .map((entry) => {
      const [hostname, path] = entry.split(":");
      if (!hostname?.trim() || !path?.trim().startsWith("/")) return undefined;

      return {
        hostname: hostname.trim(),
        path: path.trim(),
      };
    })
    .filter((site): site is TargetSiteConfig => Boolean(site));

  return sites.length ? sites : undefined;
}
