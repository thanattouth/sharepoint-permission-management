export type ManagedSiteDraft = {
  hostname: string;
  path: string;
};

export interface ManagedSiteStore {
  list(): Promise<ManagedSiteDraft[]>;
  add(site: ManagedSiteDraft): Promise<void>;
  remove(site: ManagedSiteDraft): Promise<void>;
}

export class LocalManagedSiteStore implements ManagedSiteStore {
  private readonly storageKey = "spAccess:customSites";

  async list() {
    return readLocalSites(this.storageKey);
  }

  async add(site: ManagedSiteDraft) {
    const nextSite = normalizeManagedSite(site);
    if (!nextSite) return;
    const current = readLocalSites(this.storageKey);
    window.localStorage.setItem(this.storageKey, JSON.stringify(upsertManagedSite(current, nextSite)));
  }

  async remove(site: ManagedSiteDraft) {
    const targetKey = getManagedSiteKey(site);
    const current = readLocalSites(this.storageKey);
    window.localStorage.setItem(
      this.storageKey,
      JSON.stringify(current.filter((candidate) => getManagedSiteKey(candidate) !== targetKey)),
    );
  }
}

export function normalizeManagedSite(site: ManagedSiteDraft | undefined): ManagedSiteDraft | undefined {
  const hostname = site?.hostname.trim().toLowerCase().replace(/^https?:\/\//, "").split("/")[0];
  const rawPath = site?.path.trim() ?? "";
  if (!hostname || !rawPath) return undefined;
  const path = rawPath.startsWith("/") ? rawPath : `/${rawPath}`;
  return { hostname, path };
}

export function upsertManagedSite(sites: ManagedSiteDraft[], site: ManagedSiteDraft) {
  const normalizedSite = normalizeManagedSite(site);
  if (!normalizedSite) return sites;
  const siteKey = getManagedSiteKey(normalizedSite);
  return [
    ...sites.filter((candidate) => getManagedSiteKey(candidate) !== siteKey),
    normalizedSite,
  ];
}

export function getManagedSiteKey(site: ManagedSiteDraft) {
  return `${site.hostname.trim().toLowerCase()}:${site.path.trim().toLowerCase()}`;
}

function readLocalSites(storageKey: string) {
  try {
    const raw = window.localStorage.getItem(storageKey);
    const parsed = raw ? JSON.parse(raw) as ManagedSiteDraft[] : [];
    return Array.isArray(parsed)
      ? parsed.map(normalizeManagedSite).filter((site): site is ManagedSiteDraft => Boolean(site))
      : [];
  } catch {
    return [];
  }
}
