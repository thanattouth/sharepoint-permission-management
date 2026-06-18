import { reviewScanDescendantsEnabled, reviewScanItemLimit } from "../../app-config";
import type {
  ContentItem,
  LibrarySummary,
  PermissionEntry,
  ReportPermissionRow,
  ReportSiteSummary,
  ReportSummary,
  SiteSummary,
} from "../../types";
import { normalizeSharePointPrincipals } from "../admin/permission-normalization";
import { libraryToContentItem, type SharePointInventoryReader } from "../admin/sharepoint-permission-client";
import {
  fallbackReviewScopes,
  getReviewScopesForOwner,
  hasReviewScopes,
  matchesReviewScopeLibrary,
  matchesReviewScopeSite,
  type ReviewScope,
} from "./review-scopes";

type ReportPermissionInventoryEntry = {
  item: ContentItem;
  itemPath: string;
  permissions: PermissionEntry[];
};

export interface SharePointReportClient {
  getReportSummary(options?: { ownerEmail?: string; reviewScopes?: ReviewScope[] }): Promise<ReportSummary>;
}

export class GraphSharePointReportClient implements SharePointReportClient {
  constructor(private readonly inventory: SharePointInventoryReader) {}

  async getReportSummary(options: { ownerEmail?: string; reviewScopes?: ReviewScope[] } = {}) {
    const configuredReviewScopes = options.reviewScopes ?? fallbackReviewScopes;
    const ownerScopes = getReviewScopesForOwner(options.ownerEmail, configuredReviewScopes);
    const scopeConfigured = hasReviewScopes(configuredReviewScopes);
    const scopeByOwner = ownerScopes.length > 0;
    if (scopeConfigured && !scopeByOwner) {
      return emptyReportSummary(options.ownerEmail, scopeConfigured);
    }

    const sites = (await this.inventory.listSites()).filter((site) => {
      if (!scopeByOwner) return true;
      return ownerScopes.some((scope) => matchesReviewScopeSite(scope, site));
    });
    const reportSites = await Promise.all(
      sites.map(async (site) => {
        const libraries = (await this.inventory.listLibraries(site.id)).filter((library) => {
          if (!scopeByOwner) return true;
          const siteScopes = ownerScopes.filter((scope) => matchesReviewScopeSite(scope, site));
          return siteScopes.some((scope) => matchesReviewScopeLibrary(scope, library));
        });
        const permissionInventories = await Promise.all(
          libraries.map(async (library) => ({
            library,
            inventory: scopeByOwner && reviewScanDescendantsEnabled
              ? await this.scanLibraryPermissionInventory(site, library)
              : await this.scanLibraryRootPermissionInventory(site, library),
          })),
        );
        const permissions = permissionInventories.flatMap((entry) =>
          entry.inventory.entries.flatMap((inventoryEntry) => inventoryEntry.permissions),
        );
        return {
          summary: mapReportSite(site, libraries, permissions),
          scannedItemCount: sumByNumber(permissionInventories, (entry) => entry.inventory.scannedItemCount),
          scanLimitReached: permissionInventories.some((entry) => entry.inventory.scanLimitReached),
          permissions: permissionInventories.flatMap((entry) =>
            entry.inventory.entries.flatMap((inventoryEntry) =>
              inventoryEntry.permissions.map((permission) =>
                mapReportPermission(site, entry.library, inventoryEntry.item, inventoryEntry.itemPath, permission),
              ),
            ),
          ),
        };
      }),
    );
    const siteSummaries = reportSites.map((site) => site.summary);
    const permissionRows = reportSites.flatMap((site) => site.permissions);

    return {
      generatedAt: new Date().toISOString(),
      reviewOwnerEmail: options.ownerEmail,
      reviewScopeApplied: scopeByOwner,
      reviewScopeConfigured: scopeConfigured,
      siteCount: siteSummaries.length,
      libraryCount: sumBy(siteSummaries, "libraryCount"),
      protectedLibraryCount: sumBy(siteSummaries, "protectedLibraryCount"),
      standardLibraryCount: sumBy(siteSummaries, "standardLibraryCount"),
      directPermissionCount: sumBy(siteSummaries, "directPermissionCount"),
      externalPermissionCount: sumBy(siteSummaries, "externalPermissionCount"),
      inheritedPermissionCount: sumBy(siteSummaries, "inheritedPermissionCount"),
      scannedItemCount: sumByNumber(reportSites, (site) => site.scannedItemCount),
      scanLimitReached: reportSites.some((site) => site.scanLimitReached),
      sites: siteSummaries,
      permissions: permissionRows,
    };
  }

  private async scanLibraryRootPermissionInventory(site: SiteSummary, library: LibrarySummary) {
    const rootItem = libraryToContentItem(library);
    const permissions = normalizeSharePointPrincipals(
      await this.inventory.listPermissions(rootItem.id).catch(() => []),
      site.name,
    );

    return {
      entries: [{ item: rootItem, itemPath: library.name, permissions }],
      scannedItemCount: 1,
      scanLimitReached: false,
    };
  }

  private async scanLibraryPermissionInventory(site: SiteSummary, library: LibrarySummary) {
    const rootItem = libraryToContentItem(library);
    const entries: ReportPermissionInventoryEntry[] = [];
    const queue: Array<{ item: ContentItem; itemPath: string }> = [{ item: rootItem, itemPath: library.name }];
    let scannedItemCount = 0;

    while (queue.length > 0 && scannedItemCount < reviewScanItemLimit) {
      const current = queue.shift();
      if (!current) break;

      scannedItemCount += 1;
      const permissions = normalizeSharePointPrincipals(
        await this.inventory.listPermissions(current.item.id).catch(() => []),
        site.name,
      );
      entries.push({ ...current, permissions });

      if (current.item.type === "file" || scannedItemCount >= reviewScanItemLimit) {
        continue;
      }

      const children = await this.inventory.listChildren(current.item).catch(() => []);
      children.forEach((child) => {
        queue.push({
          item: child,
          itemPath: `${current.itemPath}/${child.name}`,
        });
      });
    }

    return {
      entries,
      scannedItemCount,
      scanLimitReached: queue.length > 0,
    };
  }
}

function emptyReportSummary(ownerEmail: string | undefined, scopeConfigured: boolean): ReportSummary {
  return {
    generatedAt: new Date().toISOString(),
    reviewOwnerEmail: ownerEmail,
    reviewScopeApplied: false,
    reviewScopeConfigured: scopeConfigured,
    siteCount: 0,
    libraryCount: 0,
    protectedLibraryCount: 0,
    standardLibraryCount: 0,
    directPermissionCount: 0,
    externalPermissionCount: 0,
    inheritedPermissionCount: 0,
    scannedItemCount: 0,
    scanLimitReached: false,
    sites: [],
    permissions: [],
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
  item: ContentItem,
  itemPath: string,
  permission: PermissionEntry,
): ReportPermissionRow {
  return {
    id: `${site.id}:${library.id}:${item.itemId ?? item.id}:${permission.id}`,
    siteName: site.name,
    libraryName: library.name,
    itemName: item.name,
    itemPath,
    itemType: item.type,
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

function sumByNumber<T>(items: T[], getValue: (item: T) => number | undefined) {
  return items.reduce((total, item) => total + (getValue(item) ?? 0), 0);
}
