import { auditSite, managedSiteListId, managedSiteListName } from "../../app-config";
import type { GraphCollection, GraphRequestClient } from "../../graph-request";
import {
  getManagedSiteKey,
  normalizeManagedSite,
  type ManagedSiteDraft,
  type ManagedSiteStore,
} from "./managed-site-store";

type GraphSite = {
  id: string;
};

type GraphList = {
  id: string;
  displayName?: string;
};

type GraphListItem = {
  id: string;
  fields?: Record<string, unknown>;
};

export class SharePointListManagedSiteStore implements ManagedSiteStore {
  private managedSiteListPromise?: Promise<{ siteId: string; listId: string }>;

  constructor(private readonly graph: GraphRequestClient) {}

  async list() {
    const target = await this.getManagedSiteList();
    const response = await this.graph.request<GraphCollection<GraphListItem>>(
      `/sites/${encodeURIComponent(target.siteId)}/lists/${encodeURIComponent(target.listId)}/items?$top=500&$expand=fields`,
    );

    return (response.value ?? [])
      .map((item) => mapManagedSiteItem(item.fields ?? {}))
      .filter((site): site is ManagedSiteDraft => Boolean(site));
  }

  async add(site: ManagedSiteDraft) {
    const nextSite = normalizeManagedSite(site);
    if (!nextSite) return;

    const target = await this.getManagedSiteList();
    const existing = await this.findItem(target, nextSite);
    const fields = {
      Title: `${nextSite.hostname}${nextSite.path}`,
      Hostname: nextSite.hostname,
      Path: nextSite.path,
      Active: "Yes",
    };

    if (existing) {
      await this.graph.request<GraphListItem>(
        `/sites/${encodeURIComponent(target.siteId)}/lists/${encodeURIComponent(target.listId)}/items/${encodeURIComponent(existing.id)}/fields`,
        {
          method: "PATCH",
          body: JSON.stringify(fields),
        },
      );
      return;
    }

    await this.graph.request<GraphListItem>(
      `/sites/${encodeURIComponent(target.siteId)}/lists/${encodeURIComponent(target.listId)}/items`,
      {
        method: "POST",
        body: JSON.stringify({ fields }),
      },
    );
  }

  async remove(site: ManagedSiteDraft) {
    const target = await this.getManagedSiteList();
    const existing = await this.findItem(target, site);
    if (!existing) return;

    await this.graph.request<GraphListItem>(
      `/sites/${encodeURIComponent(target.siteId)}/lists/${encodeURIComponent(target.listId)}/items/${encodeURIComponent(existing.id)}/fields`,
      {
        method: "PATCH",
        body: JSON.stringify({ Active: "No" }),
      },
    );
  }

  private async findItem(target: { siteId: string; listId: string }, site: ManagedSiteDraft) {
    const siteKey = getManagedSiteKey(normalizeManagedSite(site) ?? site);
    const response = await this.graph.request<GraphCollection<GraphListItem>>(
      `/sites/${encodeURIComponent(target.siteId)}/lists/${encodeURIComponent(target.listId)}/items?$top=500&$expand=fields`,
    );

    return (response.value ?? []).find((item) => {
      const mapped = mapManagedSiteItem(item.fields ?? {}, { includeInactive: true });
      return mapped && getManagedSiteKey(mapped) === siteKey;
    });
  }

  private async getManagedSiteList() {
    this.managedSiteListPromise ??= this.resolveManagedSiteList();
    return this.managedSiteListPromise;
  }

  private async resolveManagedSiteList() {
    const site = await this.graph.request<GraphSite>(
      `/sites/${auditSite.hostname}:${encodeURI(auditSite.path)}?$select=id`,
    );
    if (managedSiteListId) {
      return { siteId: site.id, listId: managedSiteListId };
    }

    const directList = await this.resolveListByConvention(site.id).catch(() => undefined);
    if (directList) {
      return { siteId: site.id, listId: directList.id };
    }

    const lists = await this.graph.request<GraphCollection<GraphList>>(
      `/sites/${encodeURIComponent(site.id)}/lists?$select=id,displayName`,
    );
    const existing = (lists.value ?? []).find((list) => list.displayName === managedSiteListName);
    if (existing) return { siteId: site.id, listId: existing.id };

    const created = await this.graph.request<GraphList>(`/sites/${encodeURIComponent(site.id)}/lists`, {
      method: "POST",
      body: JSON.stringify({
        displayName: managedSiteListName,
        columns: managedSiteColumns.map((name) => ({ name, text: {} })),
        list: { template: "genericList" },
      }),
    });

    return { siteId: site.id, listId: created.id };
  }

  private async resolveListByConvention(siteId: string) {
    return this.graph.request<GraphList>(
      `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(managedSiteListName)}?$select=id,displayName`,
    );
  }
}

const managedSiteColumns = ["Hostname", "Path", "Active"];

function mapManagedSiteItem(fields: Record<string, unknown>, options: { includeInactive?: boolean } = {}) {
  const active = readText(fields.Active);
  if (!options.includeInactive && active && active.toLowerCase() !== "yes") return undefined;
  return normalizeManagedSite({
    hostname: readText(fields.Hostname),
    path: readText(fields.Path),
  });
}

function readText(value: unknown) {
  return typeof value === "string" ? value : "";
}
