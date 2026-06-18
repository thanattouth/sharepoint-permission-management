import { reviewScopeListName, reviewScopeSite } from "../../app-config";
import type { GraphCollection, GraphRequestClient } from "../../graph-request";
import { normalizeReviewScope, type ReviewScope } from "./review-scopes";
import type { ReviewScopeStore } from "./review-scope-store";

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

export class SharePointListReviewScopeStore implements ReviewScopeStore {
  private reviewScopeListPromise?: Promise<{ siteId: string; listId: string }>;

  constructor(private readonly graph: GraphRequestClient) {}

  async list(): Promise<ReviewScope[]> {
    const target = await this.getReviewScopeList();
    const response = await this.graph.request<GraphCollection<GraphListItem>>(
      `/sites/${encodeURIComponent(target.siteId)}/lists/${encodeURIComponent(target.listId)}/items?$top=500&$expand=fields`,
    );

    return (response.value ?? [])
      .map((item) => normalizeReviewScope(item.fields ?? {}))
      .filter((scope): scope is ReviewScope => Boolean(scope));
  }

  private async getReviewScopeList() {
    this.reviewScopeListPromise ??= this.resolveReviewScopeList();
    return this.reviewScopeListPromise;
  }

  private async resolveReviewScopeList() {
    const site = await this.graph.request<GraphSite>(
      `/sites/${reviewScopeSite.hostname}:${encodeURI(reviewScopeSite.path)}?$select=id`,
    );
    const lists = await this.graph.request<GraphCollection<GraphList>>(
      `/sites/${encodeURIComponent(site.id)}/lists?$select=id,displayName`,
    );
    const existing = (lists.value ?? []).find((list) => list.displayName === reviewScopeListName);

    if (!existing) {
      throw new Error(`Review scope list "${reviewScopeListName}" was not found.`);
    }

    return { siteId: site.id, listId: existing.id };
  }
}
