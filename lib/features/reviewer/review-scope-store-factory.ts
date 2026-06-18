import { reviewScopeListEnabled } from "../../app-config";
import { GraphRequestClient, type TokenProvider } from "../../graph-request";
import { EnvReviewScopeStore, type ReviewScopeStore } from "./review-scope-store";
import { SharePointListReviewScopeStore } from "./sharepoint-list-review-scope-store";

export function createReviewScopeStore(getToken: TokenProvider): ReviewScopeStore {
  if (!reviewScopeListEnabled) {
    return new EnvReviewScopeStore();
  }

  return new SharePointListReviewScopeStore(new GraphRequestClient(getToken));
}
