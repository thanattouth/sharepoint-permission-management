import { fallbackReviewScopes, type ReviewScope } from "./review-scopes";

export interface ReviewScopeStore {
  list(): Promise<ReviewScope[]>;
}

export class EnvReviewScopeStore implements ReviewScopeStore {
  async list() {
    return fallbackReviewScopes;
  }
}
