import type { ReportSummary } from "../types";
import {
  GraphSharePointAdminClient,
  GraphInviteFailureError,
  graphReadScopes,
  graphWriteScopes,
  type GrantPermissionResult,
  type InviteDiagnostic,
  type PermissionDraft,
  type SharePointPermissionClient as SharePointAdminClient,
} from "./admin";
import {
  GraphSharePointReportClient,
  type SharePointReportClient,
  type ReviewScope,
} from "./reviewer";
import type { TokenProvider } from "../graph-request";

export {
  GraphInviteFailureError,
  graphReadScopes,
  graphWriteScopes,
  type GrantPermissionResult,
  type InviteDiagnostic,
  type PermissionDraft,
};

export interface SharePointPermissionClient extends SharePointAdminClient, SharePointReportClient {}

export class GraphSharePointPermissionClient extends GraphSharePointAdminClient implements SharePointPermissionClient {
  private readonly reportClient: GraphSharePointReportClient;

  constructor(getToken: TokenProvider) {
    super(getToken);
    this.reportClient = new GraphSharePointReportClient(this);
  }

  getReportSummary(options?: { ownerEmail?: string; reviewScopes?: ReviewScope[] }): Promise<ReportSummary> {
    return this.reportClient.getReportSummary(options);
  }
}
