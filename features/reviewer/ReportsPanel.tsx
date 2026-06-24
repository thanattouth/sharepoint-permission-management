"use client";

import { RefreshCw } from "lucide-react";
import { isDefaultSharePointGroup, roleLabels } from "@/lib/features/admin";
import type { ReviewScopeOwner } from "@/lib/features/reviewer";
import type { PermissionEntry, ReportSummary } from "@/lib/types";
import { ReportSkeleton } from "@/components/shared/Skeletons";

type ReportsPanelProps = {
  report: ReportSummary | null;
  reportError: string;
  loadingLabel: string;
  reviewOwners: ReviewScopeOwner[];
  selectedOwnerEmail: string;
  onRefresh: (ownerEmail?: string) => void;
};

export function ReportsPanel({
  report,
  reportError,
  loadingLabel,
  reviewOwners,
  selectedOwnerEmail,
  onRefresh,
}: ReportsPanelProps) {
  const generatedAt = report?.generatedAt ? new Date(report.generatedAt).toLocaleString() : "Not generated";
  const isLoadingReport = loadingLabel === "Loading reports";
  const reviewerPermissions = (report?.permissions ?? []).filter((permission) => !isDefaultReportSharePointGroup(permission));
  const hiddenSystemPermissionCount = Math.max((report?.permissions.length ?? 0) - reviewerPermissions.length, 0);
  const selectedOwner = reviewOwners.find((owner) => owner.email === selectedOwnerEmail.toLowerCase());
  const hasConfiguredScopes = reviewOwners.length > 0;
  const mappedSiteText = report
    ? `${report.siteCount.toLocaleString()} site${report.siteCount === 1 ? "" : "s"} and ${report.libraryCount.toLocaleString()} librar${report.libraryCount === 1 ? "y" : "ies"}`
    : "mapped sites and libraries";

  return (
    <section className="page-section">
      <div className="page-header with-actions">
        <div>
          <p className="section-label">Reviewer</p>
          <h1>Permission Review</h1>
          <p>Read-only inventory summary across configured SharePoint sites.</p>
        </div>
        <button className="secondary-button" disabled={isLoadingReport} onClick={() => onRefresh(selectedOwnerEmail || undefined)}>
          <RefreshCw className={isLoadingReport ? "spin-icon" : ""} size={17} />
          Refresh review
        </button>
      </div>

      {reportError && <div className="auth-error">{reportError}</div>}

      {hasConfiguredScopes && (
        <div className="review-scope-bar">
          <div className="review-scope-identity">
            <span>Your review scope</span>
            <strong>{selectedOwner?.name ?? selectedOwnerEmail}</strong>
            <small>{selectedOwnerEmail}</small>
          </div>
          <span className="review-scope-note">
            {selectedOwner
              ? `Loaded from your mapped scope: ${mappedSiteText}. ${selectedOwner.scopeCount.toLocaleString()} mapping${selectedOwner.scopeCount === 1 ? "" : "s"} are assigned to you.`
              : "No review scope is mapped to this signed-in account yet. Ask an admin to add this account to the review scope list."}
          </span>
        </div>
      )}

      {isLoadingReport && <ReportSkeleton />}

      {report && !isLoadingReport && (
        <>
          <div className="report-meta">
            <span>Generated</span>
            <strong>{generatedAt}</strong>
            {typeof report.scannedItemCount === "number" && (
              <>
                <span>Scanned items</span>
                <strong>{report.scannedItemCount.toLocaleString()}</strong>
              </>
            )}
            {report.scanLimitReached && <span className="risk-text">Scan limit reached</span>}
          </div>

          <div className="permission-section-title">
            <div>
              <p className="section-label">Site Summary</p>
              <h2>Coverage by site</h2>
            </div>
            <div className="section-title-actions">
              <span>{report.sites.length}</span>
            </div>
          </div>

          <div className="report-table" role="table" aria-label="Site report">
            <div className="report-table-head" role="row">
              <span>Site</span>
              <span>Libraries</span>
              <span>Protected</span>
              <span>Can edit here</span>
              <span>External</span>
              <span>From parent</span>
            </div>
            {report.sites.map((site) => (
              <div className="report-table-row" role="row" key={site.siteId}>
                <div>
                  <strong>{site.siteName}</strong>
                  <small>{site.hostname}</small>
                </div>
                <span>{site.libraryCount}</span>
                <span>{site.protectedLibraryCount}</span>
                <span>{site.directPermissionCount}</span>
                <span className={site.externalPermissionCount > 0 ? "risk-text" : ""}>{site.externalPermissionCount}</span>
                <span>{site.inheritedPermissionCount}</span>
              </div>
            ))}
          </div>

          <div className="permission-section-title">
            <div>
              <p className="section-label">Permission Inventory</p>
              <h2>Who has access</h2>
            </div>
            <div className="section-title-actions">
              <span>{reviewerPermissions.length}</span>
            </div>
          </div>

          <div className="report-permission-table report-permission-table-scroll" role="table" aria-label="Permission inventory">
            <div className="report-permission-head" role="row">
              <span>Principal</span>
              <span>Role</span>
              <span>Item scope</span>
              <span>Tenant</span>
            </div>
            {reviewerPermissions.map((permission) => (
              <div className="report-permission-row" role="row" key={permission.id}>
                <div>
                  <strong>{permission.principalName}</strong>
                  <small>{permission.email}</small>
                </div>
                <span className={`role-chip ${permission.role}`}>{roleLabels[permission.role]}</span>
                <div>
                  <strong>{permission.itemName}</strong>
                  <small>{permission.itemType} / {permission.itemPath}</small>
                </div>
                <span className={permission.tenant === "external" ? "risk-text" : ""}>{permission.tenant}</span>
              </div>
            ))}
            {reviewerPermissions.length === 0 && (
              <div className="empty-row">No non-system permissions found in configured library roots.</div>
            )}
          </div>
          {hiddenSystemPermissionCount > 0 && (
            <p className="table-footnote">
              Hidden {hiddenSystemPermissionCount} default SharePoint Owners, Members, and Visitors group row{hiddenSystemPermissionCount > 1 ? "s" : ""}.
            </p>
          )}
          {report.scanLimitReached && (
            <p className="table-footnote risk-text">
              The review reached the configured item scan limit. Increase NEXT_PUBLIC_REVIEW_SCAN_ITEM_LIMIT to scan more items.
            </p>
          )}
        </>
      )}

      {!report && !reportError && loadingLabel !== "Loading reports" && (
        <div className="empty-row site-empty-state">
          {hasConfiguredScopes
            ? "No mapped review data is available for this signed-in account."
            : "No report data loaded."}
        </div>
      )}
    </section>
  );
}

function isDefaultReportSharePointGroup(permission: ReportSummary["permissions"][number]) {
  return isDefaultSharePointGroup(
    {
      id: permission.id,
      libraryId: permission.id,
      displayName: permission.principalName,
      email: permission.email,
      type: "group",
      role: permission.role,
      source: permission.source,
      tenant: permission.tenant,
      lastActivity: "Report",
    } satisfies PermissionEntry,
    permission.siteName,
  );
}
