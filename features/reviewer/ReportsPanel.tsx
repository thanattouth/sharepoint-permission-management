"use client";

import { RefreshCw } from "lucide-react";
import { useState } from "react";
import { isDefaultSharePointGroup } from "@/lib/permission-normalization";
import { roleLabels } from "@/lib/permission-labels";
import type { PermissionEntry, ReportSummary } from "@/lib/types";
import { ReportSkeleton } from "@/components/shared/Skeletons";

type ReportsPanelProps = {
  report: ReportSummary | null;
  reportError: string;
  loadingLabel: string;
  onRefresh: () => void;
};

export function ReportsPanel({
  report,
  reportError,
  loadingLabel,
  onRefresh,
}: ReportsPanelProps) {
  const [showAllPermissions, setShowAllPermissions] = useState(false);
  const generatedAt = report?.generatedAt ? new Date(report.generatedAt).toLocaleString() : "Not generated";
  const isLoadingReport = loadingLabel === "Loading reports";
  const reviewerPermissions = (report?.permissions ?? []).filter((permission) => !isDefaultReportSharePointGroup(permission));
  const hiddenSystemPermissionCount = Math.max((report?.permissions.length ?? 0) - reviewerPermissions.length, 0);
  const visibleReportPermissions = showAllPermissions ? reviewerPermissions : reviewerPermissions.slice(0, 8);
  const hiddenPermissionCount = Math.max(reviewerPermissions.length - visibleReportPermissions.length, 0);

  return (
    <section className="page-section">
      <div className="page-header with-actions">
        <div>
          <p className="section-label">Reviewer</p>
          <h1>Permission Review</h1>
          <p>Read-only inventory summary across configured SharePoint sites.</p>
        </div>
        <button className="secondary-button" disabled={isLoadingReport} onClick={onRefresh}>
          <RefreshCw className={isLoadingReport ? "spin-icon" : ""} size={17} />
          Refresh review
        </button>
      </div>

      {reportError && <div className="auth-error">{reportError}</div>}
      {isLoadingReport && <ReportSkeleton />}

      {report && !isLoadingReport && (
        <>
          <div className="report-meta">
            <span>Generated</span>
            <strong>{generatedAt}</strong>
          </div>

          <div className="report-metrics">
            <ReportMetric label="Sites" value={report.siteCount} />
            <ReportMetric label="Libraries" value={report.libraryCount} />
            <ReportMetric label="Protected" value={report.protectedLibraryCount} />
            <ReportMetric label="Standard" value={report.standardLibraryCount} />
            <ReportMetric label="Editable access" value={report.directPermissionCount} />
            <ReportMetric label="External access" value={report.externalPermissionCount} tone="risk" />
            <ReportMetric label="Inherited" value={report.inheritedPermissionCount} />
            <ReportMetric label="Permission rows" value={reviewerPermissions.length} />
          </div>

          <div className="permission-section-title">
            <div>
              <p className="section-label">Permission Inventory</p>
              <h2>Who has access</h2>
            </div>
            <div className="section-title-actions">
              <span>{reviewerPermissions.length}</span>
              {reviewerPermissions.length > 8 && (
                <button className="text-button" type="button" onClick={() => setShowAllPermissions((current) => !current)}>
                  {showAllPermissions ? "Show less" : "See all"}
                </button>
              )}
            </div>
          </div>

          <div className="report-permission-table" role="table" aria-label="Permission inventory">
            <div className="report-permission-head" role="row">
              <span>Principal</span>
              <span>Role</span>
              <span>Source</span>
              <span>Scope</span>
              <span>Tenant</span>
            </div>
            {visibleReportPermissions.map((permission) => (
              <div className="report-permission-row" role="row" key={permission.id}>
                <div>
                  <strong>{permission.principalName}</strong>
                  <small>{permission.email}</small>
                </div>
                <span className={`role-chip ${permission.role}`}>{roleLabels[permission.role]}</span>
                <span>{permission.source}</span>
                <div>
                  <strong>{permission.libraryName}</strong>
                  <small>{permission.siteName}</small>
                </div>
                <span className={permission.tenant === "external" ? "risk-text" : ""}>{permission.tenant}</span>
              </div>
            ))}
            {reviewerPermissions.length === 0 && (
              <div className="empty-row">No non-system permissions found in configured library roots.</div>
            )}
          </div>
          {hiddenPermissionCount > 0 && (
            <p className="table-footnote">
              Showing {visibleReportPermissions.length} of {reviewerPermissions.length}. Use See all for the full inventory.
            </p>
          )}
          {hiddenSystemPermissionCount > 0 && (
            <p className="table-footnote">
              Hidden {hiddenSystemPermissionCount} default SharePoint Owners, Members, and Visitors group row{hiddenSystemPermissionCount > 1 ? "s" : ""}.
            </p>
          )}

          <div className="permission-section-title">
            <div>
              <p className="section-label">Site Summary</p>
              <h2>Coverage by site</h2>
            </div>
            <span>{report.sites.length}</span>
          </div>

          <div className="report-table" role="table" aria-label="Site report">
            <div className="report-table-head" role="row">
              <span>Site</span>
              <span>Libraries</span>
              <span>Protected</span>
              <span>Direct</span>
              <span>External</span>
              <span>Inherited</span>
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
        </>
      )}

      {!report && !reportError && loadingLabel !== "Loading reports" && (
        <div className="empty-row site-empty-state">No report data loaded.</div>
      )}
    </section>
  );
}

function ReportMetric({
  label,
  value,
  tone,
}: {
  label: string;
  value: number;
  tone?: "risk";
}) {
  return (
    <div className={`report-metric ${tone ?? ""}`}>
      <strong>{value}</strong>
      <span>{label}</span>
    </div>
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
