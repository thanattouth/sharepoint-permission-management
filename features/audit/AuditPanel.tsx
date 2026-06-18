"use client";

import { RefreshCw, Search } from "lucide-react";
import { useState } from "react";
import { roleLabels } from "@/lib/features/admin";
import type { AccessRole, AuditLogAction, AuditLogRecord, AuditLogStatus } from "@/lib/types";
import { TableSkeleton } from "@/components/shared/Skeletons";

type AuditPanelProps = {
  auditRecords: AuditLogRecord[];
  auditError: string;
  loadingLabel: string;
  onRefresh: () => void;
};

export function AuditPanel({
  auditRecords,
  auditError,
  loadingLabel,
  onRefresh,
}: AuditPanelProps) {
  const isLoadingAudit = loadingLabel === "Loading audit";
  const [auditQuery, setAuditQuery] = useState("");
  const [auditActionFilter, setAuditActionFilter] = useState<"all" | AuditLogAction>("all");
  const [auditStatusFilter, setAuditStatusFilter] = useState<"all" | AuditLogStatus>("all");
  const filteredAuditRecords = auditRecords.filter((entry) => {
    const search = auditQuery.trim().toLowerCase();
    const matchesSearch = !search || [
      entry.action,
      entry.actorEmail,
      entry.actorName,
      entry.actorRole,
      entry.approvalRequestNo,
      entry.targetEmail,
      entry.targetName,
      entry.permissionRole,
      entry.previousRole,
      entry.siteName,
      entry.libraryName,
      entry.status,
      entry.errorMessage,
      entry.graphRequestId,
    ]
      .filter(Boolean)
      .some((value) => String(value).toLowerCase().includes(search));

    const matchesAction = auditActionFilter === "all" || entry.action === auditActionFilter;
    const matchesStatus = auditStatusFilter === "all" || entry.status === auditStatusFilter;
    return matchesSearch && matchesAction && matchesStatus;
  });

  return (
    <section className="page-section audit-page-section">
      <div className="page-header with-actions">
        <div>
          <p className="section-label">Audit</p>
          <h1>Permission Audit Trail</h1>
          <p>Track who changed permissions, what changed, where it changed, and the approved request number used as reference.</p>
        </div>
        <button className="secondary-button" disabled={isLoadingAudit} onClick={onRefresh}>
          <RefreshCw className={isLoadingAudit ? "spin-icon" : ""} size={17} />
          Refresh audit
        </button>
      </div>

      {auditError && <div className="auth-error">{auditError}</div>}

      <div className="audit-search-bar">
        <label className="search-box audit-search-box">
          <Search size={16} />
          <input
            aria-label="Search audit logs"
            onChange={(event) => setAuditQuery(event.target.value)}
            placeholder="Search actor, target, request no., site, error"
            value={auditQuery}
          />
        </label>
        <select
          aria-label="Filter audit action"
          className="audit-filter-select"
          onChange={(event) => setAuditActionFilter(event.target.value as "all" | AuditLogAction)}
          value={auditActionFilter}
        >
          <option value="all">All actions</option>
          <option value="GrantAccess">Grant</option>
          <option value="UpdateRole">Update role</option>
          <option value="RemoveAccess">Remove</option>
          <option value="RefreshReport">Refresh review</option>
          <option value="Login">Login</option>
        </select>
        <select
          aria-label="Filter audit status"
          className="audit-filter-select"
          onChange={(event) => setAuditStatusFilter(event.target.value as "all" | AuditLogStatus)}
          value={auditStatusFilter}
        >
          <option value="all">All status</option>
          <option value="Success">Success</option>
          <option value="Failed">Failed</option>
        </select>
        <span className="audit-result-count">
          {filteredAuditRecords.length} / {auditRecords.length}
        </span>
      </div>

      {isLoadingAudit ? (
        <div className="audit-table audit-table-scroll">
          <TableSkeleton columns={7} rows={6} />
        </div>
      ) : (
        <div className="audit-table audit-table-scroll" role="table" aria-label="Permission audit trail">
          <div className="audit-table-head" role="row">
            <span>Time</span>
            <span>Action</span>
            <span>Actor</span>
            <span>Target</span>
            <span>Scope</span>
            <span>Request no.</span>
            <span>Status</span>
          </div>
          {filteredAuditRecords.map((entry) => (
            <div className="audit-table-row" role="row" key={entry.id}>
              <span>{formatAuditTime(entry.createdAt)}</span>
              <strong>{formatAuditAction(entry.action, entry.permissionRole, entry.previousRole)}</strong>
              <div>
                <strong>{entry.actorName || entry.actorEmail || "Unknown"}</strong>
                <small>{entry.actorRole || entry.actorEmail}</small>
              </div>
              <div>
                <strong>{entry.targetName || entry.targetEmail || "-"}</strong>
                <small>{entry.targetEmail}</small>
              </div>
              <div>
                <strong>{entry.libraryName || "-"}</strong>
                <small>{entry.siteName}</small>
              </div>
              <span className={entry.approvalRequestNo ? "request-ref" : "muted"}>{entry.approvalRequestNo || "-"}</span>
              <span className={`status-chip ${entry.status === "Failed" ? "failed" : "success"}`}>{entry.status}</span>
            </div>
          ))}
          {filteredAuditRecords.length === 0 && !auditError && (
            <div className="empty-row">{auditRecords.length === 0 ? "No audit entries found." : "No audit entries match the current filters."}</div>
          )}
        </div>
      )}
    </section>
  );
}

function formatAuditTime(value: string) {
  if (!value) return "-";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleString();
}

function formatAuditAction(action: AuditLogAction, role?: AccessRole, previousRole?: AccessRole) {
  if (action === "GrantAccess") return role ? `Grant ${roleLabels[role]}` : "Grant access";
  if (action === "UpdateRole") {
    const nextRole = role ? roleLabels[role] : "role";
    return previousRole ? `${roleLabels[previousRole]} to ${nextRole}` : `Update to ${nextRole}`;
  }
  if (action === "RemoveAccess") return "Remove access";
  if (action === "RefreshReport") return "Refresh review";
  return "Login";
}
