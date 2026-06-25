"use client";

import {
  Archive,
  ArrowLeft,
  Check,
  ChevronRight,
  Copy,
  File,
  Folder,
  Home as HomeIcon,
  Library,
  LockKeyhole,
  Plus,
  RefreshCw,
  Search,
  ShieldCheck,
  Trash2,
  UserRound,
  UsersRound,
} from "lucide-react";
import { FormEvent, useEffect, useRef, useState } from "react";
import { isInternalEmail, tenantDomain } from "@/lib/app-config";
import { roleLabels, type PermissionDraft } from "@/lib/features/admin";
import type { AccessRole, AuditEntry, ContentItem, PermissionEntry, SiteSummary, UserSuggestion } from "@/lib/types";
import { TableSkeleton } from "@/components/shared/Skeletons";

export type PendingPermissionAction =
  | {
      type: "grant";
      draft: PermissionDraft;
    }
  | {
      type: "update";
      permission: PermissionEntry;
      role: Exclude<AccessRole, "owner">;
    }
  | {
      type: "remove";
      permission: PermissionEntry;
    };

export function SitePicker({
  sites,
  loadingLabel,
  dataError,
  dataConsentRequired,
  onSelect,
  onRequestDataAccess,
}: {
  sites: SiteSummary[];
  loadingLabel: string;
  dataError: string;
  dataConsentRequired: boolean;
  onSelect: (site: SiteSummary) => void;
  onRequestDataAccess: () => void;
}) {
  const isLoadingSites = loadingLabel === "Loading site contents" || loadingLabel === "Restoring session";

  return (
    <section className="page-section">
      <div className="site-picker-hero">
        <div>
        <p className="section-label">Admin</p>
        <h1>Permission Management</h1>
        <p>Choose the SharePoint site where you want to review or change access.</p>
        </div>
      </div>

      <div className="site-card-grid">
        {sites.map((site, index) => (
          <button className={`site-card site-card-${index % 4}`} key={site.id} onClick={() => onSelect(site)}>
            <div className="site-card-band">
              <ShieldCheck className="site-card-star" size={18} />
            </div>
            <div className="site-card-body">
              <div className="site-card-avatar">{getInitials(site.name)}</div>
              <div className="site-card-copy">
                <h2>{site.name}</h2>
                <p>{site.hostname}</p>
              </div>
              <div className="site-card-meta">
                <span>
                  <Library size={14} />
                  Libraries
                </span>
                <span>
                  <UsersRound size={14} />
                  Permissions
                </span>
              </div>
              <span className="site-card-action">
                <Folder size={16} />
                Open site
              </span>
            </div>
          </button>
        ))}
        {sites.length === 0 && loadingLabel && <SiteCardSkeleton count={2} />}
      </div>

      {sites.length === 0 && !loadingLabel && !dataError && (
        <div className="empty-row site-empty-state">
          No SharePoint sites are configured for this app.
        </div>
      )}

      {sites.length === 0 && !loadingLabel && dataError && (
        <div className="empty-row site-empty-state action-empty-state">
          <strong>SharePoint data is not available yet.</strong>
          <span>{dataError}</span>
          {dataConsentRequired && (
            <button className="secondary-button" disabled={loadingLabel === "Requesting SharePoint access"} onClick={onRequestDataAccess}>
              <ShieldCheck size={16} />
              {loadingLabel === "Requesting SharePoint access" ? "Requesting access" : "Request SharePoint access"}
            </button>
          )}
        </div>
      )}

      {isLoadingSites && <div className="loading-note">{loadingLabel}</div>}
    </section>
  );
}

export function PermissionActionDialog({
  action,
  approvalRequestNo,
  error,
  isSubmitting,
  successLink,
  onApprovalRequestNoChange,
  onCancel,
  onConfirm,
  onCopyLink,
}: {
  action: PendingPermissionAction;
  approvalRequestNo: string;
  error: string;
  isSubmitting: boolean;
  successLink?: {
    message: string;
    url: string;
  };
  onApprovalRequestNoChange: (value: string) => void;
  onCancel: () => void;
  onConfirm: (event: FormEvent<HTMLFormElement>) => void;
  onCopyLink: (url: string) => Promise<boolean>;
}) {
  const summary = getPermissionActionSummary(action);
  const isComplete = Boolean(successLink);
  const [copyStatus, setCopyStatus] = useState<"idle" | "copying" | "copied" | "failed">("idle");
  const copyResetTimer = useRef<number | undefined>(undefined);

  useEffect(() => {
    return () => {
      if (copyResetTimer.current) {
        window.clearTimeout(copyResetTimer.current);
      }
    };
  }, []);

  async function copySharePointLink() {
    if (!successLink || copyStatus === "copying") return;

    setCopyStatus("copying");
    const copied = await onCopyLink(successLink.url);
    setCopyStatus(copied ? "copied" : "failed");

    if (copyResetTimer.current) {
      window.clearTimeout(copyResetTimer.current);
    }
    copyResetTimer.current = window.setTimeout(() => setCopyStatus("idle"), 2200);
  }

  return (
    <div className="modal-backdrop" role="presentation">
      <form className="confirm-dialog permission-confirm-dialog" onSubmit={onConfirm}>
        <div className={`confirm-icon ${action.type === "remove" ? "danger" : ""}`}>
          {action.type === "remove" ? <Trash2 size={22} /> : <ShieldCheck size={22} />}
        </div>
        <div className="confirm-copy">
          <p className="section-label">Approval Reference</p>
          <h2>{summary.title}</h2>
          <p>{summary.description}</p>
        </div>

        <dl className="confirm-summary">
          <div>
            <dt>Target</dt>
            <dd>
              <strong>{summary.targetName}</strong>
              <span>{summary.targetEmail}</span>
            </dd>
          </div>
          <div>
            <dt>Change</dt>
            <dd>
              <strong>{summary.change}</strong>
              {summary.previousRole && <span>Current role: {summary.previousRole}</span>}
            </dd>
          </div>
        </dl>

        <label className="approval-field">
          <span>Approved request no.</span>
          <input
            autoFocus
            aria-label="Approved request number"
            disabled={isComplete}
            onChange={(event) => onApprovalRequestNoChange(event.target.value)}
            placeholder="e.g. REQ-2026-0001"
            value={approvalRequestNo}
          />
        </label>

        {successLink && (
          <div className="permission-link-result">
            <Check size={18} />
            <div>
              <strong>{successLink.message}</strong>
              <span>Copy this SharePoint link and send it if the invitation email is not received.</span>
              <button
                className={`secondary-button copy-link-button ${copyStatus}`}
                disabled={copyStatus === "copying"}
                type="button"
                onClick={() => void copySharePointLink()}
              >
                {copyStatus === "copying" ? (
                  <RefreshCw className="spin-icon" size={16} />
                ) : copyStatus === "copied" ? (
                  <Check size={16} />
                ) : (
                  <Copy size={16} />
                )}
                {copyStatus === "copying" ? "Copying" : copyStatus === "copied" ? "Copied" : "Copy link"}
              </button>
              {copyStatus === "copied" && <small className="copy-link-feedback success">Link copied to clipboard.</small>}
              {copyStatus === "failed" && <small className="copy-link-feedback">Copy did not finish automatically. Use the browser prompt to copy the link.</small>}
            </div>
          </div>
        )}

        {error && <div className="auth-error confirm-error">{error}</div>}

        <div className="confirm-actions">
          <button className="secondary-button" disabled={isSubmitting} type="button" onClick={onCancel}>
            {isComplete ? "Close" : "Cancel"}
          </button>
          {!isComplete && (
            <button className={`primary-button ${action.type === "remove" ? "danger-primary" : ""}`} disabled={isSubmitting || !approvalRequestNo.trim()} type="submit">
              {action.type === "remove" ? <Trash2 size={17} /> : <Check size={17} />}
              {isSubmitting ? summary.submittingLabel : summary.confirmLabel}
            </button>
          )}
        </div>
      </form>
    </div>
  );
}

export function ContentExplorer({
  site,
  contents,
  path,
  query,
  loadingLabel,
  audit,
  onQueryChange,
  onOpen,
  onManage,
  onRoot,
  onCrumb,
  onBack,
  onSites,
  onRefresh,
}: {
  site: SiteSummary;
  contents: ContentItem[];
  path: ContentItem[];
  query: string;
  loadingLabel: string;
  audit: AuditEntry[];
  onQueryChange: (value: string) => void;
  onOpen: (item: ContentItem) => void;
  onManage: (item: ContentItem) => void;
  onRoot: () => void;
  onCrumb: (index: number) => void;
  onBack: () => void;
  onSites: () => void;
  onRefresh: () => void;
}) {
  const [selectedContent, setSelectedContent] = useState<ContentItem | null>(null);
  const isLoadingContents = loadingLabel.startsWith("Loading") || loadingLabel.startsWith("Opening") || loadingLabel.startsWith("Refreshing");
  const currentSelection = selectedContent && contents.some((item) => item.id === selectedContent.id)
    ? selectedContent
    : contents[0];

  return (
    <section className="page-section">
      <div className="page-header with-actions">
        <div>
          <p className="section-label">SharePoint Workspace</p>
          <h1>{path.at(-1)?.name ?? `${site.name} Contents`}</h1>
          <p>{path.length ? "Open a folder or manage access for this location." : "Review libraries available for admin permission management."}</p>
        </div>
        <button className="icon-button" disabled={isLoadingContents} title="Refresh" onClick={onRefresh}>
          <RefreshCw className={isLoadingContents ? "spin-icon" : ""} size={17} />
        </button>
      </div>

      <div className="explorer-bar">
        <div className="breadcrumbs">
          <button onClick={onSites}>
            <HomeIcon size={14} />
            Sites
          </button>
          <ChevronRight size={14} />
          <button onClick={onRoot}>
            Site contents
          </button>
          {path.map((node, index) => (
            <span className="breadcrumb-step" key={node.id}>
              <ChevronRight size={14} />
              <button className="breadcrumb-node" onClick={() => onCrumb(index)}>
                <strong>{node.name}</strong>
              </button>
            </span>
          ))}
        </div>
        <button className="secondary-button" onClick={path.length > 0 ? onBack : onSites}>
          <ArrowLeft size={16} />
          {path.length > 0 ? "Up one level" : "All sites"}
        </button>
      </div>

      <div className="toolbar">
        <label className="search-box">
          <Search size={17} />
          <input
            aria-label="Search contents"
            onChange={(event) => onQueryChange(event.target.value)}
            placeholder="Search current view"
            value={query}
          />
        </label>
      </div>

      <div className="explorer-layout">
        <div className="items-table">
          <div className="items-head">
            <span>Name</span>
            <span>Type</span>
            <span>Protection</span>
            <span>Updated</span>
          </div>
          {isLoadingContents ? (
            <TableSkeleton columns={4} rows={5} />
          ) : contents.map((item) => (
            <div
              className={`items-row ${currentSelection?.id === item.id ? "selected" : ""}`}
              key={item.id}
              onClick={() => setSelectedContent(item)}
              onDoubleClick={() => onOpen(item)}
            >
              <button
                className="item-name-button"
                title={`Open ${item.name}`}
                type="button"
                onClick={(event) => {
                  event.stopPropagation();
                  onOpen(item);
                }}
              >
                <ItemIcon item={item} />
                <span>
                  <strong>{item.name}</strong>
                  <small>{formatItemMeta(item)}</small>
                </span>
              </button>
              <span className="muted" data-label="Type">{item.type}</span>
              <span className={`policy-badge ${item.protected ? "protected" : ""}`} data-label="Protection">
                {item.rightsPolicy}
              </span>
              <span className="muted" data-label="Updated">{formatModifiedDate(item.modified)}</span>
            </div>
          ))}
          {!isLoadingContents && contents.length === 0 && <div className="empty-row">No items found.</div>}
        </div>

        <aside className="details-pane">
          {isLoadingContents ? (
            <DetailsSkeleton />
          ) : currentSelection ? (
            <>
              <div className="details-icon">
                <ItemIcon item={currentSelection} />
              </div>
              <div className="details-copy">
                <p className="section-label">Selected Item</p>
                <h2>{currentSelection.name}</h2>
                <span>{currentSelection.type}</span>
              </div>
              <div className="details-meta">
                <div>
                  <span>Protection</span>
                  <strong>{currentSelection.rightsPolicy}</strong>
                </div>
                <div>
                  <span>Contains</span>
                  <strong>{formatItemMeta(currentSelection)}</strong>
                </div>
              </div>
              <div className="details-actions">
                {currentSelection.type !== "file" && (
                  <button className="secondary-button" onClick={() => onOpen(currentSelection)}>
                    <Folder size={16} />
                    Open
                  </button>
                )}
                <button className="primary-button" onClick={() => onManage(currentSelection)}>
                  <ShieldCheck size={16} />
                  Manage access
                </button>
              </div>
            </>
          ) : (
            <div className="empty-selection">
              <ShieldCheck size={22} />
              <strong>Select an item</strong>
              <span>Choose a library, folder, or file to manage access.</span>
            </div>
          )}

          <div className="recent-compact">
            <div className="recent-heading">
              <p className="section-label">Recent Changes</p>
              {audit.length > 2 && <span>{audit.length - 2} more</span>}
            </div>
            <div className="audit-list">
              {audit.slice(0, 2).map((entry) => (
                <article className={`audit-item ${entry.status === "Failed" ? "failed" : ""}`} key={entry.id}>
                  <span className="audit-dot" />
                  <div>
                    <strong>{entry.action}</strong>
                    <p>{entry.target}</p>
                    <small>
                      {entry.actor} - {entry.time}
                    </small>
                  </div>
                </article>
              ))}
            </div>
          </div>
        </aside>
      </div>
    </section>
  );
}

export function AccessPanel({
  item,
  permissions,
  query,
  newEmail,
  newRole,
  userSuggestions,
  suggestionsLoading,
  suggestionError,
  canManagePermissions,
  loadingLabel,
  backLabel,
  onBack,
  onRefresh,
  onQueryChange,
  onEmailChange,
  onSelectUserSuggestion,
  onRoleChange,
  onGrant,
  onUpdateRole,
  onRemove,
}: {
  item: ContentItem;
  permissions: PermissionEntry[];
  query: string;
  newEmail: string;
  newRole: AccessRole;
  userSuggestions: UserSuggestion[];
  suggestionsLoading: boolean;
  suggestionError: string;
  canManagePermissions: boolean;
  loadingLabel: string;
  backLabel: string;
  onBack: () => void;
  onRefresh: () => void;
  onQueryChange: (value: string) => void;
  onEmailChange: (value: string) => void;
  onSelectUserSuggestion: (user: UserSuggestion) => void;
  onRoleChange: (value: AccessRole) => void;
  onGrant: (event: FormEvent<HTMLFormElement>) => void;
  onUpdateRole: (permissionId: string, role: Exclude<AccessRole, "owner">) => void;
  onRemove: (permissionId: string) => void;
}) {
  const isLoadingPermissions =
    loadingLabel === "Loading permissions" ||
    loadingLabel === "Removing permission" ||
    loadingLabel === "Updating role" ||
    loadingLabel === "Granting permission";
  const externalGrantTarget = newEmail.includes("@") && !isInternalEmail(newEmail);
  const lockedCount = permissions.filter(
    (permission) => permission.canEditRole === false || permission.canDelete === false || permission.role === "owner",
  ).length;
  const editableCount = permissions.length - lockedCount;
  const directPermissions = permissions.filter(
    (permission) => permission.canEditRole !== false || permission.canDelete !== false,
  );
  const managedPermissions = permissions.filter(
    (permission) => permission.canEditRole === false && permission.canDelete === false,
  );

  return (
    <section className="page-section">
      <div className="page-header with-actions">
        <div>
          <p className="section-label">Manage Access</p>
          <h1>{item.name}</h1>
          <p>{item.protected ? "Protected library context. Downloaded Office files remain governed by Rights Management." : "Review who can open the selected item."}</p>
        </div>
        <div className="header-actions">
          <button className="secondary-button" onClick={onBack}>
            <ArrowLeft size={16} />
            Back to {backLabel}
          </button>
          <button className="icon-button" disabled={isLoadingPermissions} title="Refresh permissions" onClick={onRefresh}>
            <RefreshCw className={isLoadingPermissions ? "spin-icon" : ""} size={17} />
          </button>
        </div>
      </div>

      <div className="access-summary">
        <span className="library-icon">
          <ItemIcon item={item} />
        </span>
        <div>
          <strong>{item.type === "library" ? "Library root" : item.type}</strong>
          <span>{item.rightsPolicy}</span>
        </div>
      </div>

      <div className="access-metrics">
        <div>
          <strong>{directPermissions.length}</strong>
          <span>Can edit here</span>
        </div>
        <div>
          <strong>{managedPermissions.length}</strong>
          <span>Managed elsewhere</span>
        </div>
        <div>
          <strong>{permissions.length}</strong>
          <span>People and groups</span>
        </div>
      </div>

      {permissions.length > 0 && lockedCount > 0 && (
        <div className="info-message">
          <ShieldCheck size={18} />
          <span>
            {editableCount === 0
              ? "No editable permissions are shown here. Add Viewer/Editor access below, or manage parent access in SharePoint."
              : `${lockedCount} permission${lockedCount > 1 ? "s" : ""} ${lockedCount === 1 ? "comes" : "come"} from SharePoint or a parent folder and can only be changed there.`}
          </span>
        </div>
      )}

      {!canManagePermissions && (
        <div className="info-message readonly-message">
          <ShieldCheck size={18} />
          <span>Your app role is read-only here. You can review permissions, but only Admin can grant, change, or remove access.</span>
        </div>
      )}

      {canManagePermissions && (
        <div className="info-message approval-message">
          <ShieldCheck size={18} />
          <span>Permission changes require confirmation with an approved request number.</span>
        </div>
      )}

      <div className="toolbar">
        <label className="search-box">
          <Search size={17} />
          <input
            aria-label="Search permissions"
            onChange={(event) => onQueryChange(event.target.value)}
            placeholder="Search people, groups, roles"
            value={query}
          />
        </label>
      </div>

      {canManagePermissions && (
        <form className="grant-panel" onSubmit={onGrant}>
          <div className="grant-title">
            <strong>Grant access</strong>
            <span>Use this for people or groups who need access to this item.</span>
          </div>
          <label className="people-picker-field">
            <span>User email</span>
            <input
              onChange={(event) => onEmailChange(event.target.value)}
              placeholder={`name@${tenantDomain}`}
              type="email"
              value={newEmail}
            />
            {(suggestionsLoading || suggestionError || userSuggestions.length > 0) && (
              <div className="people-suggestions" role="listbox" aria-label="People suggestions">
                {suggestionsLoading && <div className="people-suggestion-status">Searching people</div>}
                {!suggestionsLoading && suggestionError && (
                  <div className="people-suggestion-status error">{suggestionError}</div>
                )}
                {!suggestionsLoading &&
                  !suggestionError &&
                  userSuggestions.map((user) => (
                    <button
                      className="people-suggestion"
                      key={user.id}
                      type="button"
                      role="option"
                      aria-selected={newEmail === user.email}
                      onClick={() => onSelectUserSuggestion(user)}
                    >
                      <span className="avatar small">
                        <UserRound size={15} />
                      </span>
                      <span>
                        <strong>{user.displayName}</strong>
                        <small>{user.email}</small>
                        {user.jobTitle && <small>{user.jobTitle}</small>}
                      </span>
                    </button>
                  ))}
                {!suggestionsLoading && !suggestionError && userSuggestions.length === 0 && (
                  <div className="people-suggestion-status">No matching people found</div>
                )}
              </div>
            )}
            {externalGrantTarget && (
              <small className="field-hint">
                External recipient. SharePoint will send a sharing invitation email when access is granted.
              </small>
            )}
          </label>
          <label>
            <span>Role</span>
            <select onChange={(event) => onRoleChange(event.target.value as AccessRole)} value={newRole}>
              <option value="viewer">Viewer</option>
              <option value="editor">Editor</option>
            </select>
          </label>
          <button className="primary-button" disabled={loadingLabel === "Granting permission"} type="submit">
            <Plus size={18} />
            {loadingLabel === "Granting permission" ? "Granting" : "Review grant"}
          </button>
        </form>
      )}

      <div className="permission-section-title">
        <div>
          <p className="section-label">Editable Here</p>
          <h2>People you can change</h2>
        </div>
        <span>{directPermissions.length}</span>
      </div>

      <PermissionTable
        permissions={directPermissions}
        emptyText="No editable permissions found."
        isLoading={isLoadingPermissions}
        canManagePermissions={canManagePermissions}
        onUpdateRole={onUpdateRole}
        onRemove={onRemove}
      />

      {managedPermissions.length > 0 && (
        <>
          <div className="permission-section-title muted-title">
            <div>
              <p className="section-label">Managed Elsewhere</p>
              <h2>People shown for context</h2>
            </div>
            <span>{managedPermissions.length}</span>
          </div>
          <PermissionTable
            permissions={managedPermissions}
            emptyText=""
            isLoading={isLoadingPermissions}
            canManagePermissions={canManagePermissions}
            onUpdateRole={onUpdateRole}
            onRemove={onRemove}
          />
        </>
      )}
    </section>
  );
}

function PermissionTable({
  permissions,
  emptyText,
  isLoading,
  canManagePermissions,
  onUpdateRole,
  onRemove,
}: {
  permissions: PermissionEntry[];
  emptyText: string;
  isLoading?: boolean;
  canManagePermissions: boolean;
  onUpdateRole: (permissionId: string, role: Exclude<AccessRole, "owner">) => void;
  onRemove: (permissionId: string) => void;
}) {
  return (
    <div className="permission-table" role="table" aria-label="Permissions">
      <div className="table-head" role="row">
        <span>Principal</span>
        <span>Role</span>
        <span>Access path</span>
        <span>Action</span>
      </div>
      {isLoading ? (
        <TableSkeleton columns={4} rows={4} />
      ) : permissions.map((permission) => (
        <div className="table-row" role="row" key={permission.id}>
          <div className="principal-cell">
            <span className="avatar">
              {permission.type === "group" ? <UsersRound size={17} /> : <UserRound size={17} />}
            </span>
            <span>
              <strong>{permission.displayName}</strong>
              <small>{permission.email}</small>
            </span>
          </div>
          <span className={`role-chip role-display ${permission.role}`} data-label="Role">{roleLabels[permission.role]}</span>
          <span className="muted" data-label="Access path">{formatPermissionSource(permission)}</span>
          <div className="row-actions" data-label="Action">
            {!canManagePermissions ? (
              <span className="locked-badge">Read-only</span>
            ) : permission.canDelete === false ? (
              <span className="locked-badge">{permission.source === "inherited" ? "Parent access" : "SharePoint"}</span>
            ) : (
              <>
                <button
                  className="permission-action-button viewer"
                  disabled={permission.role === "viewer" || permission.canEditRole === false}
                  type="button"
                  onClick={() => onUpdateRole(permission.id, "viewer")}
                >
                  Viewer
                </button>
                <button
                  className="permission-action-button editor"
                  disabled={permission.role === "editor" || permission.canEditRole === false}
                  type="button"
                  onClick={() => onUpdateRole(permission.id, "editor")}
                >
                  Editor
                </button>
                <button
                  className="permission-action-button danger"
                  title="Remove this access"
                  type="button"
                  onClick={() => onRemove(permission.id)}
                >
                  <Trash2 size={14} />
                  Remove
                </button>
              </>
            )}
          </div>
        </div>
      ))}
      {!isLoading && permissions.length === 0 && emptyText && <div className="empty-row">{emptyText}</div>}
    </div>
  );
}

function SiteCardSkeleton({ count }: { count: number }) {
  return Array.from({ length: count }).map((_, index) => (
    <div className="site-card skeleton-card" key={`site-skeleton-${index}`}>
      <div className="site-card-band skeleton-block" />
      <div className="site-card-body">
        <div className="site-card-avatar skeleton-block" />
        <div className="skeleton-line wide" />
        <div className="skeleton-line medium" />
        <div className="skeleton-pill-row">
          <span className="skeleton-pill" />
          <span className="skeleton-pill" />
        </div>
      </div>
    </div>
  ));
}

function formatModifiedDate(value: string | undefined) {
  return value && value !== "Live Graph" ? value : "Current";
}

function formatPermissionSource(permission: PermissionEntry) {
  if (permission.source === "inherited") return "From parent folder";
  if (permission.source === "group") return "Via SharePoint group";
  if (permission.source === "link") return "Sharing link";
  return "Added here";
}

function DetailsSkeleton() {
  return (
    <div className="skeleton-stack" aria-label="Loading details">
      <span className="details-icon skeleton-block" />
      <span className="skeleton-line short" />
      <span className="skeleton-line wide" />
      <span className="skeleton-line medium" />
      <div className="details-meta">
        <div>
          <span className="skeleton-line short" />
          <strong className="skeleton-line medium" />
        </div>
        <div>
          <span className="skeleton-line short" />
          <strong className="skeleton-line medium" />
        </div>
      </div>
    </div>
  );
}

function ItemIcon({ item }: { item: ContentItem }) {
  if (item.protected && item.type === "library") {
    return (
      <span className="item-icon-stack protected-library-icon" aria-hidden="true">
        <Archive size={20} />
        <LockKeyhole className="item-icon-lock" size={11} strokeWidth={2.7} />
      </span>
    );
  }
  if (item.type === "library") return <Archive size={20} />;
  if (item.type === "folder") return <Folder size={20} />;
  return <File size={20} />;
}

function formatItemMeta(item: ContentItem) {
  if (item.type === "file") return item.size ? formatBytes(item.size) : "File";
  return `${item.childCount ?? 0} items`;
}

function formatBytes(bytes: number) {
  if (!bytes) return "0 B";
  const units = ["B", "KB", "MB", "GB"];
  const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  return `${Math.round((bytes / 1024 ** index) * 10) / 10} ${units[index]}`;
}

function getPermissionActionSummary(action: PendingPermissionAction) {
  if (action.type === "grant") {
    return {
      title: "Confirm grant access",
      description: "Enter the approved request number before granting access.",
      targetName: action.draft.displayName,
      targetEmail: action.draft.email,
      change: `Grant ${roleLabels[action.draft.role]}`,
      previousRole: undefined,
      confirmLabel: "Grant access",
      submittingLabel: "Granting",
    };
  }

  if (action.type === "update") {
    return {
      title: "Confirm role update",
      description: "Enter the approved request number before changing this permission.",
      targetName: action.permission.displayName,
      targetEmail: action.permission.email,
      change: `Change to ${roleLabels[action.role]}`,
      previousRole: roleLabels[action.permission.role],
      confirmLabel: "Update role",
      submittingLabel: "Updating",
    };
  }

  return {
    title: "Confirm remove access",
    description: "Enter the approved request number before removing this access.",
    targetName: action.permission.displayName,
    targetEmail: action.permission.email,
    change: "Remove access",
    previousRole: roleLabels[action.permission.role],
    confirmLabel: "Remove access",
    submittingLabel: "Removing",
  };
}

function getInitials(label: string) {
  const words = label.split(/\s+/).filter(Boolean);
  if (words.length === 1) return words[0].slice(0, 2).toUpperCase();
  return words
    .slice(0, 2)
    .map((word) => word[0])
    .join("")
    .toUpperCase();
}
