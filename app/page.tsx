"use client";

import type { AccountInfo } from "@azure/msal-browser";
import {
  ArrowLeft,
  ChevronRight,
  File,
  FileLock2,
  Folder,
  BarChart3,
  Home as HomeIcon,
  Library,
  LogIn,
  LogOut,
  Plus,
  RefreshCw,
  Search,
  ShieldCheck,
  Trash2,
  UserRound,
  UsersRound,
} from "lucide-react";
import { FormEvent, useEffect, useMemo, useRef, useState } from "react";
import { acquireGraphToken, isAuthConfigured, signInMicrosoft365 } from "@/lib/auth";
import { isInternalEmail, tenantDomain } from "@/lib/app-config";
import { filterContentItemsForRoles, getAccountRoles, getCapabilities, getPrimaryRole, getRoleLabel } from "@/lib/app-roles";
import { normalizeSharePointPrincipals } from "@/lib/permission-normalization";
import {
  GraphSharePointPermissionClient,
  type PermissionDraft,
  type SharePointPermissionClient,
} from "@/lib/graph";
import type {
  AccessRole,
  AuditEntry,
  ContentItem,
  PermissionEntry,
  ReportSummary,
  SiteSummary,
  UserSuggestion,
} from "@/lib/types";

const roleLabels: Record<AccessRole, string> = {
  viewer: "Viewer",
  editor: "Editor",
  owner: "Owner",
};

type AppHistoryView = {
  selectedSite: SiteSummary | null;
  path: ContentItem[];
  selectedItem: ContentItem | null;
};

type AppHistoryState = {
  spAccessView?: AppHistoryView;
};

type PortalView = "workspace" | "reports";

export default function Home() {
  const [signedIn, setSignedIn] = useState(false);
  const [account, setAccount] = useState<AccountInfo | null>(null);
  const [accountLabel, setAccountLabel] = useState("Admin");
  const [roleLabel, setRoleLabel] = useState("No app role");
  const [sites, setSites] = useState<SiteSummary[]>([]);
  const [portalView, setPortalView] = useState<PortalView>("workspace");
  const [selectedSite, setSelectedSite] = useState<SiteSummary | null>(null);
  const [contents, setContents] = useState<ContentItem[]>([]);
  const [path, setPath] = useState<ContentItem[]>([]);
  const [selectedItem, setSelectedItem] = useState<ContentItem | null>(null);
  const [permissions, setPermissions] = useState<PermissionEntry[]>([]);
  const [audit, setAudit] = useState<AuditEntry[]>([]);
  const [query, setQuery] = useState("");
  const [newEmail, setNewEmail] = useState("");
  const [newRole, setNewRole] = useState<AccessRole>("viewer");
  const [userSuggestions, setUserSuggestions] = useState<UserSuggestion[]>([]);
  const [suggestionError, setSuggestionError] = useState("");
  const [suggestionsLoading, setSuggestionsLoading] = useState(false);
  const [reportSummary, setReportSummary] = useState<ReportSummary | null>(null);
  const [reportError, setReportError] = useState("");
  const [authError, setAuthError] = useState("");
  const [dataError, setDataError] = useState("");
  const [loadingLabel, setLoadingLabel] = useState("");
  const restoringHistoryRef = useRef(false);

  const graphClient = useMemo<SharePointPermissionClient>(() => {
    return new GraphSharePointPermissionClient(() => acquireGraphToken(account));
  }, [account]);

  const appRoles = useMemo(() => getAccountRoles(account), [account]);
  const capabilities = useMemo(() => getCapabilities(appRoles), [appRoles]);

  useEffect(() => {
    if (!signedIn) return;

    replaceAppHistory(getHistoryView());

    function handlePopState(event: PopStateEvent) {
      const view = (event.state as AppHistoryState | null)?.spAccessView;
      if (!view) return;

      void restoreHistoryView(view);
    }

    window.addEventListener("popstate", handlePopState);
    return () => window.removeEventListener("popstate", handlePopState);
    // The handler intentionally restores through the latest Graph client after sign-in.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [signedIn, graphClient]);

  useEffect(() => {
    if (!signedIn || !selectedItem) {
      setUserSuggestions([]);
      setSuggestionsLoading(false);
      setSuggestionError("");
      return;
    }

    const search = newEmail.trim();
    if (search.length < 2) {
      setUserSuggestions([]);
      setSuggestionsLoading(false);
      setSuggestionError("");
      return;
    }

    let cancelled = false;
    setSuggestionsLoading(true);
    setSuggestionError("");

    const timeout = window.setTimeout(() => {
      void graphClient
        .searchUsers(search)
        .then((users) => {
          if (!cancelled) setUserSuggestions(users);
        })
        .catch((error) => {
          if (!cancelled) {
            setUserSuggestions([]);
            setSuggestionError(error instanceof Error ? error.message : "Unable to search people.");
          }
        })
        .finally(() => {
          if (!cancelled) setSuggestionsLoading(false);
        });
    }, 300);

    return () => {
      cancelled = true;
      window.clearTimeout(timeout);
    };
  }, [graphClient, newEmail, selectedItem, signedIn]);

  const visibleContents = useMemo(() => {
    const search = query.trim().toLowerCase();
    if (!search || selectedItem) return contents;
    return contents.filter((item) => item.name.toLowerCase().includes(search));
  }, [contents, query, selectedItem]);

  const visiblePermissions = useMemo(() => {
    const search = query.trim().toLowerCase();
    return permissions.filter((permission) => {
      if (!search) return true;
      return (
        permission.displayName.toLowerCase().includes(search) ||
        permission.email.toLowerCase().includes(search) ||
        permission.role.toLowerCase().includes(search)
      );
    });
  }, [permissions, query]);

  async function connectMicrosoft365() {
    setAuthError("");
    setDataError("");
    setLoadingLabel("Connecting Microsoft 365");

    try {
      if (!isAuthConfigured) {
        setAuthError("Microsoft Entra configuration is missing. Set NEXT_PUBLIC_MSAL_CLIENT_ID and NEXT_PUBLIC_MSAL_TENANT_ID before signing in.");
        return;
      }

      const response = await signInMicrosoft365();
      const nextRoles = getAccountRoles(response.account);
      if (nextRoles.length === 0) {
        setAuthError("Your account is signed in but has no app role assigned. Ask an administrator to assign Admin, InternalUser, GuestUser, or ExecutiveUser.");
        return;
      }

      const nextClient = new GraphSharePointPermissionClient(() => acquireGraphToken(response.account));
      const nextSites = await nextClient.listSites();

      setSignedIn(true);
      setAccount(response.account);
      setAccountLabel(response.account?.username ?? "Microsoft 365 Admin");
      setRoleLabel(getRoleLabel(getPrimaryRole(nextRoles)));
      setSites(nextSites);
      replaceAppHistory({ selectedSite: null, path: [], selectedItem: null });
    } catch (error) {
      setAuthError(error instanceof Error ? error.message : "Unable to connect Microsoft 365.");
    } finally {
      setLoadingLabel("");
    }
  }

  function signOut() {
    setSignedIn(false);
    setAccount(null);
    setAccountLabel("Admin");
    setRoleLabel("No app role");
    setPortalView("workspace");
    setSelectedSite(null);
    setContents([]);
    setPath([]);
    setSelectedItem(null);
    setPermissions([]);
    setQuery("");
    setNewEmail("");
    setAuthError("");
    setDataError("");
    setReportSummary(null);
    setReportError("");
    replaceAppHistory({ selectedSite: null, path: [], selectedItem: null });
  }

  function returnToSites() {
    setPortalView("workspace");
    setSelectedSite(null);
    setContents([]);
    setPath([]);
    setSelectedItem(null);
    setPermissions([]);
    setQuery("");
    setDataError("");
    pushAppHistory({ selectedSite: null, path: [], selectedItem: null });
  }

  async function openReports() {
    if (!capabilities.canViewReports) return;

    setPortalView("reports");
    setSelectedSite(null);
    setContents([]);
    setPath([]);
    setSelectedItem(null);
    setPermissions([]);
    setQuery("");
    setDataError("");
    setReportError("");
    setLoadingLabel("Loading reports");

    try {
      setReportSummary(await graphClient.getReportSummary());
    } catch (error) {
      setReportSummary(null);
      setReportError(error instanceof Error ? error.message : "Unable to load reports.");
    } finally {
      setLoadingLabel("");
    }
  }

  async function chooseSite(site: SiteSummary, options: { updateHistory?: boolean } = {}) {
    const updateHistory = options.updateHistory ?? true;
    setSelectedSite(site);
    setSelectedItem(null);
    setPath([]);
    setQuery("");
    setPermissions([]);
    setDataError("");
    setLoadingLabel("Loading site contents");

    try {
      const nextContents = await graphClient.listContentItems(site.id);
      setContents(filterContentItemsForRoles(nextContents, appRoles));
      if (updateHistory) {
        pushAppHistory({ selectedSite: site, path: [], selectedItem: null });
      }
    } catch (error) {
      setDataError(error instanceof Error ? error.message : "Unable to load site contents.");
      setContents([]);
    } finally {
      setLoadingLabel("");
    }
  }

  async function openItem(item: ContentItem) {
    if (item.type === "file") {
      await manageAccess(item);
      return;
    }

    setSelectedItem(null);
    setPermissions([]);
    setQuery("");
    setDataError("");
    setLoadingLabel(`Opening ${item.name}`);

    try {
      const children = await graphClient.listChildren(item);
      setContents(children);
      const nextPath = [...path, item];
      setPath(nextPath);
      pushAppHistory({ selectedSite, path: nextPath, selectedItem: null });
    } catch (error) {
      setDataError(error instanceof Error ? error.message : "Unable to open this item.");
    } finally {
      setLoadingLabel("");
    }
  }

  async function goToSiteRoot() {
    if (!selectedSite) return;
    await chooseSite(selectedSite);
  }

  async function goToPath(index: number) {
    if (!selectedSite) return;

    if (index < 0) {
      await goToSiteRoot();
      return;
    }

    const target = path[index];
    if (!target) return;

    setPath(path.slice(0, index + 1));
    setSelectedItem(null);
    setPermissions([]);
    setQuery("");
    setDataError("");
    setLoadingLabel(`Opening ${target.name}`);

    try {
      setContents(await graphClient.listChildren(target));
      pushAppHistory({ selectedSite, path: path.slice(0, index + 1), selectedItem: null });
    } catch (error) {
      setDataError(error instanceof Error ? error.message : "Unable to open this location.");
    } finally {
      setLoadingLabel("");
    }
  }

  async function goBackOneLevel() {
    if (!selectedSite) return;
    const nextPath = path.slice(0, -1);
    const parent = nextPath.at(-1);
    setPath(nextPath);
    setSelectedItem(null);
    setPermissions([]);

    if (!parent) {
      await chooseSite(selectedSite);
      return;
    }

    setLoadingLabel(`Opening ${parent.name}`);
    try {
      setContents(await graphClient.listChildren(parent));
      pushAppHistory({ selectedSite, path: nextPath, selectedItem: null });
    } catch (error) {
      setDataError(error instanceof Error ? error.message : "Unable to go back.");
    } finally {
      setLoadingLabel("");
    }
  }

  async function manageAccess(item: ContentItem, options: { updateHistory?: boolean } = {}) {
    const updateHistory = options.updateHistory ?? true;
    setSelectedItem(item);
    setQuery("");
    setDataError("");
    setLoadingLabel("Loading permissions");

    try {
      const nextPermissions = await graphClient.listPermissions(item.id);
      setPermissions(normalizeSharePointPrincipals(nextPermissions, selectedSite?.name ?? "Site"));
      if (selectedSite && updateHistory) {
        pushAppHistory({ selectedSite, path, selectedItem: item });
      }
    } catch (error) {
      setDataError(error instanceof Error ? error.message : "Unable to load permissions.");
      setPermissions([]);
    } finally {
      setLoadingLabel("");
    }
  }

  async function refreshCurrentView() {
    if (selectedItem) {
      await manageAccess(selectedItem, { updateHistory: false });
      return;
    }

    if (path.length === 0) {
      if (selectedSite) await chooseSite(selectedSite, { updateHistory: false });
      return;
    }

    const current = path.at(-1);
    if (current) {
      setLoadingLabel(`Refreshing ${current.name}`);
      try {
        setContents(await graphClient.listChildren(current));
      } catch (error) {
        setDataError(error instanceof Error ? error.message : "Unable to refresh.");
      } finally {
        setLoadingLabel("");
      }
    }
  }

  function leaveAccessPanel() {
    setSelectedItem(null);
    setPermissions([]);
    setQuery("");
    if (selectedSite) {
      pushAppHistory({ selectedSite, path, selectedItem: null });
    }
  }

  async function addPermission(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!selectedItem || !newEmail.trim()) return;

    const draft: PermissionDraft = {
      displayName: newEmail.trim(),
      email: newEmail.trim(),
      role: newRole,
    };

    setDataError("");
    setLoadingLabel("Granting permission");

    try {
      const created = await graphClient.grantPermission(selectedItem, draft);
      setPermissions((current) => [...created, ...current]);

      addAudit(`Granted ${roleLabels[newRole].toLowerCase()}`, draft.email);
      setNewEmail("");
      setUserSuggestions([]);
    } catch (error) {
      setDataError(error instanceof Error ? error.message : "Unable to grant permission.");
    } finally {
      setLoadingLabel("");
    }
  }

  async function updateRole(permissionId: string, role: AccessRole) {
    const changed = permissions.find((permission) => permission.id === permissionId);
    if (!changed || changed.role === role) return;

    setDataError("");
    setLoadingLabel("Updating role");

    try {
      const updated = await graphClient.updatePermissionRole(changed, role);
      setPermissions((current) =>
        current.map((permission) => (permission.id === permissionId ? updated : permission)),
      );
      addAudit(`Changed role to ${roleLabels[role]}`, changed.email);
    } catch (error) {
      setDataError(error instanceof Error ? error.message : "Unable to update role.");
    } finally {
      setLoadingLabel("");
    }
  }

  async function removePermission(permissionId: string) {
    const removed = permissions.find((permission) => permission.id === permissionId);
    if (!removed) return;

    setDataError("");
    setLoadingLabel("Removing permission");

    try {
      await graphClient.removePermission(removed);
      setPermissions((current) => current.filter((permission) => permission.id !== permissionId));
      addAudit("Removed permission", removed.email);
    } catch (error) {
      setDataError(error instanceof Error ? error.message : "Unable to remove permission.");
    } finally {
      setLoadingLabel("");
    }
  }

  function addAudit(action: string, target: string) {
    setAudit((current) => [
      {
        id: `audit-${Date.now()}`,
        actor: accountLabel,
        action,
        target,
        time: "Just now",
      },
      ...current,
    ]);
  }

  function selectUserSuggestion(user: UserSuggestion) {
    setNewEmail(user.email);
    setUserSuggestions([]);
    setSuggestionError("");
  }

  function getHistoryView(): AppHistoryView {
    return { selectedSite, path, selectedItem };
  }

  function pushAppHistory(view: AppHistoryView) {
    if (restoringHistoryRef.current) return;
    window.history.pushState({ spAccessView: view } satisfies AppHistoryState, "", window.location.href);
  }

  function replaceAppHistory(view: AppHistoryView) {
    if (typeof window === "undefined") return;
    window.history.replaceState({ spAccessView: view } satisfies AppHistoryState, "", window.location.href);
  }

  async function restoreHistoryView(view: AppHistoryView) {
    restoringHistoryRef.current = true;
    setDataError("");
    setQuery("");
    setSelectedSite(view.selectedSite);
    setPath(view.path);
    setSelectedItem(view.selectedItem);
    setPermissions([]);

    if (!view.selectedSite) {
      setContents([]);
      setLoadingLabel("");
      restoringHistoryRef.current = false;
      return;
    }

    setLoadingLabel(view.selectedItem ? "Loading permissions" : "Loading site contents");

    try {
      const nextContents = await loadContentsForHistoryView(view.selectedSite, view.path);
      setContents(nextContents);

      if (view.selectedItem) {
        const nextPermissions = await graphClient.listPermissions(view.selectedItem.id);
        setPermissions(normalizeSharePointPrincipals(nextPermissions, view.selectedSite.name));
      }
    } catch (error) {
      setDataError(error instanceof Error ? error.message : "Unable to restore the previous view.");
      setContents([]);
      setPermissions([]);
    } finally {
      setLoadingLabel("");
      restoringHistoryRef.current = false;
    }
  }

  async function loadContentsForHistoryView(site: SiteSummary, nextPath: ContentItem[]) {
    const currentFolder = nextPath.at(-1);
    if (currentFolder) {
      return graphClient.listChildren(currentFolder);
    }
    const rootContents = await graphClient.listContentItems(site.id);
    return filterContentItemsForRoles(rootContents, appRoles);
  }

  if (!signedIn) {
    return (
      <main className="auth-shell">
        <section className="auth-card" aria-labelledby="auth-title">
          <div className="brand auth-brand">
            <span className="brand-mark">
              <ShieldCheck size={19} strokeWidth={2.4} />
            </span>
            <div>
              <strong>SP Access</strong>
              <span>Permission Console</span>
            </div>
          </div>

          <div className="auth-copy">
            <p className="section-label">Microsoft 365 Access</p>
            <h1 id="auth-title">Sign in to manage SharePoint permissions</h1>
            <p>Only authenticated Microsoft 365 users can view sites and manage access.</p>
          </div>

          {authError && <div className="auth-error">{authError}</div>}

          <button className="login-button auth-login-button" onClick={connectMicrosoft365}>
            <LogIn size={18} />
            {loadingLabel === "Connecting Microsoft 365" ? "Connecting" : "Sign in with Microsoft 365"}
          </button>

          <div className="auth-meta">
            <span className={`status-dot ${isAuthConfigured ? "live" : ""}`} />
            <span>{isAuthConfigured ? "Entra app configured" : "Entra app not configured"}</span>
          </div>
        </section>
      </main>
    );
  }

  return (
    <main className="portal-shell">
      <aside className="portal-sidebar" aria-label="Workspace navigation">
        <div className="brand">
          <span className="brand-mark solid">
            <ShieldCheck size={19} strokeWidth={2.4} />
          </span>
          <div>
            <strong>SP Access</strong>
            <span>Permission Console</span>
          </div>
        </div>

        <nav className="sidebar-nav">
          <button
            className={`sidebar-nav-item ${portalView === "workspace" && !selectedSite ? "active" : ""}`}
            onClick={returnToSites}
          >
            <HomeIcon size={18} />
            Sites
          </button>
          <button className={`sidebar-nav-item ${portalView === "workspace" && selectedSite ? "active" : ""}`} disabled={!selectedSite}>
            <Library size={18} />
            Site contents
          </button>
          {capabilities.canViewReports && (
            <button className={`sidebar-nav-item ${portalView === "reports" ? "active" : ""}`} onClick={openReports}>
              <BarChart3 size={18} />
              Reports
            </button>
          )}
        </nav>

        <section className="workspace-card">
          <p className="section-label">Current Workspace</p>
          <strong>{selectedSite?.name ?? "No site selected"}</strong>
          <span>{selectedSite ? selectedSite.hostname : "Choose a site to begin."}</span>
        </section>
      </aside>

      <section className="portal-main">
        <header className="portal-topbar">
          <div className="top-context">
            <ShieldCheck size={17} />
            <span>
              {selectedItem
                ? `Managing access: ${selectedItem.name}`
                : portalView === "reports"
                  ? "Reports"
                  : selectedSite
                  ? `Browsing: ${selectedSite.name}`
                  : "Select a SharePoint site"}
            </span>
          </div>

          <div className="user-cluster">
            <span className="status-dot live" />
            <span>Live Graph</span>
            <span className={`role-pill ${capabilities.isReadOnly ? "readonly" : ""}`}>{roleLabel}</span>
            <button className="avatar-button" title={accountLabel}>
              {accountLabel.charAt(0).toUpperCase()}
            </button>
            <button className="icon-button" title="Sign out" onClick={signOut}>
              <LogOut size={17} />
            </button>
          </div>
        </header>

        <div className="portal-content">
          {dataError && <div className="auth-error">{dataError}</div>}

          {portalView === "reports" ? (
            <ReportsPanel
              loadingLabel={loadingLabel}
              report={reportSummary}
              reportError={reportError}
              onRefresh={openReports}
            />
          ) : !selectedSite ? (
            <SitePicker sites={sites} onSelect={chooseSite} loadingLabel={loadingLabel} />
          ) : selectedItem ? (
            <AccessPanel
              item={selectedItem}
              permissions={visiblePermissions}
              query={query}
              newEmail={newEmail}
              newRole={newRole}
              userSuggestions={userSuggestions}
              suggestionsLoading={suggestionsLoading}
              suggestionError={suggestionError}
              canManagePermissions={capabilities.canManagePermissions}
              loadingLabel={loadingLabel}
              backLabel={path.at(-1)?.name ?? `${selectedSite.name} contents`}
              onBack={leaveAccessPanel}
              onRefresh={refreshCurrentView}
              onQueryChange={setQuery}
              onEmailChange={setNewEmail}
              onSelectUserSuggestion={selectUserSuggestion}
              onRoleChange={setNewRole}
              onGrant={addPermission}
              onUpdateRole={updateRole}
              onRemove={removePermission}
            />
          ) : (
            <ContentExplorer
              site={selectedSite}
              contents={visibleContents}
              path={path}
              query={query}
              loadingLabel={loadingLabel}
              audit={audit}
              onQueryChange={setQuery}
              onOpen={openItem}
              onManage={manageAccess}
              onRoot={goToSiteRoot}
              onCrumb={goToPath}
              onBack={goBackOneLevel}
              onSites={returnToSites}
              onRefresh={refreshCurrentView}
            />
          )}
        </div>
      </section>
    </main>
  );
}

function SitePicker({
  sites,
  loadingLabel,
  onSelect,
}: {
  sites: SiteSummary[];
  loadingLabel: string;
  onSelect: (site: SiteSummary) => void;
}) {
  return (
    <section className="page-section">
      <div className="site-picker-hero">
        <div>
        <p className="section-label">Workspace Selection</p>
        <h1>Sites</h1>
        <p>Choose the SharePoint site you want to review.</p>
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
      </div>

      {sites.length === 0 && !loadingLabel && (
        <div className="empty-row site-empty-state">
          No SharePoint sites are configured for this app.
        </div>
      )}

      {loadingLabel && <div className="loading-note">{loadingLabel}</div>}
    </section>
  );
}

function ReportsPanel({
  report,
  reportError,
  loadingLabel,
  onRefresh,
}: {
  report: ReportSummary | null;
  reportError: string;
  loadingLabel: string;
  onRefresh: () => void;
}) {
  const generatedAt = report?.generatedAt ? new Date(report.generatedAt).toLocaleString() : "Not generated";

  return (
    <section className="page-section">
      <div className="page-header with-actions">
        <div>
          <p className="section-label">Executive Report</p>
          <h1>Permission Overview</h1>
          <p>Read-only summary across configured SharePoint sites.</p>
        </div>
        <button className="secondary-button" onClick={onRefresh}>
          <RefreshCw size={17} />
          Refresh report
        </button>
      </div>

      {reportError && <div className="auth-error">{reportError}</div>}
      {loadingLabel === "Loading reports" && <div className="loading-note">Loading reports</div>}

      {report && (
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
            <ReportMetric label="Permission rows" value={report.permissions.length} />
          </div>

          <div className="permission-section-title">
            <div>
              <p className="section-label">Permission Inventory</p>
              <h2>Who has access</h2>
            </div>
            <span>{report.permissions.length}</span>
          </div>

          <div className="report-permission-table" role="table" aria-label="Permission inventory">
            <div className="report-permission-head" role="row">
              <span>Principal</span>
              <span>Role</span>
              <span>Source</span>
              <span>Scope</span>
              <span>Tenant</span>
            </div>
            {report.permissions.map((permission) => (
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
            {report.permissions.length === 0 && (
              <div className="empty-row">No permissions found in configured library roots.</div>
            )}
          </div>

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

function ContentExplorer({
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
  const currentSelection = selectedContent && contents.some((item) => item.id === selectedContent.id)
    ? selectedContent
    : contents[0];

  return (
    <section className="page-section">
      <div className="page-header with-actions">
        <div>
          <p className="section-label">SharePoint Workspace</p>
          <h1>{path.at(-1)?.name ?? `${site.name} Contents`}</h1>
          <p>{path.length ? "Open a folder or manage access for this location." : "Review libraries available in this site."}</p>
        </div>
        <button className="icon-button" title="Refresh" onClick={onRefresh}>
          <RefreshCw size={17} />
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
            <span>Modified</span>
          </div>
          {contents.map((item) => (
            <div
              className={`items-row ${currentSelection?.id === item.id ? "selected" : ""}`}
              key={item.id}
              onClick={() => setSelectedContent(item)}
            >
              <button
                className="item-name-button"
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
              <span className="muted" data-label="Modified">{item.modified ?? "Live Graph"}</span>
            </div>
          ))}
          {contents.length === 0 && <div className="empty-row">{loadingLabel || "No items found."}</div>}
        </div>

        <aside className="details-pane">
          {currentSelection ? (
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
            <p className="section-label">Recent Changes</p>
            <div className="audit-list">
              {audit.slice(0, 3).map((entry) => (
                <article className="audit-item" key={entry.id}>
                  <span className="audit-dot" />
                  <div>
                    <strong>{entry.action}</strong>
                    <p>{entry.target}</p>
                    <small>
                      {entry.actor} · {entry.time}
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

function AccessPanel({
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
  onUpdateRole: (permissionId: string, role: AccessRole) => void;
  onRemove: (permissionId: string) => void;
}) {
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
          <p>{item.protected ? "Protected library context. Downloaded Office files remain governed by Rights Management." : "Review direct access for the selected item."}</p>
        </div>
        <div className="header-actions">
          <button className="secondary-button" onClick={onBack}>
            <ArrowLeft size={16} />
            Back to {backLabel}
          </button>
          <button className="icon-button" title="Refresh permissions" onClick={onRefresh}>
            <RefreshCw size={17} />
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
          <span>Direct</span>
        </div>
        <div>
          <strong>{managedPermissions.length}</strong>
          <span>Inherited / managed</span>
        </div>
        <div>
          <strong>{permissions.length}</strong>
          <span>Total visible</span>
        </div>
      </div>

      {permissions.length > 0 && lockedCount > 0 && (
        <div className="info-message">
          <ShieldCheck size={18} />
          <span>
            {editableCount === 0
              ? "No direct permissions are visible here. Add a direct Viewer/Editor below, or manage inherited access at the parent site/library."
              : `${lockedCount} inherited or system-managed permission${lockedCount > 1 ? "s" : ""} ${lockedCount === 1 ? "is" : "are"} shown for context only.`}
          </span>
        </div>
      )}

      {!canManagePermissions && (
        <div className="info-message readonly-message">
          <ShieldCheck size={18} />
          <span>Your app role is read-only here. You can review permissions, but only Admin can grant, change, or remove access.</span>
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
            <strong>Grant direct access</strong>
            <span>Use this for individual users or groups who need access to this item.</span>
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
                External recipient. This app grants direct access without sending an invitation email.
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
          <button className="primary-button" type="submit">
            <Plus size={18} />
            {loadingLabel === "Granting permission" ? "Granting" : "Grant"}
          </button>
        </form>
      )}

      <div className="permission-section-title">
        <div>
          <p className="section-label">Direct Access</p>
          <h2>Editable permissions</h2>
        </div>
        <span>{directPermissions.length}</span>
      </div>

      <PermissionTable
        permissions={directPermissions}
        emptyText={loadingLabel || "No direct permissions found."}
        canManagePermissions={canManagePermissions}
        onUpdateRole={onUpdateRole}
        onRemove={onRemove}
      />

      {managedPermissions.length > 0 && (
        <>
          <div className="permission-section-title muted-title">
            <div>
              <p className="section-label">Inherited And Managed</p>
              <h2>Shown for context</h2>
            </div>
            <span>{managedPermissions.length}</span>
          </div>
          <PermissionTable
            permissions={managedPermissions}
            emptyText=""
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
  canManagePermissions,
  onUpdateRole,
  onRemove,
}: {
  permissions: PermissionEntry[];
  emptyText: string;
  canManagePermissions: boolean;
  onUpdateRole: (permissionId: string, role: AccessRole) => void;
  onRemove: (permissionId: string) => void;
}) {
  return (
    <div className="permission-table" role="table" aria-label="Permissions">
      <div className="table-head" role="row">
        <span>Principal</span>
        <span>Role</span>
        <span>Source</span>
        <span>Activity</span>
        <span>Remove</span>
      </div>
      {permissions.map((permission) => (
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
          <select
            aria-label={`Role for ${permission.displayName}`}
            className={`role-select ${permission.role}`}
            disabled={!canManagePermissions || permission.role === "owner" || permission.canEditRole === false}
            onChange={(event) => onUpdateRole(permission.id, event.target.value as AccessRole)}
            value={permission.role}
          >
            <option value="viewer">{roleLabels.viewer}</option>
            <option value="editor">{roleLabels.editor}</option>
            <option value="owner">{roleLabels.owner}</option>
          </select>
          <span className="muted" data-label="Source">{permission.source}</span>
          <span className="muted" data-label="Activity">{permission.lastActivity}</span>
          <div className="row-actions" data-label="Remove">
            {!canManagePermissions ? (
              <span className="locked-badge">Read-only</span>
            ) : permission.canDelete === false ? (
              <span className="locked-badge">{permission.source === "inherited" ? "Inherited" : "Managed"}</span>
            ) : (
              <button
                className="remove-access-button"
                title="Remove this direct permission"
                type="button"
                onClick={() => onRemove(permission.id)}
              >
                <Trash2 size={15} />
                Remove access
              </button>
            )}
          </div>
        </div>
      ))}
      {permissions.length === 0 && emptyText && <div className="empty-row">{emptyText}</div>}
    </div>
  );
}

function ItemIcon({ item }: { item: ContentItem }) {
  if (item.protected && item.type === "library") return <FileLock2 size={20} />;
  if (item.type === "library") return <Library size={20} />;
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

function getInitials(label: string) {
  const words = label.split(/\s+/).filter(Boolean);
  if (words.length === 1) return words[0].slice(0, 2).toUpperCase();
  return words
    .slice(0, 2)
    .map((word) => word[0])
    .join("")
    .toUpperCase();
}

