"use client";

import type { AccountInfo } from "@azure/msal-browser";
import {
  ArrowLeft,
  ChevronRight,
  Eye,
  File,
  FileLock2,
  Folder,
  Home as HomeIcon,
  Library,
  LogIn,
  LogOut,
  Pencil,
  Plus,
  RefreshCw,
  Search,
  ShieldCheck,
  Trash2,
  UserRound,
  UsersRound,
} from "lucide-react";
import { FormEvent, useMemo, useState } from "react";
import { acquireGraphToken, isAuthConfigured, signInMicrosoft365 } from "@/lib/auth";
import { demoAudit, demoSites } from "@/lib/mock-data";
import {
  GraphSharePointPermissionClient,
  MockSharePointPermissionClient,
  type PermissionDraft,
  type SharePointPermissionClient,
} from "@/lib/graph";
import type { AccessRole, AuditEntry, ContentItem, PermissionEntry, SiteSummary } from "@/lib/types";

const roleLabels: Record<AccessRole, string> = {
  viewer: "Viewer",
  editor: "Editor",
  owner: "Owner",
};

const roleDescriptions: Record<AccessRole, string> = {
  viewer: "Read access through SharePoint. Rights Management still controls protected downloads.",
  editor: "Read and write access through SharePoint.",
  owner: "Full control. Review before assigning.",
};

export default function Home() {
  const [signedIn, setSignedIn] = useState(false);
  const [mode, setMode] = useState<"demo" | "live">("demo");
  const [account, setAccount] = useState<AccountInfo | null>(null);
  const [accountLabel, setAccountLabel] = useState("Admin");
  const [sites, setSites] = useState<SiteSummary[]>(demoSites);
  const [selectedSite, setSelectedSite] = useState<SiteSummary | null>(null);
  const [contents, setContents] = useState<ContentItem[]>([]);
  const [path, setPath] = useState<ContentItem[]>([]);
  const [selectedItem, setSelectedItem] = useState<ContentItem | null>(null);
  const [permissions, setPermissions] = useState<PermissionEntry[]>([]);
  const [audit, setAudit] = useState<AuditEntry[]>(demoAudit);
  const [query, setQuery] = useState("");
  const [newEmail, setNewEmail] = useState("");
  const [newRole, setNewRole] = useState<AccessRole>("viewer");
  const [authError, setAuthError] = useState("");
  const [dataError, setDataError] = useState("");
  const [loadingLabel, setLoadingLabel] = useState("");

  const graphClient = useMemo<SharePointPermissionClient>(() => {
    if (!signedIn || !isAuthConfigured || mode !== "live") return new MockSharePointPermissionClient();
    return new GraphSharePointPermissionClient(() => acquireGraphToken(account));
  }, [account, mode, signedIn]);

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
        setSignedIn(true);
        setMode("demo");
        setAccountLabel("Demo Admin");
        setSites(demoSites);
        return;
      }

      const response = await signInMicrosoft365();
      const nextClient = new GraphSharePointPermissionClient(() => acquireGraphToken(response.account));
      const nextSites = await nextClient.listSites();

      setSignedIn(true);
      setMode("live");
      setAccount(response.account);
      setAccountLabel(response.account?.username ?? "Microsoft 365 Admin");
      setSites(nextSites);
    } catch (error) {
      setAuthError(error instanceof Error ? error.message : "Unable to connect Microsoft 365.");
    } finally {
      setLoadingLabel("");
    }
  }

  function signOut() {
    setSignedIn(false);
    setMode("demo");
    setAccount(null);
    setAccountLabel("Admin");
    setSelectedSite(null);
    setContents([]);
    setPath([]);
    setSelectedItem(null);
    setPermissions([]);
    setQuery("");
    setNewEmail("");
    setAuthError("");
    setDataError("");
  }

  async function chooseSite(site: SiteSummary) {
    setSelectedSite(site);
    setSelectedItem(null);
    setPath([]);
    setQuery("");
    setPermissions([]);
    setDataError("");
    setLoadingLabel("Loading site contents");

    try {
      const nextContents = await graphClient.listContentItems(site.id);
      setContents(nextContents);
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
      setPath((current) => [...current, item]);
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
    } catch (error) {
      setDataError(error instanceof Error ? error.message : "Unable to go back.");
    } finally {
      setLoadingLabel("");
    }
  }

  async function manageAccess(item: ContentItem) {
    setSelectedItem(item);
    setQuery("");
    setDataError("");
    setLoadingLabel("Loading permissions");

    try {
      const nextPermissions = await graphClient.listPermissions(item.id);
      setPermissions(normalizeSharePointPrincipals(nextPermissions, selectedSite?.name ?? "Site"));
    } catch (error) {
      setDataError(error instanceof Error ? error.message : "Unable to load permissions.");
      setPermissions([]);
    } finally {
      setLoadingLabel("");
    }
  }

  async function refreshCurrentView() {
    if (selectedItem) {
      await manageAccess(selectedItem);
      return;
    }

    if (path.length === 0) {
      if (selectedSite) await chooseSite(selectedSite);
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
      if (mode === "live" && graphClient.grantPermission) {
        const created = await graphClient.grantPermission(selectedItem, draft);
        setPermissions((current) => [...created, ...current]);
      } else {
        setPermissions((current) => [
          {
            id: `perm-${Date.now()}`,
            libraryId: selectedItem.id,
            driveId: selectedItem.driveId,
            itemId: selectedItem.itemId,
            displayName: draft.email.split("@")[0],
            email: draft.email,
            type: "user",
            role: draft.role,
            source: "direct",
            tenant: draft.email.endsWith("@baht.net") ? "baht.net" : "external",
            lastActivity: "Just granted",
            canEditRole: true,
            canDelete: true,
          },
          ...current,
        ]);
      }

      addAudit(`Granted ${roleLabels[newRole].toLowerCase()}`, draft.email);
      setNewEmail("");
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
      if (mode === "live" && graphClient.updatePermissionRole) {
        const updated = await graphClient.updatePermissionRole(changed, role);
        setPermissions((current) =>
          current.map((permission) => (permission.id === permissionId ? updated : permission)),
        );
      } else {
        setPermissions((current) =>
          current.map((permission) =>
            permission.id === permissionId ? { ...permission, role, lastActivity: "Role updated" } : permission,
          ),
        );
      }
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
      if (mode === "live" && graphClient.removePermission) {
        await graphClient.removePermission(removed);
      }
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
        actor: mode === "live" ? accountLabel : "Admin",
        action,
        target,
        time: "Just now",
      },
      ...current,
    ]);
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
            <span>{isAuthConfigured ? "Entra app configured" : "Demo mode available"}</span>
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
            className={`sidebar-nav-item ${!selectedSite ? "active" : ""}`}
            onClick={() => {
              setSelectedSite(null);
              setSelectedItem(null);
              setPath([]);
              setContents([]);
              setPermissions([]);
              setQuery("");
            }}
          >
            <HomeIcon size={18} />
            Sites
          </button>
          <button className={`sidebar-nav-item ${selectedSite ? "active" : ""}`} disabled={!selectedSite}>
            <Library size={18} />
            Site contents
          </button>
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
                : selectedSite
                  ? `Browsing: ${selectedSite.name}`
                  : "Select a SharePoint site"}
            </span>
          </div>

          <div className="user-cluster">
            <span className={`status-dot ${mode}`} />
            <span>{mode === "live" ? "Live Graph" : "Demo"}</span>
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

          {!selectedSite ? (
            <SitePicker sites={sites} onSelect={chooseSite} loadingLabel={loadingLabel} />
          ) : selectedItem ? (
            <AccessPanel
              item={selectedItem}
              permissions={visiblePermissions}
              query={query}
              newEmail={newEmail}
              newRole={newRole}
              loadingLabel={loadingLabel}
              onBack={() => {
                setSelectedItem(null);
                setPermissions([]);
                setQuery("");
              }}
              onRefresh={refreshCurrentView}
              onQueryChange={setQuery}
              onEmailChange={setNewEmail}
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
              onBack={goBackOneLevel}
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

      {loadingLabel && <div className="loading-note">{loadingLabel}</div>}
    </section>
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
  onBack,
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
  onBack: () => void;
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
          <button onClick={onRoot}>
            <HomeIcon size={14} />
            Site contents
          </button>
          {path.map((node) => (
            <span key={node.id}>
              <ChevronRight size={14} />
              <strong>{node.name}</strong>
            </span>
          ))}
        </div>
        {path.length > 0 && (
          <button className="secondary-button" onClick={onBack}>
            <ArrowLeft size={16} />
            Back
          </button>
        )}
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
              <span className="muted">{item.type}</span>
              <span className={`policy-badge ${item.protected ? "protected" : ""}`}>{item.rightsPolicy}</span>
              <span className="muted">{item.modified ?? "Live Graph"}</span>
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
  loadingLabel,
  onBack,
  onRefresh,
  onQueryChange,
  onEmailChange,
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
  loadingLabel: string;
  onBack: () => void;
  onRefresh: () => void;
  onQueryChange: (value: string) => void;
  onEmailChange: (value: string) => void;
  onRoleChange: (value: AccessRole) => void;
  onGrant: (event: FormEvent<HTMLFormElement>) => void;
  onUpdateRole: (permissionId: string, role: AccessRole) => void;
  onRemove: (permissionId: string) => void;
}) {
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
            Back
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

      <form className="grant-panel" onSubmit={onGrant}>
        <div className="grant-title">
          <strong>Grant direct access</strong>
          <span>Use this for individual users or groups who need access to this item.</span>
        </div>
        <label>
          <span>User email</span>
          <input
            onChange={(event) => onEmailChange(event.target.value)}
            placeholder="name@baht.net"
            type="email"
            value={newEmail}
          />
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
  onUpdateRole,
  onRemove,
}: {
  permissions: PermissionEntry[];
  emptyText: string;
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
        <span>Actions</span>
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
            disabled={permission.role === "owner" || permission.canEditRole === false}
            onChange={(event) => onUpdateRole(permission.id, event.target.value as AccessRole)}
            value={permission.role}
          >
            <option value="viewer">{roleLabels.viewer}</option>
            <option value="editor">{roleLabels.editor}</option>
            <option value="owner">{roleLabels.owner}</option>
          </select>
          <span className="muted">{permission.source}</span>
          <span className="muted">{permission.lastActivity}</span>
          <div className="row-actions">
            {permission.canEditRole === false && permission.canDelete === false ? (
              <span className="locked-badge">{permission.source === "inherited" ? "Inherited" : "Managed"}</span>
            ) : (
              <>
                <button
                  className="mini-button"
                  disabled={permission.canEditRole === false}
                  title={roleDescriptions.viewer}
                  type="button"
                  onClick={() => onUpdateRole(permission.id, "viewer")}
                >
                  <Eye size={16} />
                </button>
                <button
                  className="mini-button"
                  disabled={permission.canEditRole === false}
                  title={roleDescriptions.editor}
                  type="button"
                  onClick={() => onUpdateRole(permission.id, "editor")}
                >
                  <Pencil size={16} />
                </button>
                <button
                  className="mini-button danger"
                  disabled={permission.canDelete === false}
                  title={permission.canDelete === false ? "Inherited permission cannot be removed here" : "Remove permission"}
                  type="button"
                  onClick={() => onRemove(permission.id)}
                >
                  <Trash2 size={16} />
                </button>
              </>
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

function normalizeSharePointPrincipals(permissions: PermissionEntry[], siteName: string) {
  return permissions.map((permission) => {
    if (!isOpaqueSharePointPrincipal(permission.displayName)) {
      if (isDefaultSharePointGroup(permission, siteName)) {
        return {
          ...permission,
          lastActivity: permission.source === "inherited" ? "Inherited from parent" : "System-managed SharePoint group",
          canEditRole: false,
          canDelete: false,
        };
      }

      return permission;
    }

    const decoded = decodeSharePointPrincipal(permission.displayName);
    const principalId = decoded?.split("_").at(-1);
    const groupLabel = getDefaultSharePointGroupName(permission.role, principalId);

    return {
      ...permission,
      displayName: `${siteName} ${groupLabel}`,
      email: principalId ? `Inherited SharePoint group · Principal ${principalId}` : "Inherited SharePoint group",
      type: "group" as const,
      source: "inherited" as const,
      lastActivity: "Inherited from parent",
      canEditRole: false,
      canDelete: false,
    };
  });
}

function isDefaultSharePointGroup(permission: PermissionEntry, siteName: string) {
  if (permission.type !== "group") return false;
  const normalizedName = permission.displayName.toLowerCase();
  const normalizedSite = siteName.toLowerCase();
  return (
    normalizedName === `${normalizedSite} owners` ||
    normalizedName === `${normalizedSite} members` ||
    normalizedName === `${normalizedSite} visitors`
  );
}

function isOpaqueSharePointPrincipal(value: string) {
  if (value.length < 28 || /\s/.test(value)) return false;
  return /^[A-Za-z0-9_-]+={0,2}$/.test(value);
}

function decodeSharePointPrincipal(value: string) {
  try {
    const normalized = value.replace(/-/g, "+").replace(/_/g, "/");
    const padded = normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "=");
    const decoded = atob(padded).replace(/[^\x20-\x7E_:-]/g, "");
    return decoded.includes("_") ? decoded : undefined;
  } catch {
    return undefined;
  }
}

function getDefaultSharePointGroupName(role: AccessRole, principalId?: string) {
  if (principalId === "3") return "Owners";
  if (principalId === "4") return "Visitors";
  if (principalId === "5") return "Members";
  if (role === "owner") return "Owners";
  if (role === "editor") return "Members";
  return "Visitors";
}
