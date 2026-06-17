"use client";

import type { AccountInfo } from "@azure/msal-browser";
import {
  ArrowLeft,
  Check,
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
  ScrollText,
  Trash2,
  UserRound,
  UsersRound,
} from "lucide-react";
import { type CSSProperties, FormEvent, useEffect, useMemo, useRef, useState } from "react";
import { acquireGraphToken, getSignedInAccount, isAuthConfigured, signInMicrosoft365, signOutMicrosoft365 } from "@/lib/auth";
import { isInternalEmail, tenantDomain } from "@/lib/app-config";
import { filterContentItemsForRoles, getAccountRoles, getCapabilities, getPrimaryRole, getRoleLabel } from "@/lib/app-roles";
import { createAuditStore } from "@/lib/audit-store-factory";
import { normalizeSharePointPrincipals } from "@/lib/permission-normalization";
import {
  graphReadScopes,
  graphWriteScopes,
  GraphSharePointPermissionClient,
  type PermissionDraft,
  type SharePointPermissionClient,
} from "@/lib/graph";
import type { AuditStore } from "@/lib/audit-store";
import type {
  AccessRole,
  AuditEntry,
  AuditLogRecord,
  AuditLogAction,
  AuditLogStatus,
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

type PortalView = "workspace" | "reports" | "audit";

type SavedSessionView = {
  portalView: PortalView;
  view: AppHistoryView;
};

type PendingPermissionAction =
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

const savedSessionViewKey = "spAccess:lastView";
const signedOutMarkerKey = "spAccess:signedOut";

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
  const [auditRecords, setAuditRecords] = useState<AuditLogRecord[]>([]);
  const [auditError, setAuditError] = useState("");
  const [query, setQuery] = useState("");
  const [newEmail, setNewEmail] = useState("");
  const [newRole, setNewRole] = useState<AccessRole>("viewer");
  const [approvalRequestNo, setApprovalRequestNo] = useState("");
  const [pendingPermissionAction, setPendingPermissionAction] = useState<PendingPermissionAction | null>(null);
  const [userSuggestions, setUserSuggestions] = useState<UserSuggestion[]>([]);
  const [suggestionError, setSuggestionError] = useState("");
  const [suggestionsLoading, setSuggestionsLoading] = useState(false);
  const [reportSummary, setReportSummary] = useState<ReportSummary | null>(null);
  const [reportError, setReportError] = useState("");
  const [authError, setAuthError] = useState("");
  const [dataError, setDataError] = useState("");
  const [dataConsentRequired, setDataConsentRequired] = useState(false);
  const [loadingLabel, setLoadingLabel] = useState("");
  const [restoringSession, setRestoringSession] = useState(true);
  const restoringHistoryRef = useRef(false);

  const graphClient = useMemo<SharePointPermissionClient>(() => {
    return new GraphSharePointPermissionClient(() => acquireGraphToken(account, undefined, { allowPopup: false }));
  }, [account]);
  const auditStore = useMemo<AuditStore>(() => {
    return createAuditStore(() => acquireGraphToken(account, undefined, { allowPopup: false }));
  }, [account]);
  const writeGraphClient = useMemo<SharePointPermissionClient>(() => {
    return new GraphSharePointPermissionClient(() => acquireGraphToken(account, graphWriteScopes));
  }, [account]);

  const appRoles = useMemo(() => getAccountRoles(account), [account]);
  const capabilities = useMemo(() => getCapabilities(appRoles), [appRoles]);

  useEffect(() => {
    let cancelled = false;

    async function restoreSignedInSession() {
      if (!isAuthConfigured) {
        setRestoringSession(false);
        return;
      }

      if (window.sessionStorage.getItem(signedOutMarkerKey) === "true") {
        setRestoringSession(false);
        return;
      }

      setLoadingLabel("Restoring session");

      try {
        const restoredAccount = await getSignedInAccount();
        if (!restoredAccount || cancelled) return;

        const nextRoles = getAccountRoles(restoredAccount);
        if (nextRoles.length === 0) return;

        const nextClient = new GraphSharePointPermissionClient(() =>
          acquireGraphToken(restoredAccount, undefined, { allowPopup: false }),
        );
        const nextSites = await nextClient.listSites().catch((error) => {
          handleSharePointDataError(error);
          return [];
        });
        if (cancelled) return;

        setSignedIn(true);
        setAccount(restoredAccount);
        setAccountLabel(restoredAccount.username ?? "Microsoft 365 Admin");
        setRoleLabel(getRoleLabel(getPrimaryRole(nextRoles)));
        setSites(nextSites);

        const nextAuditStore = createAuditStore(() => acquireGraphToken(restoredAccount, undefined, { allowPopup: false }));
        await restoreSavedSessionView(nextClient, nextAuditStore, nextRoles);
      } catch {
        clearSavedSessionView();
      } finally {
        if (!cancelled) {
          setLoadingLabel("");
          setRestoringSession(false);
        }
      }
    }

    void restoreSignedInSession();

    return () => {
      cancelled = true;
    };
    // Restore runs once on browser refresh using the MSAL session cache.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

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
    if (!signedIn || restoringSession) return;
    saveSessionView({ portalView, view: getHistoryView() });
    // Persists only navigation state; data is reloaded from Graph after refresh.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [signedIn, restoringSession, portalView, selectedSite, path, selectedItem]);

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

      window.sessionStorage.removeItem(signedOutMarkerKey);
      const response = await signInMicrosoft365();
      if (!response) return;

      const nextRoles = getAccountRoles(response.account);
      if (nextRoles.length === 0) {
        setAuthError("Your account is signed in but has no app role assigned. Ask an administrator to assign Admin, Reviewer, InternalUser, GuestUser, or ExecutiveUser.");
        return;
      }
      const nextCapabilities = getCapabilities(nextRoles);

      const nextClient = new GraphSharePointPermissionClient(() =>
        acquireGraphToken(response.account, undefined, { allowPopup: false }),
      );
      const nextSites = await nextClient.listSites().catch((error) => {
        handleSharePointDataError(error);
        return [];
      });
      const nextAuditStore = createAuditStore(() => acquireGraphToken(response.account, undefined, { allowPopup: false }));
      void writeAuditWithStore(nextAuditStore, {
        action: "Login",
        status: "Success",
        actorEmail: response.account?.username ?? "Unknown",
        actorName: response.account?.name ?? response.account?.username ?? "Microsoft 365 user",
        actorRole: getRoleLabel(getPrimaryRole(nextRoles)),
      });
      const initialPortalView: PortalView = nextCapabilities.canManagePermissions
        ? "workspace"
        : nextCapabilities.canViewReports
          ? "reports"
          : nextCapabilities.canViewAudit
            ? "audit"
            : "workspace";
      const initialReportSummary = initialPortalView === "reports"
        ? await nextClient.getReportSummary().catch(() => null)
        : null;
      const initialAuditRecords = initialPortalView === "audit"
        ? await nextAuditStore.list(100).catch(() => [])
        : [];

      setSignedIn(true);
      setAccount(response.account);
      setAccountLabel(response.account?.username ?? "Microsoft 365 Admin");
      setRoleLabel(getRoleLabel(getPrimaryRole(nextRoles)));
      setSites(nextSites);
      setPortalView(initialPortalView);
      setReportSummary(initialReportSummary);
      setAuditRecords(initialAuditRecords);
      replaceAppHistory({ selectedSite: null, path: [], selectedItem: null });
      saveSessionView({ portalView: initialPortalView, view: { selectedSite: null, path: [], selectedItem: null } });
    } catch (error) {
      setAuthError(error instanceof Error ? error.message : "Unable to connect Microsoft 365.");
    } finally {
      setLoadingLabel("");
    }
  }

  async function signOut() {
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
    setApprovalRequestNo("");
    setAuthError("");
    setDataError("");
    setReportSummary(null);
    setReportError("");
    setAuditRecords([]);
    setAuditError("");
    window.sessionStorage.setItem(signedOutMarkerKey, "true");
    clearSavedSessionView();
    replaceAppHistory({ selectedSite: null, path: [], selectedItem: null });

    try {
      await signOutMicrosoft365(account);
    } catch {
      // Local session state is already cleared; MSAL may fail if the popup is blocked.
    }
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
    setDataConsentRequired(false);
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
    setDataConsentRequired(false);
    setReportError("");
    setLoadingLabel("Loading reports");

    try {
      setReportSummary(await graphClient.getReportSummary());
      void writeAudit({
        action: "RefreshReport",
        status: "Success",
      });
    } catch (error) {
      setReportSummary(null);
      setReportError(error instanceof Error ? error.message : "Unable to load reports.");
      void writeAudit({
        action: "RefreshReport",
        status: "Failed",
        errorMessage: getErrorMessage(error),
      });
    } finally {
      setLoadingLabel("");
    }
  }

  async function openAudit() {
    if (!capabilities.canViewAudit) return;

    setPortalView("audit");
    setSelectedSite(null);
    setContents([]);
    setPath([]);
    setSelectedItem(null);
    setPermissions([]);
    setQuery("");
    setDataError("");
    setDataConsentRequired(false);
    setAuditError("");
    setLoadingLabel("Loading audit");

    try {
      setAuditRecords(await auditStore.list(100));
    } catch (error) {
      setAuditRecords([]);
      setAuditError(error instanceof Error ? error.message : "Unable to load audit trail.");
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
    setDataConsentRequired(false);
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
    setDataConsentRequired(false);
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

    openPermissionConfirmation({
      type: "grant",
      draft,
    });
  }

  function updateRole(permissionId: string, role: Exclude<AccessRole, "owner">) {
    const changed = permissions.find((permission) => permission.id === permissionId);
    if (!changed || changed.role === role) return;

    openPermissionConfirmation({
      type: "update",
      permission: changed,
      role,
    });
  }

  function removePermission(permissionId: string) {
    const removed = permissions.find((permission) => permission.id === permissionId);
    if (!removed) return;

    openPermissionConfirmation({
      type: "remove",
      permission: removed,
    });
  }

  function openPermissionConfirmation(action: PendingPermissionAction) {
    setDataError("");
    setApprovalRequestNo("");
    setPendingPermissionAction(action);
  }

  function closePermissionConfirmation() {
    if (isPermissionActionLoading()) return;
    setPendingPermissionAction(null);
    setApprovalRequestNo("");
    setDataError("");
  }

  async function confirmPermissionAction(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!pendingPermissionAction) return;
    const approvedRequestNo = approvalRequestNo.trim();
    if (!approvedRequestNo) {
      setDataError("Approval request number is required before changing permissions.");
      return;
    }

    if (pendingPermissionAction.type === "grant") {
      await grantPermission(pendingPermissionAction.draft, approvedRequestNo);
      return;
    }

    if (pendingPermissionAction.type === "update") {
      await applyRoleUpdate(pendingPermissionAction.permission, pendingPermissionAction.role, approvedRequestNo);
      return;
    }

    await applyPermissionRemoval(pendingPermissionAction.permission, approvedRequestNo);
  }

  async function grantPermission(draft: PermissionDraft, approvedRequestNo: string) {
    if (!selectedItem) return;

    setDataError("");
    setLoadingLabel("Granting permission");

    try {
      const created = await writeGraphClient.grantPermission(selectedItem, draft);
      setPermissions((current) => [...created, ...current]);

      addAudit(`Granted ${roleLabels[newRole].toLowerCase()}`, draft.email, "Success");
      void writeAudit({
        action: "GrantAccess",
        status: "Success",
        approvalRequestNo: approvedRequestNo,
        targetEmail: draft.email,
        targetName: draft.displayName,
        permissionRole: draft.role,
        tenantType: isInternalEmail(draft.email) ? "internal" : "external",
      });
      setNewEmail("");
      setUserSuggestions([]);
      setPendingPermissionAction(null);
      setApprovalRequestNo("");
    } catch (error) {
      const message = getErrorMessage(error, "Unable to grant permission.");
      setDataError(message);
      addAudit("Grant failed", draft.email, "Failed");
      void writeAudit({
        action: "GrantAccess",
        status: "Failed",
        approvalRequestNo: approvedRequestNo,
        targetEmail: draft.email,
        targetName: draft.displayName,
        permissionRole: draft.role,
        tenantType: isInternalEmail(draft.email) ? "internal" : "external",
        errorMessage: message,
        graphRequestId: extractGraphRequestId(message),
      });
    } finally {
      setLoadingLabel("");
    }
  }

  async function applyRoleUpdate(changed: PermissionEntry, role: Exclude<AccessRole, "owner">, approvedRequestNo: string) {
    setDataError("");
    setLoadingLabel("Updating role");

    try {
      const updated = await writeGraphClient.updatePermissionRole(changed, role);
      setPermissions((current) =>
        current.map((permission) => (permission.id === changed.id ? updated : permission)),
      );
      addAudit(`Changed role to ${roleLabels[role]}`, changed.email, "Success");
      void writeAudit({
        action: "UpdateRole",
        status: "Success",
        approvalRequestNo: approvedRequestNo,
        targetEmail: changed.email,
        targetName: changed.displayName,
        permissionRole: role,
        previousRole: changed.role,
        source: changed.source,
        tenantType: changed.tenant,
      });
      setPendingPermissionAction(null);
      setApprovalRequestNo("");
    } catch (error) {
      const message = getErrorMessage(error, "Unable to update role.");
      setDataError(message);
      addAudit("Role update failed", changed.email, "Failed");
      void writeAudit({
        action: "UpdateRole",
        status: "Failed",
        approvalRequestNo: approvedRequestNo,
        targetEmail: changed.email,
        targetName: changed.displayName,
        permissionRole: role,
        previousRole: changed.role,
        source: changed.source,
        tenantType: changed.tenant,
        errorMessage: message,
        graphRequestId: extractGraphRequestId(message),
      });
    } finally {
      setLoadingLabel("");
    }
  }

  async function applyPermissionRemoval(removed: PermissionEntry, approvedRequestNo: string) {
    setDataError("");
    setLoadingLabel("Removing permission");

    try {
      await writeGraphClient.removePermission(removed);
      setPermissions((current) => current.filter((permission) => permission.id !== removed.id));
      addAudit("Removed permission", removed.email, "Success");
      void writeAudit({
        action: "RemoveAccess",
        status: "Success",
        approvalRequestNo: approvedRequestNo,
        targetEmail: removed.email,
        targetName: removed.displayName,
        permissionRole: removed.role,
        source: removed.source,
        tenantType: removed.tenant,
      });
      setPendingPermissionAction(null);
      setApprovalRequestNo("");
    } catch (error) {
      const message = getErrorMessage(error, "Unable to remove permission.");
      setDataError(message);
      addAudit("Remove failed", removed.email, "Failed");
      void writeAudit({
        action: "RemoveAccess",
        status: "Failed",
        approvalRequestNo: approvedRequestNo,
        targetEmail: removed.email,
        targetName: removed.displayName,
        permissionRole: removed.role,
        source: removed.source,
        tenantType: removed.tenant,
        errorMessage: message,
        graphRequestId: extractGraphRequestId(message),
      });
    } finally {
      setLoadingLabel("");
    }
  }

  function isPermissionActionLoading() {
    return loadingLabel === "Granting permission" || loadingLabel === "Updating role" || loadingLabel === "Removing permission";
  }

  function addAudit(action: string, target: string, status: AuditLogStatus = "Success") {
    setAudit((current) => [
      {
        id: `audit-${Date.now()}`,
        actor: accountLabel,
        action,
        target,
        time: "Just now",
        status,
      },
      ...current,
    ]);
  }

  async function requestSharePointDataAccess() {
    if (!account) return;

    setDataError("");
    setDataConsentRequired(false);
    setLoadingLabel("Requesting SharePoint access");

    try {
      const consentClient = new GraphSharePointPermissionClient(() => acquireGraphToken(account, graphReadScopes));
      const nextSites = await consentClient.listSites();
      setSites(nextSites);

      if (portalView === "reports" && capabilities.canViewReports) {
        setReportSummary(await consentClient.getReportSummary());
        setReportError("");
      }

      if (portalView === "audit" && capabilities.canViewAudit) {
        setAuditRecords(await auditStore.list(100));
        setAuditError("");
      }
    } catch (error) {
      handleSharePointDataError(error);
    } finally {
      setLoadingLabel("");
    }
  }

  function handleSharePointDataError(error: unknown) {
    const message = getErrorMessage(error, "Unable to load SharePoint data.");
    setDataError(message);
    setDataConsentRequired(message.includes("Additional Microsoft Graph consent"));
  }

  function writeAudit(entry: {
    action: AuditLogAction;
    status: AuditLogStatus;
    targetEmail?: string;
    targetName?: string;
    permissionRole?: AccessRole;
    previousRole?: AccessRole;
    approvalRequestNo?: string;
    source?: PermissionEntry["source"];
    tenantType?: PermissionEntry["tenant"];
    errorMessage?: string;
    graphRequestId?: string;
  }) {
    return writeAuditWithStore(auditStore, {
      ...entry,
      actorEmail: account?.username ?? accountLabel,
      actorName: account?.name ?? accountLabel,
      actorRole: roleLabel,
      approvalRequestNo: entry.approvalRequestNo,
      siteId: selectedSite?.id ?? selectedItem?.siteId,
      siteName: selectedSite?.name,
      libraryName: selectedItem?.name ?? path.at(-1)?.name,
      itemId: selectedItem?.itemId ?? selectedItem?.id,
    });
  }

  async function writeAuditWithStore(store: AuditStore, entry: Parameters<AuditStore["write"]>[0]) {
    try {
      await store.write(entry);
    } catch (error) {
      addAudit("Audit log failed", getErrorMessage(error, "Unable to write SharePoint audit log."), "Failed");
    }
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

  async function restoreSavedSessionView(client: SharePointPermissionClient, store: AuditStore, roles: ReturnType<typeof getAccountRoles>) {
    const saved = readSavedSessionView();
    const roleCapabilities = getCapabilities(roles);
    if (!saved) {
      replaceAppHistory({ selectedSite: null, path: [], selectedItem: null });
      if (!roleCapabilities.canManagePermissions && roleCapabilities.canViewReports) {
        setPortalView("reports");
        setReportSummary(await client.getReportSummary());
      } else if (!roleCapabilities.canManagePermissions && roleCapabilities.canViewAudit) {
        setPortalView("audit");
        setAuditRecords(await store.list(100));
      }
      return;
    }

    restoringHistoryRef.current = true;
    const restoredPortalView = saved.portalView === "workspace" && !roleCapabilities.canManagePermissions
      ? roleCapabilities.canViewReports
        ? "reports"
        : roleCapabilities.canViewAudit
          ? "audit"
          : "workspace"
      : saved.portalView;
    setPortalView(restoredPortalView);
    setSelectedSite(saved.view.selectedSite);
    setPath(saved.view.path);
    setSelectedItem(saved.view.selectedItem);

    try {
      if (restoredPortalView === "reports") {
        setReportSummary(await client.getReportSummary());
        return;
      }

      if (restoredPortalView === "audit") {
        setAuditRecords(await store.list(100));
        return;
      }

      if (!saved.view.selectedSite) {
        setContents([]);
        setPermissions([]);
        replaceAppHistory(saved.view);
        return;
      }

      const currentFolder = saved.view.path.at(-1);
      const nextContents = currentFolder
        ? await client.listChildren(currentFolder)
        : filterContentItemsForRoles(await client.listContentItems(saved.view.selectedSite.id), roles);

      setContents(nextContents);

      if (saved.view.selectedItem) {
        const nextPermissions = await client.listPermissions(saved.view.selectedItem.id);
        setPermissions(normalizeSharePointPrincipals(nextPermissions, saved.view.selectedSite.name));
      } else {
        setPermissions([]);
      }

      replaceAppHistory(saved.view);
    } finally {
      restoringHistoryRef.current = false;
    }
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

          <button className="login-button auth-login-button" disabled={restoringSession} onClick={connectMicrosoft365}>
            <LogIn size={18} />
            {restoringSession
              ? "Restoring session"
              : loadingLabel === "Connecting Microsoft 365"
                ? "Connecting"
                : "Sign in with Microsoft 365"}
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
          {capabilities.canManagePermissions && (
            <button
              className={`sidebar-nav-item ${portalView === "workspace" ? "active" : ""}`}
              onClick={returnToSites}
            >
              <HomeIcon size={18} />
              Admin
            </button>
          )}
          {capabilities.canViewReports && (
            <button className={`sidebar-nav-item ${portalView === "reports" ? "active" : ""}`} onClick={openReports}>
              <BarChart3 size={18} />
              Reviewer
            </button>
          )}
          {capabilities.canViewAudit && (
            <button className={`sidebar-nav-item ${portalView === "audit" ? "active" : ""}`} onClick={openAudit}>
              <ScrollText size={18} />
              Audit
            </button>
          )}
        </nav>

        <section className="sidebar-account-card" aria-label="Signed in user">
          <div className="sidebar-account-status">
            <span className="status-dot live" />
            <span>Live Graph</span>
          </div>
          <div className="sidebar-account-main">
            <button className="avatar-button" title={accountLabel}>
              {accountLabel.charAt(0).toUpperCase()}
            </button>
            <div>
              <strong>{accountLabel}</strong>
              <span>{account?.username ?? accountLabel}</span>
            </div>
          </div>
          <div className="sidebar-account-actions">
            <span className={`role-pill ${capabilities.isReadOnly ? "readonly" : ""}`}>{roleLabel}</span>
            <button className="icon-button" title="Sign out" onClick={signOut}>
              <LogOut size={17} />
            </button>
          </div>
        </section>
      </aside>

      <section className="portal-main">
        <div className="portal-content">
          {dataError && (
            <div className="auth-error action-error">
              <span>{dataError}</span>
              {dataConsentRequired && (
                <button className="secondary-button" disabled={loadingLabel === "Requesting SharePoint access"} onClick={requestSharePointDataAccess}>
                  <ShieldCheck size={16} />
                  {loadingLabel === "Requesting SharePoint access" ? "Requesting access" : "Request SharePoint access"}
                </button>
              )}
            </div>
          )}

          {portalView === "reports" ? (
            <ReportsPanel
              loadingLabel={loadingLabel}
              report={reportSummary}
              reportError={reportError}
              onRefresh={openReports}
            />
          ) : portalView === "audit" ? (
            <AuditPanel
              auditRecords={auditRecords}
              auditError={auditError}
              loadingLabel={loadingLabel}
              onRefresh={openAudit}
            />
          ) : !selectedSite ? (
            <SitePicker
              sites={sites}
              onSelect={chooseSite}
              loadingLabel={loadingLabel}
              dataError={dataError}
              dataConsentRequired={dataConsentRequired}
              onRequestDataAccess={requestSharePointDataAccess}
            />
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

      {pendingPermissionAction && (
        <PermissionActionDialog
          action={pendingPermissionAction}
          approvalRequestNo={approvalRequestNo}
          error={dataError}
          isSubmitting={isPermissionActionLoading()}
          onApprovalRequestNoChange={setApprovalRequestNo}
          onCancel={closePermissionConfirmation}
          onConfirm={confirmPermissionAction}
        />
      )}

    </main>
  );
}

function SitePicker({
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
            <button className="secondary-button" onClick={onRequestDataAccess}>
              <ShieldCheck size={16} />
              Request SharePoint access
            </button>
          )}
        </div>
      )}

      {isLoadingSites && <div className="loading-note">{loadingLabel}</div>}
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
  const [showAllPermissions, setShowAllPermissions] = useState(false);
  const generatedAt = report?.generatedAt ? new Date(report.generatedAt).toLocaleString() : "Not generated";
  const isLoadingReport = loadingLabel === "Loading reports";
  const visibleReportPermissions = showAllPermissions ? report?.permissions ?? [] : report?.permissions.slice(0, 8) ?? [];
  const hiddenPermissionCount = Math.max((report?.permissions.length ?? 0) - visibleReportPermissions.length, 0);

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
            <ReportMetric label="Permission rows" value={report.permissions.length} />
          </div>

          <div className="permission-section-title">
            <div>
              <p className="section-label">Permission Inventory</p>
              <h2>Who has access</h2>
            </div>
            <div className="section-title-actions">
              <span>{report.permissions.length}</span>
              {report.permissions.length > 8 && (
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
            {report.permissions.length === 0 && (
              <div className="empty-row">No permissions found in configured library roots.</div>
            )}
          </div>
          {hiddenPermissionCount > 0 && (
            <p className="table-footnote">
              Showing {visibleReportPermissions.length} of {report.permissions.length}. Use See all for the full inventory.
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

function AuditPanel({
  auditRecords,
  auditError,
  loadingLabel,
  onRefresh,
}: {
  auditRecords: AuditLogRecord[];
  auditError: string;
  loadingLabel: string;
  onRefresh: () => void;
}) {
  const isLoadingAudit = loadingLabel === "Loading audit";

  return (
    <section className="page-section">
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

      {isLoadingAudit ? (
        <div className="audit-table">
          <TableSkeleton columns={7} rows={6} />
        </div>
      ) : (
        <div className="audit-table" role="table" aria-label="Permission audit trail">
          <div className="audit-table-head" role="row">
            <span>Time</span>
            <span>Action</span>
            <span>Actor</span>
            <span>Target</span>
            <span>Scope</span>
            <span>Request no.</span>
            <span>Status</span>
          </div>
          {auditRecords.map((entry) => (
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
          {auditRecords.length === 0 && !auditError && (
            <div className="empty-row">No audit entries found.</div>
          )}
        </div>
      )}
    </section>
  );
}

function PermissionActionDialog({
  action,
  approvalRequestNo,
  error,
  isSubmitting,
  onApprovalRequestNoChange,
  onCancel,
  onConfirm,
}: {
  action: PendingPermissionAction;
  approvalRequestNo: string;
  error: string;
  isSubmitting: boolean;
  onApprovalRequestNoChange: (value: string) => void;
  onCancel: () => void;
  onConfirm: (event: FormEvent<HTMLFormElement>) => void;
}) {
  const summary = getPermissionActionSummary(action);

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
            onChange={(event) => onApprovalRequestNoChange(event.target.value)}
            placeholder="e.g. REQ-2026-0001"
            value={approvalRequestNo}
          />
        </label>

        {error && <div className="auth-error confirm-error">{error}</div>}

        <div className="confirm-actions">
          <button className="secondary-button" disabled={isSubmitting} type="button" onClick={onCancel}>
            Cancel
          </button>
          <button className={`primary-button ${action.type === "remove" ? "danger-primary" : ""}`} disabled={isSubmitting || !approvalRequestNo.trim()} type="submit">
            {action.type === "remove" ? <Trash2 size={17} /> : <Check size={17} />}
            {isSubmitting ? summary.submittingLabel : summary.confirmLabel}
          </button>
        </div>
      </form>
    </div>
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
            <span>Modified</span>
          </div>
          {isLoadingContents ? (
            <TableSkeleton columns={4} rows={5} />
          ) : contents.map((item) => (
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
          <p>{item.protected ? "Protected library context. Downloaded Office files remain governed by Rights Management." : "Review direct access for the selected item."}</p>
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
          <button className="primary-button" disabled={loadingLabel === "Granting permission"} type="submit">
            <Plus size={18} />
            {loadingLabel === "Granting permission" ? "Granting" : "Review grant"}
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
        emptyText="No direct permissions found."
        isLoading={isLoadingPermissions}
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
        <span>Source</span>
        <span>Activity</span>
        <span>Action</span>
      </div>
      {isLoading ? (
        <TableSkeleton columns={5} rows={4} />
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
          <span className="muted" data-label="Source">{permission.source}</span>
          <span className="muted" data-label="Activity">{permission.lastActivity}</span>
          <div className="row-actions" data-label="Action">
            {!canManagePermissions ? (
              <span className="locked-badge">Read-only</span>
            ) : permission.canDelete === false ? (
              <span className="locked-badge">{permission.source === "inherited" ? "Inherited" : "Managed"}</span>
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
                  title="Remove this direct permission"
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

function ReportSkeleton() {
  return (
    <div className="skeleton-stack" aria-label="Loading report">
      <div className="report-meta skeleton-meta">
        <span className="skeleton-line short" />
        <strong className="skeleton-line medium" />
      </div>
      <div className="report-metrics">
        {Array.from({ length: 8 }).map((_, index) => (
          <div className="report-metric skeleton-metric" key={`metric-skeleton-${index}`}>
            <span className="skeleton-line short" />
            <span className="skeleton-line medium" />
          </div>
        ))}
      </div>
      <TableSkeleton columns={5} rows={5} />
    </div>
  );
}

function TableSkeleton({ columns, rows }: { columns: number; rows: number }) {
  return (
    <>
      {Array.from({ length: rows }).map((_, rowIndex) => (
        <div
          className="skeleton-table-row"
          role="row"
          key={`skeleton-row-${rowIndex}`}
          style={{ "--skeleton-columns": columns } as CSSProperties}
        >
          {Array.from({ length: columns }).map((__, columnIndex) => (
            <span
              className={`skeleton-line ${columnIndex === 0 ? "wide" : columnIndex % 2 === 0 ? "medium" : "short"}`}
              key={`skeleton-cell-${rowIndex}-${columnIndex}`}
            />
          ))}
        </div>
      ))}
    </>
  );
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

function getPermissionActionSummary(action: PendingPermissionAction) {
  if (action.type === "grant") {
    return {
      title: "Confirm grant access",
      description: "Enter the approved request number before granting direct access.",
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
    description: "Enter the approved request number before removing this direct permission.",
    targetName: action.permission.displayName,
    targetEmail: action.permission.email,
    change: "Remove direct access",
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

function getErrorMessage(error: unknown, fallback = "Unexpected error.") {
  return error instanceof Error ? error.message : fallback;
}

function extractGraphRequestId(message: string) {
  return message.match(/Request ID:\s*([0-9a-f-]+)/i)?.[1];
}

function readSavedSessionView(): SavedSessionView | undefined {
  try {
    const raw = window.sessionStorage.getItem(savedSessionViewKey);
    return raw ? (JSON.parse(raw) as SavedSessionView) : undefined;
  } catch {
    clearSavedSessionView();
    return undefined;
  }
}

function saveSessionView(view: SavedSessionView) {
  window.sessionStorage.setItem(savedSessionViewKey, JSON.stringify(view));
}

function clearSavedSessionView() {
  window.sessionStorage.removeItem(savedSessionViewKey);
}

