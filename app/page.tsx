"use client";

import type { AccountInfo } from "@azure/msal-browser";
import {
  BarChart3,
  Home as HomeIcon,
  LogIn,
  LogOut,
  ShieldCheck,
  ScrollText,
} from "lucide-react";
import { FormEvent, useEffect, useMemo, useRef, useState } from "react";
import {
  AccessPanel,
  ContentExplorer,
  PermissionActionDialog,
  SitePicker,
  type PendingPermissionAction,
} from "@/features/admin/AdminPanels";
import { AuditPanel } from "@/features/audit/AuditPanel";
import { ReportsPanel } from "@/features/reviewer/ReportsPanel";
import { acquireGraphToken, appSessionMaxAgeMs, getSignedInAccount, isAuthConfigured, signInMicrosoft365, signOutMicrosoft365 } from "@/lib/auth";
import { isInternalEmail } from "@/lib/app-config";
import { filterContentItemsForRoles, getAccountRoles, getCapabilities, getPrimaryRole, getRoleLabel } from "@/lib/app-roles";
import { createAuditStore, type AuditStore } from "@/lib/features/audit";
import { normalizeSharePointPrincipals, roleLabels } from "@/lib/features/admin";
import { createReviewScopeStore, fallbackReviewScopes, getReviewScopeOwners, type ReviewScope, type ReviewScopeStore } from "@/lib/features/reviewer";
import {
  graphReadScopes,
  graphWriteScopes,
  GraphSharePointPermissionClient,
  type PermissionDraft,
  type SharePointPermissionClient,
} from "@/lib/features/sharepoint-client";
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

const savedSessionViewKey = "spAccess:lastView";
const signedOutMarkerKey = "spAccess:signedOut";
const signedInAtKey = "spAccess:signedInAt";
const auditRecordLimit = 500;

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
  const [permissionLinkNotice, setPermissionLinkNotice] = useState("");
  const [pendingPermissionAction, setPendingPermissionAction] = useState<PendingPermissionAction | null>(null);
  const [userSuggestions, setUserSuggestions] = useState<UserSuggestion[]>([]);
  const [suggestionError, setSuggestionError] = useState("");
  const [suggestionsLoading, setSuggestionsLoading] = useState(false);
  const [reportSummary, setReportSummary] = useState<ReportSummary | null>(null);
  const [reportError, setReportError] = useState("");
  const [reviewScopes, setReviewScopes] = useState<ReviewScope[]>(fallbackReviewScopes);
  const [selectedReviewOwnerEmail, setSelectedReviewOwnerEmail] = useState("");
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
  const reviewScopeStore = useMemo<ReviewScopeStore>(() => {
    return createReviewScopeStore(() => acquireGraphToken(account, undefined, { allowPopup: false }));
  }, [account]);
  const writeGraphClient = useMemo<SharePointPermissionClient>(() => {
    return new GraphSharePointPermissionClient(() => acquireGraphToken(account, graphWriteScopes));
  }, [account]);

  const appRoles = useMemo(() => getAccountRoles(account), [account]);
  const capabilities = useMemo(() => getCapabilities(appRoles), [appRoles]);
  const reviewScopeOwners = useMemo(() => getReviewScopeOwners(reviewScopes), [reviewScopes]);
  const showGlobalDataError = Boolean(dataError && (portalView !== "workspace" || selectedSite));

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
        if (isAppSessionExpired()) {
          await expireAppSession();
          if (!cancelled) {
            setAuthError("Your session expired. Sign in again to continue.");
          }
          return;
        }
        ensureAppSessionStarted();

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
        const nextReviewScopeStore = createReviewScopeStore(() => acquireGraphToken(restoredAccount, undefined, { allowPopup: false }));
        const nextReviewScopes = await loadReviewScopesWithStore(nextReviewScopeStore);
        if (cancelled) return;

        setReviewScopes(nextReviewScopes);
        await restoreSavedSessionView(nextClient, nextAuditStore, nextRoles, nextReviewScopes, getSignedInEmail(restoredAccount));
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
    if (!signedIn) return;

    const remainingMs = getRemainingAppSessionMs();
    const expireCurrentSession = async () => {
      await signOut();
      setAuthError("Your session expired. Sign in again to continue.");
    };

    if (remainingMs <= 0) {
      void expireCurrentSession();
      return;
    }

    const timeout = window.setTimeout(() => {
      void expireCurrentSession();
    }, remainingMs);

    return () => window.clearTimeout(timeout);
    // The timeout intentionally follows the current signed-in session lifetime.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [signedIn]);

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

  async function loadReviewScopesWithStore(store: ReviewScopeStore) {
    try {
      const scopes = await store.list();
      return scopes.length > 0 ? scopes : fallbackReviewScopes;
    } catch {
      return fallbackReviewScopes;
    }
  }

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
      startAppSession();

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
      const nextReviewScopeStore = createReviewScopeStore(() => acquireGraphToken(response.account, undefined, { allowPopup: false }));
      const nextReviewScopes = await loadReviewScopesWithStore(nextReviewScopeStore);
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
      const initialReviewOwnerEmail = getSignedInEmail(response.account);
      const initialReportSummary = initialPortalView === "reports"
        ? await nextClient.getReportSummary({
            ownerEmail: initialReviewOwnerEmail,
            reviewScopes: nextReviewScopes,
          }).catch(() => null)
        : null;
      const initialAuditRecords = initialPortalView === "audit"
        ? await nextAuditStore.list(auditRecordLimit).catch(() => [])
        : [];

      setSignedIn(true);
      setAccount(response.account);
      setAccountLabel(response.account?.username ?? "Microsoft 365 Admin");
      setRoleLabel(getRoleLabel(getPrimaryRole(nextRoles)));
      setSites(nextSites);
      setReviewScopes(nextReviewScopes);
      setSelectedReviewOwnerEmail(initialReviewOwnerEmail);
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
    setPermissionLinkNotice("");
    setAuthError("");
    setDataError("");
    setReportSummary(null);
    setReportError("");
    setReviewScopes(fallbackReviewScopes);
    setSelectedReviewOwnerEmail("");
    setAuditRecords([]);
    setAuditError("");
    window.sessionStorage.setItem(signedOutMarkerKey, "true");
    window.sessionStorage.removeItem(signedInAtKey);
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
    setPermissionLinkNotice("");
    setDataError("");
    setDataConsentRequired(false);
    pushAppHistory({ selectedSite: null, path: [], selectedItem: null });
  }

  async function openReports(ownerEmail = getSignedInEmail(account)) {
    if (!capabilities.canViewReports) return;

    setPortalView("reports");
    setSelectedSite(null);
    setContents([]);
    setPath([]);
    setSelectedItem(null);
    setPermissions([]);
    setQuery("");
    setPermissionLinkNotice("");
    setDataError("");
    setDataConsentRequired(false);
    setReportError("");

    setSelectedReviewOwnerEmail(ownerEmail);
    setLoadingLabel("Loading reports");

    try {
      const nextReviewScopes = await loadReviewScopesWithStore(reviewScopeStore);
      setReviewScopes(nextReviewScopes);

      setReportSummary(await graphClient.getReportSummary({ ownerEmail, reviewScopes: nextReviewScopes }));
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
    setPermissionLinkNotice("");
    setDataError("");
    setDataConsentRequired(false);
    setAuditError("");
    setLoadingLabel("Loading audit");

    try {
      setAuditRecords(await auditStore.list(auditRecordLimit));
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
    setPermissionLinkNotice("");
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
    setPermissionLinkNotice("");
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
    setPermissionLinkNotice("");
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
    setPermissionLinkNotice("");
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
      setPermissionLinkNotice(getPermissionLinkNotice("Access granted", selectedItem, draft.email));
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
      setPermissionLinkNotice(getPermissionLinkNotice("Role updated", selectedItem, changed.email));
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
      const nextReviewScopes = await loadReviewScopesWithStore(reviewScopeStore);
      setSites(nextSites);
      setReviewScopes(nextReviewScopes);

      if (portalView === "reports" && capabilities.canViewReports) {
        const ownerEmail = selectedReviewOwnerEmail || getSignedInEmail(account);
        setSelectedReviewOwnerEmail(ownerEmail);
        setReportSummary(await consentClient.getReportSummary({
          ownerEmail,
          reviewScopes: nextReviewScopes,
        }));
        setReportError("");
      }

      if (portalView === "audit" && capabilities.canViewAudit) {
        setAuditRecords(await auditStore.list(auditRecordLimit));
        setAuditError("");
      }
    } catch (error) {
      handleSharePointDataError(error);
    } finally {
      setLoadingLabel("");
    }
  }

  function handleSharePointDataError(error: unknown) {
    const nextError = getSharePointDataError(error);
    setDataError(nextError.message);
    setDataConsentRequired(nextError.consentRequired);
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

  async function restoreSavedSessionView(
    client: SharePointPermissionClient,
    store: AuditStore,
    roles: ReturnType<typeof getAccountRoles>,
    nextReviewScopes = reviewScopes,
    reviewerEmail = getSignedInEmail(account),
  ) {
    const saved = readSavedSessionView();
    const roleCapabilities = getCapabilities(roles);
    if (!saved) {
      replaceAppHistory({ selectedSite: null, path: [], selectedItem: null });
      if (!roleCapabilities.canManagePermissions && roleCapabilities.canViewReports) {
        setPortalView("reports");
        setSelectedReviewOwnerEmail(reviewerEmail);
        setReportSummary(await client.getReportSummary({ ownerEmail: reviewerEmail, reviewScopes: nextReviewScopes }));
      } else if (!roleCapabilities.canManagePermissions && roleCapabilities.canViewAudit) {
        setPortalView("audit");
        setAuditRecords(await store.list(auditRecordLimit));
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
        setSelectedReviewOwnerEmail(reviewerEmail);
        setReportSummary(await client.getReportSummary({ ownerEmail: reviewerEmail, reviewScopes: nextReviewScopes }));
        return;
      }

      if (restoredPortalView === "audit") {
        setAuditRecords(await store.list(auditRecordLimit));
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
            <button className={`sidebar-nav-item ${portalView === "reports" ? "active" : ""}`} onClick={() => openReports()}>
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
            <span>Connected</span>
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
          {showGlobalDataError && (
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
              reviewOwners={reviewScopeOwners}
              selectedOwnerEmail={selectedReviewOwnerEmail}
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
              shareLink={
                permissionLinkNotice && selectedItem.webUrl
                  ? {
                      message: permissionLinkNotice,
                      url: selectedItem.webUrl,
                    }
                  : undefined
              }
              onBack={leaveAccessPanel}
              onRefresh={refreshCurrentView}
              onQueryChange={setQuery}
              onEmailChange={setNewEmail}
              onSelectUserSuggestion={selectUserSuggestion}
              onRoleChange={setNewRole}
              onGrant={addPermission}
              onUpdateRole={updateRole}
              onRemove={removePermission}
              onCopyLink={copyPermissionLink}
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
          successLink={
            permissionLinkNotice && selectedItem?.webUrl
              ? {
                  message: permissionLinkNotice,
                  url: selectedItem.webUrl,
                }
              : undefined
          }
          onApprovalRequestNoChange={setApprovalRequestNo}
          onCancel={closePermissionConfirmation}
          onConfirm={confirmPermissionAction}
          onCopyLink={copyPermissionLink}
        />
      )}

    </main>
  );
}

function getErrorMessage(error: unknown, fallback = "Unexpected error.") {
  return error instanceof Error ? error.message : fallback;
}

function getSharePointDataError(error: unknown) {
  const message = getErrorMessage(error, "Unable to load SharePoint data.");
  const normalized = message.toLowerCase();
  const consentRequired =
    normalized.includes("additional microsoft graph consent") ||
    normalized.includes("consent_required") ||
    normalized.includes("aadsts65001") ||
    normalized.includes("admin approval") ||
    normalized.includes("interactionrequiredautherror");

  if (consentRequired) {
    return {
      consentRequired: true,
      message:
        "SharePoint data needs Microsoft Graph consent before it can load. Click Request SharePoint access to request User.ReadBasic.All and Sites.Read.All. If Microsoft shows admin approval required, ask an Entra admin to grant admin consent for this app.",
    };
  }

  if (
    normalized.includes("graph 401") ||
    normalized.includes("invalid_grant") ||
    normalized.includes("token") && normalized.includes("expired")
  ) {
    return {
      consentRequired: false,
      message: "Your Microsoft session expired or could not be refreshed. Sign out, then sign in again.",
    };
  }

  if (
    normalized.includes("graph 403") ||
    normalized.includes("accessdenied") ||
    normalized.includes("access denied") ||
    normalized.includes("forbidden")
  ) {
    return {
      consentRequired: false,
      message:
        "You are signed in, but this account cannot read the configured SharePoint sites. Confirm the account has access to the target site and that the tenant admin has granted the required Microsoft Graph permissions.",
    };
  }

  if (
    normalized.includes("graph 404") ||
    normalized.includes("itemnotfound") ||
    normalized.includes("not found")
  ) {
    return {
      consentRequired: false,
      message:
        "The configured SharePoint site could not be found. Check NEXT_PUBLIC_TARGET_SITES, NEXT_PUBLIC_AUDIT_SITE, and the site URL/path in SharePoint.",
    };
  }

  return {
    consentRequired: false,
    message,
  };
}

function extractGraphRequestId(message: string) {
  return message.match(/Request ID:\s*([0-9a-f-]+)/i)?.[1];
}

function getPermissionLinkNotice(action: string, item: ContentItem | null, targetEmail: string) {
  if (!item?.webUrl) return "";
  return `${action} for ${targetEmail}.`;
}

function getSignedInEmail(account: AccountInfo | null | undefined) {
  return account?.username?.trim().toLowerCase() ?? "";
}

async function copyPermissionLink(url: string | undefined) {
  if (!url || typeof navigator === "undefined") return false;

  try {
    await navigator.clipboard.writeText(url);
    return true;
  } catch {
    window.prompt("Copy this SharePoint link", url);
    return false;
  }
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

function startAppSession() {
  window.sessionStorage.setItem(signedInAtKey, Date.now().toString());
}

function ensureAppSessionStarted() {
  if (!window.sessionStorage.getItem(signedInAtKey)) {
    startAppSession();
  }
}

function getAppSessionStartedAt() {
  const raw = window.sessionStorage.getItem(signedInAtKey);
  const parsed = Number.parseInt(raw ?? "", 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : undefined;
}

function getRemainingAppSessionMs() {
  const startedAt = getAppSessionStartedAt();
  if (!startedAt) return appSessionMaxAgeMs;
  return Math.max(appSessionMaxAgeMs - (Date.now() - startedAt), 0);
}

function isAppSessionExpired() {
  return getRemainingAppSessionMs() <= 0;
}

async function expireAppSession() {
  window.sessionStorage.setItem(signedOutMarkerKey, "true");
  window.sessionStorage.removeItem(signedInAtKey);
  clearSavedSessionView();
  await signOutMicrosoft365();
}

