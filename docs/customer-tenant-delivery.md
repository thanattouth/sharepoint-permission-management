# Customer Tenant Delivery Guide

This document lists what a customer must prepare before the SharePoint Permission Management app can be deployed to their Microsoft 365 tenant.

Use this guide when moving the app from an internal/demo tenant to a customer tenant.

## 1. Delivery Overview

The app is a tenant-scoped Microsoft 365 web app. It signs users in with Microsoft Entra ID, reads configured SharePoint Online sites through Microsoft Graph, lets Admin users grant/update/remove direct permissions, and stores audit evidence in a SharePoint List.

For a customer delivery, the customer must provide:

- A Microsoft Entra tenant where the app will be registered.
- SharePoint Online sites and document libraries that the app is allowed to manage/review.
- Entra users or groups that will receive app roles.
- Admin consent for the required Microsoft Graph delegated permissions.
- A SharePoint site/list location for audit logs.
- A SharePoint List for reviewer owner-to-library scope mappings, if reviewer scoping is required.

## 2. Customer Prerequisites

Ask the customer to confirm these prerequisites before implementation.

| Area | Required preparation |
| --- | --- |
| Microsoft 365 tenant | Customer must have Microsoft Entra ID and SharePoint Online. |
| Tenant admin access | A Global Administrator, Cloud Application Administrator, or Application Administrator is required for app registration and consent. |
| SharePoint admin access | A SharePoint Administrator or site owner is required to confirm target sites/libraries and sharing settings. |
| App hosting | Confirm the final web app URL, for example `https://<app-name>.azurewebsites.net`. |
| User groups | Confirm the Entra users/groups that should receive app roles. |
| SharePoint targets | Confirm hostnames and site paths, for example `contoso.sharepoint.com:/sites/Finance`. |
| Protected libraries | Confirm which library names should be treated as protected, for example `Confidential,Secret`. |
| Internal domains | Confirm all internal email domains, for example `contoso.com,contoso.onmicrosoft.com`. |
| External file access | Confirm whether external recipients must open shared files from invitation email, copied link, or both. SharePoint and Entra external sharing policies must allow this. |

## 3. Information We Need From The Customer

Collect these values from the customer.

| Item | Example | Notes |
| --- | --- | --- |
| Tenant ID | `00000000-0000-0000-0000-000000000000` | Microsoft Entra tenant ID. |
| Primary tenant domain | `contoso.com` | Used for UI hints and internal/external classification. |
| Internal domains | `contoso.com,contoso.onmicrosoft.com` | Used to classify permission rows as internal or external. |
| App URL | `https://sp-access-contoso.azurewebsites.net` | Must be added as an SPA redirect URI. |
| Target SharePoint sites | `contoso.sharepoint.com:/sites/Finance` | Comma-separated list. |
| Protected library names | `Confidential,Secret` | Exact document library display names. |
| Audit site | `contoso.sharepoint.com:/sites/Governance` | Central site where all audit entries are written. |
| Audit list name | `PermissionAuditLog` | The app can create this list on first write if the Admin has permission. |
| Managed site list | `ManagedSites` | Create this central list once on the audit site so Admin-added SharePoint sites are stored by convention without extra app settings. |
| Review scope site | `contoso.sharepoint.com:/sites/Governance` | Site containing the reviewer mapping list. |
| Review scope list name | `PermissionReviewScopes` | Must exist before reviewer scope loading. |
| Review scan limit | `2000` | Maximum items scanned in one review refresh. |

## 4. Microsoft Entra App Registration

Create one App Registration in the customer's tenant.

Recommended settings:

- Name: `SharePoint Permission Management`
- Supported account types: `Accounts in this organizational directory only`
- Platform type: `Single-page application (SPA)`
- Redirect URI: the final app URL origin, for example `https://sp-access-contoso.azurewebsites.net`
- Optional local redirect URI for testing: `http://localhost:3000`

Do not use My Apps launcher URLs as redirect URIs. My Apps should be configured as a linked tile that opens the web app URL. The web app performs MSAL sign-in itself.

After creation, record:

- Application (client) ID
- Directory (tenant) ID

## 5. App Roles

Create app roles in the App Registration and assign them to users/groups in the Enterprise Application.

Current role values accepted by the app:

| Role value | Intended use |
| --- | --- |
| `Admin` | Can view sites, review permissions, view audit logs, and grant/update/remove direct permissions. |
| `Reviewer` | Read-only reviewer. Can view reports and audit logs. |
| `InternalUser` | Read-only limited browsing role. Current behavior only shows standard/internal libraries. |
| `GuestUser` | Minimal access role. Current behavior signs in but does not expose content. |
| `ExecutiveUser` | Backward-compatible read-only role that behaves like `Reviewer`. |

If the customer wants only `Admin`, `Reviewer`, and `GuestUser`, reassign any users from `InternalUser`/`ExecutiveUser` before disabling or removing those roles.

## 6. Microsoft Graph API Permissions

Add these delegated Microsoft Graph permissions and grant admin consent.

| Permission | Type | Used for |
| --- | --- | --- |
| `User.Read` | Delegated | Basic sign-in and user profile. |
| `User.ReadBasic.All` | Delegated | People search and readable user display information. |
| `Sites.Read.All` | Delegated | Read SharePoint sites, libraries, files, and permissions. |
| `Sites.ReadWrite.All` | Delegated | Write SharePoint permission changes and audit/review list content. |
| `Files.ReadWrite.All` | Delegated | Drive item permission updates through Microsoft Graph. |

The app uses delegated permissions. The signed-in user's SharePoint and Graph rights still matter. A user assigned the app `Admin` role must also have sufficient SharePoint/Graph permission to perform the requested action.

## 7. SharePoint Preparation

### External File Sharing

For external recipients to open files from the invitation email or from a copied SharePoint link, the customer must confirm these tenant controls before UAT:

- SharePoint admin center organization sharing allows at least new and existing guests.
- Each target site allows external sharing at a level no more restrictive than the organization setting.
- Microsoft Entra External Identities B2B collaboration settings allow invitations for the recipient domain.
- Cross-tenant access and Conditional Access policies allow the guest to redeem the invitation and sign in to the resource tenant.
- If the file is in a protected library, Microsoft Purview sensitivity labels, encryption, or Rights Management also allow that external recipient to open the content.

The app grants item permissions through Microsoft Graph and sends the SharePoint invitation email. After a successful grant, the app shows the best sharing link returned by Graph. If the invite response does not include a link, the app creates a signed-in users sharing link for the same item before falling back to the direct SharePoint item URL. This avoids sending path-based URLs that can fail for external users who cannot traverse parent folders. If the recipient still sees "You need access", troubleshoot the tenant controls above before re-granting the same permission.

### Target Sites

The customer must identify all SharePoint sites the app should manage/review.

Required format:

```text
hostname:/sites/path
```

Example:

```text
contoso.sharepoint.com:/sites/Finance,contoso.sharepoint.com:/sites/Legal
```

### Protected Libraries

Protected libraries are identified by exact document library display name.

Example:

```text
Confidential,Secret
```

These names control app labeling and internal report counts. They do not create SharePoint sensitivity labels or Microsoft Purview labels by themselves.

### Audit List

Default list name:

```text
PermissionAuditLog
```

The app writes audit records for:

- Login
- Report refresh
- Grant access
- Update role
- Remove access
- Success/failure status
- Approved request number
- Actor, target, site/library, role, tenant type, error message, and Graph request ID

The app can create the audit list and missing text columns on first write if the signed-in Admin has sufficient permission.

Expected audit columns:

```text
Action
ActorEmail
ActorName
ActorRole
ApprovalRequestNo
TargetEmail
TargetName
PermissionRole
PreviousRole
SiteName
LibraryName
ItemId
Source
TenantType
Status
ErrorMessage
GraphRequestId
CreatedAt
```

### Managed Sites List

Default list name:

```text
ManagedSites
```

Create this list once on the audit/governance site. The portal uses it as the central registry for SharePoint sites added by Admin users.

Required columns:

```text
Hostname
Path
Active
```

Example row:

```text
Hostname: contoso.sharepoint.com
Path: /sites/Finance
Active: Yes
```

### Reviewer Scope List

Default list name:

```text
PermissionReviewScopes
```

Unlike the audit list, this list should be prepared by the customer or project team before reviewer testing.

Minimum column:

```text
OwnerEmail
```

Recommended columns:

```text
OwnerName
OwnerRole
SiteName
Hostname
Path
LibraryName
SensitivityLabel
Department
Section
Active
```

Example row:

```text
OwnerEmail: owner.rep@contoso.com
OwnerName: Owner Rep
OwnerRole: OwnerRep
SiteName: Finance
Hostname: contoso.sharepoint.com
Path: /sites/Finance
LibraryName: Confidential
SensitivityLabel: Confidential
Department: Finance
Section: Treasury
Active: Yes
```

When active scope rows exist, the Reviewer screen asks the user to select an owner before loading the report.

## 8. Environment Variables

Use the customer's tenant values in the deployment environment.

```bash
NEXT_PUBLIC_MSAL_CLIENT_ID=<customer-app-client-id>
NEXT_PUBLIC_MSAL_TENANT_ID=<customer-tenant-id>
NEXT_PUBLIC_APP_SESSION_MAX_MINUTES=480
NEXT_PUBLIC_TENANT_DOMAIN=contoso.com
NEXT_PUBLIC_INTERNAL_DOMAINS=contoso.com,contoso.onmicrosoft.com
NEXT_PUBLIC_TARGET_SITES=contoso.sharepoint.com:/sites/Finance,contoso.sharepoint.com:/sites/Legal
NEXT_PUBLIC_PROTECTED_LIBRARY_NAMES=Confidential,Secret
NEXT_PUBLIC_AUDIT_SITE=contoso.sharepoint.com:/sites/Governance
NEXT_PUBLIC_AUDIT_LIST_NAME=PermissionAuditLog
NEXT_PUBLIC_AUDIT_LOG_ENABLED=true
NEXT_PUBLIC_REVIEW_SCOPE_SITE=contoso.sharepoint.com:/sites/Governance
NEXT_PUBLIC_REVIEW_SCOPE_LIST_NAME=PermissionReviewScopes
NEXT_PUBLIC_REVIEW_SCOPE_LIST_ENABLED=true
NEXT_PUBLIC_REVIEW_SCAN_DESCENDANTS=true
NEXT_PUBLIC_REVIEW_SCAN_ITEM_LIMIT=2000
NEXT_PUBLIC_REVIEW_SCOPES=[]
```

Important: Next.js embeds `NEXT_PUBLIC_*` values during `next build`. If using GitHub Actions to build for Azure App Service, set these values as GitHub repository/environment variables before build, not only as Azure App Settings after deployment.

## 9. My Apps Tile

If the customer wants the app to appear in Microsoft My Apps:

1. Create or use an Enterprise Application tile.
2. Configure it as a linked application.
3. Set the Sign-on URL/Homepage URL to the deployed app URL.
4. Set `Visible to users?` to `Yes`.
5. Assign users/groups to the Enterprise Application.

The My Apps tile should open the web app URL. The tile should not perform OIDC sign-in through `launcher.myapps.microsoft.com`.

## 10. Customer UAT Checklist

Use this checklist after deployment.

| Test | Expected result |
| --- | --- |
| Admin signs in | Admin reaches the Admin workspace. |
| Reviewer signs in | Reviewer reaches the Reviewer report view and cannot change permissions. |
| Guest/minimal user signs in | User does not see protected management capabilities. |
| Target sites load | Configured SharePoint sites appear in the site picker. |
| Libraries load | Document libraries appear for the selected site. |
| Permissions load | Direct/inherited permissions are visible for selected items. |
| Grant permission | Admin can grant Viewer/Editor after entering an approved request number. |
| Update role | Admin can change Viewer/Editor where Graph allows it. |
| Remove access | Admin can remove direct permissions where Graph allows it. |
| Audit write | Permission actions create records in `PermissionAuditLog`. |
| Audit view | Audit screen loads recent records. |
| Reviewer scope | Owner dropdown loads from `PermissionReviewScopes`. |
| External classification | Non-internal domains appear as external in reports. |
| Session timeout | App requires sign-in again after `NEXT_PUBLIC_APP_SESSION_MAX_MINUTES`. |

## 11. Handover Notes

Customer operations team should own:

- Entra app role assignments.
- SharePoint site/library selection.
- Audit log retention policy.
- Reviewer scope list maintenance.
- SharePoint sharing and external collaboration policy.
- Approval workflow outside this web app.

The web app records an approved request number for traceability, but it does not implement the customer's approval workflow.
