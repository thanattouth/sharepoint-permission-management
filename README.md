# SharePoint Permission Management

Production-style SharePoint permission console for SharePoint Online permission governance.

The web app scope is intentionally limited to three role-based screens:

1. Admin permission management with required approved request number and audit trail
2. Reviewer read-only permission review with Excel-like summary
3. Audit trail view

Approval routing is separate from the web app. Power Apps and Power Automate can be delivered as a customer-customizable sample flow that returns an approved request number; admins enter that number before changing SharePoint permissions.

The app connects directly to Microsoft Graph. Microsoft Entra configuration is required before sign-in.

## Architecture

Target architecture diagrams are maintained as diagram-as-code in [docs/architecture.md](docs/architecture.md).

Standalone Mermaid sources are available in [docs/diagrams](docs/diagrams):

- `system-context.mmd`
- `application-screens.mmd`
- `permission-change-flow.mmd`
- `audit-data-model.mmd`

Backend-facing application code is organized by feature under `lib/features`:

- `lib/features/admin`: SharePoint inventory browsing, permission changes, permission labels, and SharePoint principal normalization.
- `lib/features/reviewer`: owner scope parsing/loading plus permission review report and deep scan logic.
- `lib/features/audit`: audit store interfaces and SharePoint List persistence.

Shared cross-feature code stays in `lib`, such as authentication, app configuration, Microsoft Graph request handling, roles, and shared data types. Feature-specific code should import from `lib/features/...`.

## Run

```bash
npm install
npm run dev
```

Open [http://localhost:3000](http://localhost:3000).

## Microsoft Entra Configuration

Create `.env.local` or `.env` when the App Registration is ready:

```bash
NEXT_PUBLIC_MSAL_CLIENT_ID=your-client-id
NEXT_PUBLIC_MSAL_TENANT_ID=your-tenant-id
NEXT_PUBLIC_TENANT_DOMAIN=baht.net
NEXT_PUBLIC_INTERNAL_DOMAINS=baht.net,bahtnet.onmicrosoft.com
NEXT_PUBLIC_TARGET_SITES=bahtnet.sharepoint.com:/sites/DGCS,bahtnet.sharepoint.com:/sites/EngineerSite
NEXT_PUBLIC_PROTECTED_LIBRARY_NAMES=Confidential,Secret
NEXT_PUBLIC_AUDIT_SITE=bahtnet.sharepoint.com:/sites/DGCS
NEXT_PUBLIC_AUDIT_LIST_NAME=PermissionAuditLog
NEXT_PUBLIC_AUDIT_LOG_ENABLED=true
NEXT_PUBLIC_REVIEW_SCOPE_SITE=bahtnet.sharepoint.com:/sites/DGCS
NEXT_PUBLIC_REVIEW_SCOPE_LIST_NAME=PermissionReviewScopes
NEXT_PUBLIC_REVIEW_SCOPE_LIST_ENABLED=true
NEXT_PUBLIC_REVIEW_SCAN_DESCENDANTS=true
NEXT_PUBLIC_REVIEW_SCAN_ITEM_LIMIT=2000
NEXT_PUBLIC_REVIEW_SCOPES=[{"ownerEmail":"thanattouth@baht.net","ownerName":"Thanattouth Niticharoen","ownerRole":"OwnerRep","siteName":"DGCS","hostname":"bahtnet.sharepoint.com","path":"/sites/DGCS","libraryName":"Confidential","sensitivityLabel":"Confidential"}]
```

Initial Graph scopes are declared in `lib/features/admin/sharepoint-permission-client.ts`:

- `User.Read`
- `User.ReadBasic.All`
- `Sites.Read.All`
- `Sites.ReadWrite.All`
- `Files.ReadWrite.All`

`NEXT_PUBLIC_TARGET_SITES` is a comma-separated list of `hostname:/sites/path` entries. `NEXT_PUBLIC_INTERNAL_DOMAINS` controls which email domains are treated as internal in reports. `NEXT_PUBLIC_PROTECTED_LIBRARY_NAMES` controls which document libraries are labeled as protected by the permission console. Reviewer scope mappings are loaded from the `PermissionReviewScopes` SharePoint List configured by `NEXT_PUBLIC_REVIEW_SCOPE_SITE` and `NEXT_PUBLIC_REVIEW_SCOPE_LIST_NAME`; `NEXT_PUBLIC_REVIEW_SCOPES` remains as an optional local fallback/demo JSON array when the list is empty or unavailable. For the role model, Admin can change permissions, Reviewer/ExecutiveUser can review configured sites and audit entries, and Internal User can browse configured sites but only sees standard/internal libraries.

The configured sites are loaded from Microsoft Graph after sign-in, followed by drives, drive items, and permissions for the selected workspace.

## Audit Log

Permission changes are written to a SharePoint List so governance evidence stays in Microsoft 365. By default, the app uses the first configured site, or `NEXT_PUBLIC_AUDIT_SITE` when set, and writes to `PermissionAuditLog`.

The app creates the list automatically on first write when the signed-in Admin has enough SharePoint/Graph permission. Logged events include login, report refresh, grant access, role update, remove access, success/failure status, actor, target, site/library, role, approved request number, tenant type, error message, and Graph request id when available.

Set `NEXT_PUBLIC_AUDIT_LOG_ENABLED=false` to disable SharePoint audit writes without removing the local recent-changes UI.

For Azure App Service deployments, add the same `NEXT_PUBLIC_*` values as GitHub repository variables too. Next.js embeds `NEXT_PUBLIC_*` values during `next build`, so Azure App Settings alone are not enough when GitHub Actions builds the artifact.

Audit persistence is isolated behind `AuditStore` in `lib/audit-store.ts`. The current implementation is `SharePointListAuditStore`; a future database-backed store can be added without changing permission or UI logic.

For site-scoped permission actions, audit entries are written to the `PermissionAuditLog` list on the selected SharePoint site. App-wide events such as login and report refresh use `NEXT_PUBLIC_AUDIT_SITE` as the default audit location.

## Reviewer Scope List

Reviewer owner mappings are read from the `PermissionReviewScopes` SharePoint List. The minimum required column is `OwnerEmail`; recommended columns are `OwnerName`, `OwnerRole`, `SiteName`, `Hostname`, `Path`, `LibraryName`, `SensitivityLabel`, `Department`, `Section`, and `Active`.

Example test item:

```text
OwnerEmail: .....@baht.net
OwnerName: firstname lastname
OwnerRole: OwnerRep
SiteName: DGCS
Hostname: bahtnet.sharepoint.com
Path: /sites/DGCS
LibraryName: Confidential
SensitivityLabel: Confidential
Active: Yes
```

When active scope mappings exist, the Reviewer screen asks the user to select an owner before loading a report. This prevents the app from scanning every configured site by default.

With `NEXT_PUBLIC_REVIEW_SCAN_DESCENDANTS=true`, an owner-scoped review scans the full configured library hierarchy, including library root, folders, and files. `NEXT_PUBLIC_REVIEW_SCAN_ITEM_LIMIT` caps how many items a single refresh will traverse to protect large enterprise libraries from long-running browser scans.

## Graph Integration

When the Entra settings are present, the app connects to Microsoft Graph after sign-in and loads:

- Sites from the configured SharePoint paths in `lib/app-config.ts`
- Document libraries from `GET /sites/{siteId}/drives`
- Library root permissions from `GET /drives/{driveId}/items/{itemId}/permissions`

Permission actions use Graph where possible:

- Grant Viewer/Editor: `driveItem: invite`
- Change Viewer/Editor: `PATCH permission.roles`
- Remove direct permission: `DELETE permission`

Inherited permissions and some sharing links cannot be changed directly from the selected library root; the UI disables those actions when Graph returns that metadata.
