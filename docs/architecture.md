# SharePoint Permission Management Architecture

This document describes the target architecture for the SharePoint Permission Management web app after separating approval workflow ownership from the core web app.

## Scope

The web app is responsible for SharePoint permission review, direct permission changes by authorized admins, and audit trail capture.

Approval workflow is intentionally separate. Power Apps and Power Automate provide a sample approval flow that the customer can customize. The web app records the approved request number when an admin changes permissions, so auditors can reference the external workflow.

## System Context

```mermaid
flowchart LR
  Admin["Admin"]
  Reviewer["Reviewer"]
  Approver["Owner / VP / EVP Approver"]

  WebApp["SPO Permission Management Web App<br/>Azure App Service / Next.js"]
  Entra["Microsoft Entra ID<br/>Authentication + App Roles"]
  Graph["Microsoft Graph"]
  SharePoint["SharePoint Online<br/>Sites / Libraries / Folders / Files"]
  AuditList["SharePoint List<br/>PermissionAuditLog"]
  PowerPlatform["Power Apps + Power Automate<br/>Customer-customizable approval sample"]

  Admin -->|"Manage permissions<br/>with approved request no."| WebApp
  Reviewer -->|"Review permissions<br/>read-only"| WebApp
  WebApp -->|"Sign-in + app roles"| Entra
  WebApp -->|"Read hierarchy + permissions<br/>Apply admin changes"| Graph
  Graph --> SharePoint
  WebApp -->|"Write audit trail"| AuditList

  Admin -.->|"Request approval outside app"| PowerPlatform
  Approver -.->|"Approve / reject"| PowerPlatform
  PowerPlatform -.->|"Returns approval request no."| Admin
```

## Application Screens

```mermaid
flowchart TD
  Login["Microsoft Entra sign-in"]
  RoleCheck{"App role"}

  AdminHome["Admin console"]
  ReviewerHome["Reviewer console"]
  Audit["Audit list"]

  Manage["Permission management<br/>Grant / update / remove"]
  Review["Permission review<br/>Site hierarchy, library, folder, file, roles"]
  Reports["Summary report<br/>Excel-like review view"]

  Login --> RoleCheck
  RoleCheck -->|"Admin"| AdminHome
  RoleCheck -->|"Reviewer / Executive read-only"| ReviewerHome
  RoleCheck -->|"No role"| Denied["Access denied"]

  AdminHome --> Manage
  AdminHome --> Review
  AdminHome --> Reports
  AdminHome --> Audit

  ReviewerHome --> Review
  ReviewerHome --> Reports
  ReviewerHome --> Audit
```

## Permission Change Flow

```mermaid
sequenceDiagram
  autonumber
  actor Admin
  participant Approval as Power App / Power Automate sample
  participant App as Web App
  participant Graph as Microsoft Graph
  participant SPO as SharePoint Online
  participant Audit as PermissionAuditLog

  Admin->>Approval: Submit access approval request
  Approval->>Approval: Route to owner / VP / EVP approver
  Approval-->>Admin: Approved request number
  Admin->>App: Enter approved request number
  Admin->>App: Grant / update / remove permission
  App->>Graph: Apply permission change
  Graph->>SPO: Update SharePoint permission
  Graph-->>App: Result
  App->>Audit: Write audit trail with approved request number
  App-->>Admin: Show success or failure
```

## Reviewer Flow

```mermaid
flowchart TD
  Reviewer["Reviewer"]
  App["Web App"]
  Graph["Microsoft Graph"]
  SPO["SharePoint Online"]
  View["Reviewer View<br/>Hierarchy + permissions"]

  Reviewer -->|"Open review screen"| App
  App -->|"Read sites, libraries, folders, files, permissions"| Graph
  Graph --> SPO
  Graph -->|"Permission inventory"| App
  App --> View
```

## Audit Data Model

```mermaid
erDiagram
  PERMISSION_AUDIT_LOG {
    string Title
    string Action
    string ActorEmail
    string ActorName
    string ActorRole
    string ApprovalRequestNo
    string TargetEmail
    string TargetName
    string PermissionRole
    string PreviousRole
    string SiteName
    string LibraryName
    string ItemId
    string Source
    string TenantType
    string Status
    string ErrorMessage
    string GraphRequestId
    datetime CreatedAt
  }
```

## Role Responsibilities

| Role | Can change permissions | Can review permission inventory | Can view audit | Notes |
| --- | --- | --- | --- | --- |
| Admin | Yes | Yes | Yes | Must enter approved request number before permission changes. |
| Reviewer | No | Yes | Yes | Read-only review and audit access. |
| ExecutiveUser | No | Yes | Yes | Backward-compatible read-only role that maps to Reviewer behavior. |

## Implementation Impact

The web app should remove core dependencies on approval request processing:

- Remove internal `PermissionAccessRequests` workflow from the main app path.
- Remove backend apply endpoint that exists only for Power Automate HTTP callbacks.
- Keep direct Graph permission changes for Admin, guarded by approved request number input.
- Extend `PermissionAuditLog` with `ApprovalRequestNo`.
- Add a dedicated Audit screen.
- Evolve the Reports screen into a Reviewer-oriented permission inventory.

Power Apps and Power Automate can remain as a separate sample package:

- Requester submits approval.
- Owner / VP / EVP approves.
- The flow returns or records an approved request number.
- Admin references that number in the web app before applying SharePoint permission changes.
