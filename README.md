# SharePoint Permission Management

Production-style SharePoint permission console focused on the first five MVP flows:

1. Microsoft Entra ID login
2. Site selection
3. Document Library selection
4. Member/permission visibility
5. Viewer/Editor permission management

The app connects directly to Microsoft Graph. Microsoft Entra configuration is required before sign-in.

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
```

Initial Graph scopes are declared in `lib/graph.ts`:

- `User.Read`
- `User.ReadBasic.All`
- `Sites.Read.All`
- `Sites.ReadWrite.All`
- `Files.ReadWrite.All`

`NEXT_PUBLIC_TARGET_SITES` is a comma-separated list of `hostname:/sites/path` entries. `NEXT_PUBLIC_INTERNAL_DOMAINS` controls which email domains are treated as internal in reports. `NEXT_PUBLIC_PROTECTED_LIBRARY_NAMES` controls which document libraries are labeled as protected by the permission console. For the demo role model, Internal User can browse configured sites but only sees standard/internal libraries.

The configured sites are loaded from Microsoft Graph after sign-in, followed by drives, drive items, and permissions for the selected workspace.

## Graph Integration

When the Entra settings are present, the app connects to Microsoft Graph after sign-in and loads:

- Sites from the configured SharePoint paths in `lib/graph.ts`
- Document libraries from `GET /sites/{siteId}/drives`
- Library root permissions from `GET /drives/{driveId}/items/{itemId}/permissions`

Permission actions use Graph where possible:

- Grant Viewer/Editor: `driveItem: invite`
- Change Viewer/Editor: `PATCH permission.roles`
- Remove direct permission: `DELETE permission`

Inherited permissions and some sharing links cannot be changed directly from the selected library root; the UI disables those actions when Graph returns that metadata.
