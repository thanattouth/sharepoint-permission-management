# SharePoint Permission Management

Production-style SharePoint permission console focused on the first five MVP flows:

1. Microsoft Entra ID login
2. Site selection
3. Document Library selection
4. Member/permission visibility
5. Viewer/Editor permission management

The app currently ships with demo data for the DGCS and EngineerSite test sites, plus an MSAL-ready login path. If Microsoft Entra environment variables are not configured, the Connect button uses demo mode so the UI can still be tested locally.

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
```

Initial Graph scopes are declared in `lib/graph.ts`:

- `User.Read`
- `Sites.Read.All`
- `Sites.ReadWrite.All`
- `Files.ReadWrite.All`

The UI is wired through a mock client boundary so the next step is replacing demo data calls with Microsoft Graph API calls for sites, drives, drive items, and permissions.

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
