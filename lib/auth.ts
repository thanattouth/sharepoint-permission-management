import {
  InteractionRequiredAuthError,
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
  type Configuration,
  type PopupRequest,
} from "@azure/msal-browser";
import { graphReadScopes } from "./graph";

const clientId = process.env.NEXT_PUBLIC_MSAL_CLIENT_ID;
const tenantId = process.env.NEXT_PUBLIC_MSAL_TENANT_ID ?? "common";

export const isAuthConfigured = Boolean(clientId);

export const msalConfig: Configuration = {
  auth: {
    clientId: clientId ?? "",
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri: typeof window === "undefined" ? undefined : window.location.origin,
  },
  cache: {
    cacheLocation: "sessionStorage",
  },
};

export const loginRequest: PopupRequest = {
  scopes: graphReadScopes,
};

let msalInstance: PublicClientApplication | undefined;

export function getMsalInstance() {
  msalInstance ??= new PublicClientApplication(msalConfig);
  return msalInstance;
}

export async function signInMicrosoft365(): Promise<AuthenticationResult> {
  const msal = getMsalInstance();
  await msal.initialize();
  const response = await msal.loginPopup(loginRequest);
  msal.setActiveAccount(response.account);
  return response;
}

export async function getSignedInAccount() {
  const msal = getMsalInstance();
  await msal.initialize();
  const account = msal.getActiveAccount() ?? msal.getAllAccounts()[0] ?? null;
  if (account) {
    msal.setActiveAccount(account);
  }
  return account;
}

export async function signOutMicrosoft365(account?: AccountInfo | null) {
  const msal = getMsalInstance();
  await msal.initialize();
  const activeAccount = account ?? msal.getActiveAccount() ?? msal.getAllAccounts()[0] ?? undefined;

  await msal.logoutPopup({
    account: activeAccount,
    mainWindowRedirectUri: typeof window === "undefined" ? undefined : window.location.origin,
  });
}

export async function acquireGraphToken(account?: AccountInfo | null, scopes = graphReadScopes) {
  const msal = getMsalInstance();
  await msal.initialize();
  const activeAccount = account ?? msal.getActiveAccount() ?? msal.getAllAccounts()[0];

  if (!activeAccount) {
    const response = await signInMicrosoft365();
    return response.accessToken;
  }

  try {
    const response = await msal.acquireTokenSilent({
      scopes,
      account: activeAccount,
    });
    return response.accessToken;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      const response = await msal.acquireTokenPopup({
        scopes,
        account: activeAccount,
      });
      return response.accessToken;
    }
    throw error;
  }
}
