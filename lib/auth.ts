import {
  InteractionRequiredAuthError,
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
  type Configuration,
  type RedirectRequest,
} from "@azure/msal-browser";
import { graphReadScopes } from "./features/admin";

const clientId = process.env.NEXT_PUBLIC_MSAL_CLIENT_ID;
const tenantId = process.env.NEXT_PUBLIC_MSAL_TENANT_ID ?? "common";
const defaultAppSessionMaxMinutes = 480;

export const isAuthConfigured = Boolean(clientId);
export const appSessionMaxAgeMs = parsePositiveInteger(
  process.env.NEXT_PUBLIC_APP_SESSION_MAX_MINUTES,
  defaultAppSessionMaxMinutes,
) * 60 * 1000;

export const msalConfig: Configuration = {
  auth: {
    clientId: clientId ?? "",
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri: typeof window === "undefined" ? undefined : window.location.origin,
    navigateToLoginRequestUrl: false,
  },
  cache: {
    cacheLocation: "sessionStorage",
  },
};

export const loginRequest: RedirectRequest = {
  scopes: ["User.Read"],
  prompt: "select_account",
};

let msalInstance: PublicClientApplication | undefined;
let interactiveRequest: Promise<unknown> | undefined;
let redirectResponse: Promise<AuthenticationResult | null> | undefined;

export function getMsalInstance() {
  msalInstance ??= new PublicClientApplication(msalConfig);
  return msalInstance;
}

async function initializeMsal() {
  const msal = getMsalInstance();
  await msal.initialize();
  redirectResponse ??= msal.handleRedirectPromise();
  const response = await redirectResponse;

  if (response?.account) {
    msal.setActiveAccount(response.account);
  }

  return msal;
}

function isInteractionInProgressError(error: unknown) {
  if (!(error instanceof Error)) return false;
  return error.message.includes("interaction_in_progress") || error.name === "BrowserAuthError";
}

function clearStaleInteractionState() {
  if (typeof window === "undefined") return;

  const storages = [window.sessionStorage, window.localStorage];
  for (const storage of storages) {
    for (let index = storage.length - 1; index >= 0; index -= 1) {
      const key = storage.key(index);
      if (!key) continue;

      const normalizedKey = key.toLowerCase();
      const isMsalTemporaryKey =
        normalizedKey.includes("msal") &&
        (normalizedKey.includes("interaction.status") ||
          normalizedKey.includes("interaction_in_progress") ||
          normalizedKey.includes("request.state") ||
          normalizedKey.includes("request.params") ||
          normalizedKey.includes("urlhash"));

      if (isMsalTemporaryKey) {
        storage.removeItem(key);
      }
    }
  }
}

async function runInteractiveRequest<T>(request: () => Promise<T>) {
  if (interactiveRequest) {
    await interactiveRequest.catch(() => undefined);
  }

  const nextRequest = request();
  interactiveRequest = nextRequest;

  try {
    return await nextRequest;
  } finally {
    if (interactiveRequest === nextRequest) {
      interactiveRequest = undefined;
    }
  }
}

async function runInteractiveRequestWithRecovery<T>(request: () => Promise<T>) {
  try {
    return await runInteractiveRequest(request);
  } catch (error) {
    if (!isInteractionInProgressError(error)) throw error;
    clearStaleInteractionState();
    return runInteractiveRequest(request);
  }
}

export async function signInMicrosoft365(): Promise<AuthenticationResult | null> {
  const msal = await initializeMsal();
  const existingAccount = msal.getActiveAccount() ?? msal.getAllAccounts()[0] ?? null;

  if (existingAccount) {
    msal.setActiveAccount(existingAccount);
    const response = await msal.acquireTokenSilent({
      scopes: loginRequest.scopes,
      account: existingAccount,
    });
    return response;
  }

  await runInteractiveRequestWithRecovery(() =>
    msal.loginRedirect({
      ...loginRequest,
      redirectStartPage: window.location.href,
    }),
  );
  return null;
}

export async function getSignedInAccount() {
  const msal = await initializeMsal();
  const account = msal.getActiveAccount() ?? msal.getAllAccounts()[0] ?? null;
  if (account) {
    msal.setActiveAccount(account);
  }
  return account;
}

export async function signOutMicrosoft365(account?: AccountInfo | null) {
  const msal = await initializeMsal();
  const activeAccount = account ?? msal.getActiveAccount() ?? msal.getAllAccounts()[0] ?? undefined;

  await msal.clearCache({ account: activeAccount });
  clearStaleInteractionState();
}

function parsePositiveInteger(value: string | undefined, fallback: number) {
  const parsed = Number.parseInt(value ?? "", 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

export async function acquireGraphToken(
  account?: AccountInfo | null,
  scopes = graphReadScopes,
  options: { allowPopup?: boolean } = {},
) {
  const allowPopup = options.allowPopup ?? true;
  const msal = await initializeMsal();
  const activeAccount = account ?? msal.getActiveAccount() ?? msal.getAllAccounts()[0];

  if (!activeAccount) {
    const response = await signInMicrosoft365();
    if (!response) {
      throw new Error("Microsoft 365 sign-in is required before loading SharePoint data.");
    }
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
      if (!allowPopup) {
        throw new Error("Additional Microsoft Graph consent is required before this data can be loaded.");
      }

      const response = await runInteractiveRequestWithRecovery(() => msal.acquireTokenPopup({
        scopes,
        account: activeAccount,
      }));
      return response.accessToken;
    }
    throw error;
  }
}
