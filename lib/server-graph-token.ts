import type { TokenProvider } from "./graph-request";

type TokenResponse = {
  access_token?: string;
  expires_in?: number;
  error?: string;
  error_description?: string;
};

let cachedToken: {
  accessToken: string;
  expiresAt: number;
} | undefined;

export function createClientCredentialTokenProvider(): TokenProvider {
  return async () => {
    const now = Date.now();
    if (cachedToken && cachedToken.expiresAt - 60_000 > now) {
      return cachedToken.accessToken;
    }

    const clientId = process.env.GRAPH_CLIENT_ID ?? process.env.NEXT_PUBLIC_MSAL_CLIENT_ID;
    const tenantId = process.env.GRAPH_TENANT_ID ?? process.env.NEXT_PUBLIC_MSAL_TENANT_ID;
    const clientSecret = process.env.GRAPH_CLIENT_SECRET;

    if (!clientId || !tenantId || !clientSecret) {
      throw new Error("Server Graph credentials are missing. Set GRAPH_CLIENT_ID, GRAPH_TENANT_ID, and GRAPH_CLIENT_SECRET.");
    }

    const body = new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      grant_type: "client_credentials",
      scope: "https://graph.microsoft.com/.default",
    });

    const response = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body,
    });

    const token = (await response.json()) as TokenResponse;
    if (!response.ok || !token.access_token) {
      throw new Error(token.error_description || token.error || "Unable to acquire Microsoft Graph application token.");
    }

    cachedToken = {
      accessToken: token.access_token,
      expiresAt: now + (token.expires_in ?? 3600) * 1000,
    };

    return cachedToken.accessToken;
  };
}
