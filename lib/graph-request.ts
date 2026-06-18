export type TokenProvider = () => Promise<string>;

export type GraphCollection<T> = {
  value?: T[];
  "@odata.nextLink"?: string;
};

export class GraphRequestClient {
  constructor(private readonly getToken: TokenProvider) {}

  async request<T>(path: string, init: RequestInit = {}) {
    const token = await this.getToken();
    const url = path.startsWith("https://graph.microsoft.com/")
      ? path
      : `https://graph.microsoft.com/v1.0${path}`;
    const response = await fetch(url, {
      ...init,
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        ...init.headers,
      },
    });

    if (!response.ok) {
      const body = await response.text();
      throw new Error(formatGraphError(response.status, body || response.statusText));
    }

    if (response.status === 204) {
      return undefined as T;
    }

    return (await response.json()) as T;
  }
}

function formatGraphError(status: number, body: string) {
  try {
    const parsed = JSON.parse(body) as {
      error?: {
        code?: string;
        message?: string;
        innerError?: {
          "request-id"?: string;
          date?: string;
        };
      };
    };
    const code = parsed.error?.code;
    const message = parsed.error?.message;
    const requestId = parsed.error?.innerError?.["request-id"];

    if (code === "sharingFailed") {
      return [
        "Graph sharing failed. SharePoint rejected this invitation.",
        message ? `Graph message: ${message}` : "",
        "Check SharePoint organization sharing, site sharing, domain restrictions, Microsoft Entra external collaboration settings, and whether this email domain is allowed.",
        requestId ? `Request ID: ${requestId}` : "",
      ]
        .filter(Boolean)
        .join(" ");
    }

    return [`Graph ${status}${code ? ` ${code}` : ""}: ${message || body}`, requestId ? `Request ID: ${requestId}` : ""]
      .filter(Boolean)
      .join(" ");
  } catch {
    return `Graph ${status}: ${body}`;
  }
}
