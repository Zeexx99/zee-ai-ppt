import { STORAGE_PREFIX, ZeeSettings } from "../types/settings";

export const ALLOWED_PROXY_ACTIONS = [
  "generateContent",
  "streamGenerateContent",
  "countTokens",
] as const;

export type AllowedGeminiAction = (typeof ALLOWED_PROXY_ACTIONS)[number];

/**
 * Always builds a relative URL from the current page location so the exact
 * same build works on both https://localhost:3300/ and https://zeexx99.github.io/zee-ai-ppt/
 */
export function buildRelativeAuthUrl(
  relativePath: string = "auth/start.html",
  params: Record<string, string> = {},
  baseHref: string = typeof window !== "undefined" ? window.location.href : "https://localhost:3300/index.html"
): string {
  const normalizedBase = baseHref.endsWith("/")
    ? baseHref
    : baseHref.replace(/\/[^/]*$/, "/");
  const cleanRel = relativePath.replace(/^\/+/, "");
  const url = new URL(cleanRel, normalizedBase);
  for (const [k, v] of Object.entries(params)) {
    if (v) {
      url.searchParams.set(k, v);
    }
  }
  return url.toString();
}

export function isAuthConfigured(settings: ZeeSettings): boolean {
  if (settings.authMode === "proxy") {
    return Boolean(settings.proxyUrl.trim() && settings.proxyToken.trim());
  }
  if (settings.authMode === "oauth") {
    return Boolean(
      settings.oauthAccessToken.trim() &&
        (!settings.oauthExpiry || settings.oauthExpiry > Date.now() - 60_000)
    );
  }
  return Boolean(settings.geminiApiKey.trim());
}

export function buildGeminiEndpoint(
  settings: ZeeSettings,
  action: AllowedGeminiAction,
  streamSse: boolean = false
): {
  url: string;
  headers: Record<string, string>;
} {
  if (!ALLOWED_PROXY_ACTIONS.includes(action)) {
    throw new Error(`Unsupported Gemini API action: ${action}`);
  }

  const model = (settings.model || "gemini-2.5-flash").replace(/^models\//, "");
  const headers: Record<string, string> = {
    "Content-Type": "application/json",
  };

  if (settings.authMode === "proxy") {
    const baseProxy = settings.proxyUrl.trim().replace(/\/+$/, "");
    if (!baseProxy) {
      throw new Error("Team proxy URL is missing in Settings.");
    }
    const token = settings.proxyToken.trim();
    if (!token) {
      throw new Error("Team proxy token is missing in Settings.");
    }
    headers["Authorization"] = `Bearer ${token}`;
    headers["X-Zee-Token"] = token;
    headers["X-Zee-Addin"] = "zee-ai-ppt";
    const query = streamSse ? "?alt=sse" : "";
    return {
      url: `${baseProxy}/v1beta/models/${model}:${action}${query}`,
      headers,
    };
  }

  if (settings.authMode === "oauth") {
    const token = settings.oauthAccessToken.trim();
    if (!token) {
      throw new Error("Please sign in with Google in Settings or switch to Gemini API Key.");
    }
    headers["Authorization"] = `Bearer ${token}`;
    const query = streamSse ? "?alt=sse" : "";
    return {
      url: `https://generativelanguage.googleapis.com/v1beta/models/${model}:${action}${query}`,
      headers,
    };
  }

  const apiKey = settings.geminiApiKey.trim();
  if (!apiKey) {
    throw new Error("Gemini API Key is not set. Open Settings (gear icon) to enter your key.");
  }
  headers["x-goog-api-key"] = apiKey;
  const query = streamSse ? "?alt=sse" : "";
  return {
    url: `https://generativelanguage.googleapis.com/v1beta/models/${model}:${action}${query}`,
    headers,
  };
}

export interface OAuthSignInResult {
  accessToken: string;
  expiresIn: number;
  email?: string;
}

export async function startGoogleOAuthSignIn(
  clientId: string
): Promise<OAuthSignInResult> {
  if (!clientId.trim()) {
    throw new Error("Please enter a Google OAuth Client ID in Settings first.");
  }

  const state = Math.random().toString(36).slice(2);
  const startUrl = buildRelativeAuthUrl("auth/start.html", {
    client_id: clientId.trim(),
    state,
  });

  localStorage.removeItem(`${STORAGE_PREFIX}oauth-callback`);

  return new Promise<OAuthSignInResult>((resolve, reject) => {
    let settled = false;

    const finish = (rawPayload: string) => {
      if (settled) return;
      try {
        const parsed = JSON.parse(rawPayload);
        if (parsed.error) {
          settled = true;
          reject(new Error(`Google Sign-In error: ${parsed.error}`));
          return;
        }
        if (parsed.accessToken) {
          settled = true;
          resolve({
            accessToken: parsed.accessToken,
            expiresIn: Number(parsed.expiresIn) || 3600,
          });
        }
      } catch (err) {
        settled = true;
        reject(err instanceof Error ? err : new Error(String(err)));
      }
    };

    const storagePoller = window.setInterval(() => {
      const saved = localStorage.getItem(`${STORAGE_PREFIX}oauth-callback`);
      if (saved) {
        window.clearInterval(storagePoller);
        localStorage.removeItem(`${STORAGE_PREFIX}oauth-callback`);
        finish(saved);
      }
    }, 600);

    const onWindowMessage = (ev: MessageEvent) => {
      if (typeof ev.data === "string" && ev.data.includes("zee-ppt-oauth")) {
        window.clearInterval(storagePoller);
        window.removeEventListener("message", onWindowMessage);
        finish(ev.data);
      }
    };
    window.addEventListener("message", onWindowMessage);

    try {
      if (
        typeof Office !== "undefined" &&
        Office.context &&
        Office.context.ui &&
        typeof Office.context.ui.displayDialogAsync === "function"
      ) {
        Office.context.ui.displayDialogAsync(
          startUrl,
          { height: 65, width: 45, displayInIframe: false },
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              window.clearInterval(storagePoller);
              window.removeEventListener("message", onWindowMessage);
              reject(new Error(asyncResult.error?.message || "Failed to open sign-in dialog."));
              return;
            }
            const dialog = asyncResult.value;
            dialog.addEventHandler(
              Office.EventType.DialogMessageReceived,
              (arg: { message: string } | { error: number }) => {
                window.clearInterval(storagePoller);
                window.removeEventListener("message", onWindowMessage);
                dialog.close();
                if ("message" in arg) {
                  finish(arg.message);
                }
              }
            );
            dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
              window.clearInterval(storagePoller);
              window.removeEventListener("message", onWindowMessage);
              if (!settled) {
                reject(new Error("Sign-in window was closed."));
              }
            });
          }
        );
        return;
      }
    } catch {
      // Fall back to window.open below
    }

    const popup = window.open(startUrl, "zee_ppt_oauth", "width=500,height=640");
    if (!popup) {
      window.clearInterval(storagePoller);
      window.removeEventListener("message", onWindowMessage);
      reject(new Error("Popup blocked. Please allow popups for Google Sign-In."));
    }
  });
}
