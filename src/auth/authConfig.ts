/**
 * MSAL configuration and token acquisition helpers.
 *
 * Supports two strategies:
 *   1. Nested App Authentication (NAA) — preferred when the host supports it.
 *      Uses createNestablePublicClientApplication for seamless SSO with the
 *      signed-in Microsoft Account.
 *   2. Office Dialog API fallback — opens a popup dialog that redirects to
 *      Authentik's OAuth2 authorize endpoint directly.
 *
 * After obtaining a Microsoft token (or directly authenticating via Authentik),
 * the token is exchanged with Authentik's token endpoint (server-side or
 * client-side) to obtain an API access token.
 */

import {
  type IPublicClientApplication,
  type AuthenticationResult,
  InteractionRequiredAuthError,
  createNestablePublicClientApplication,
} from "@azure/msal-browser";
import { getConfig } from "../shared/config";
import { setAuthenticated, clearAuth } from "../shared/state";

let msalInstance: IPublicClientApplication | null = null;
let naaSupported = false;

export async function initAuth(): Promise<void> {
  const config = getConfig();

  try {
    naaSupported = Office.context.requirements.isSetSupported("NestedAppAuth", "1.1");
  } catch {
    naaSupported = false;
  }

  if (naaSupported) {
    try {
      msalInstance = await createNestablePublicClientApplication({
        auth: {
          clientId: config.msalClientId,
          authority: config.msalAuthority,
        },
        cache: {
          cacheLocation: "localStorage",
        },
      });
    } catch (err) {
      console.warn("NAA initialization failed, will use Dialog API fallback.", err);
      naaSupported = false;
    }
  }
}

export function isNaaSupported(): boolean {
  return naaSupported;
}

/**
 * Acquire a Microsoft identity token via NAA (silent → popup fallback).
 */
export async function acquireMsToken(): Promise<AuthenticationResult> {
  if (!msalInstance) {
    throw new Error("MSAL not initialized. Call initAuth() first.");
  }

  const config = getConfig();
  const tokenRequest = { scopes: config.msalScopes };

  try {
    return await msalInstance.acquireTokenSilent(tokenRequest);
  } catch (err) {
    if (err instanceof InteractionRequiredAuthError) {
      return await msalInstance.acquireTokenPopup(tokenRequest);
    }
    throw err;
  }
}

/**
 * Exchange the Microsoft access token with Authentik to get an API token.
 *
 * This performs an OAuth2 token exchange: the MS token is sent to Authentik's
 * token endpoint. Authentik validates the Microsoft identity against its user
 * base and returns an access token scoped to the API.
 *
 * In production, this exchange should happen server-side for security.
 * This skeleton provides the client-side flow for demonstration.
 */
export async function exchangeTokenWithAuthentik(msAccessToken: string): Promise<{
  access_token: string;
  refresh_token?: string;
  expires_in?: number;
}> {
  const config = getConfig();

  const body = new URLSearchParams({
    grant_type: "urn:ietf:params:oauth:grant-type:token-exchange",
    client_id: config.authentikClientId,
    subject_token: msAccessToken,
    subject_token_type: "urn:ietf:params:oauth:token-type:access_token",
    scope: config.authentikScopes,
  });

  const response = await fetch(config.authentikTokenEndpoint, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });

  if (!response.ok) {
    const errBody = await response.text();
    throw new Error(`Authentik token exchange failed (${response.status}): ${errBody}`);
  }

  return response.json();
}

/**
 * Full sign-in flow: acquire MS token → exchange with Authentik → update shared state.
 */
export async function signIn(): Promise<void> {
  if (naaSupported && msalInstance) {
    const msResult = await acquireMsToken();
    const authentikResult = await exchangeTokenWithAuthentik(msResult.accessToken);

    setAuthenticated({
      msAccessToken: msResult.accessToken,
      authentikAccessToken: authentikResult.access_token,
      authentikRefreshToken: authentikResult.refresh_token,
      tokenExpiry: authentikResult.expires_in
        ? Date.now() + authentikResult.expires_in * 1000
        : undefined,
      userDisplayName: msResult.account?.name ?? undefined,
      userEmail: msResult.account?.username ?? undefined,
    });
  } else {
    await signInViaDialog();
  }
}

/**
 * Sign in using the Office Dialog API — used when NAA is not available.
 * Opens a dialog window that drives the Authentik OAuth2 authorization code flow.
 */
async function signInViaDialog(): Promise<void> {
  const config = getConfig();

  return new Promise((resolve, reject) => {
    const dialogUrl = `${config.addinHost}/auth-dialog.html`;

    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 60, width: 30, displayInIframe: false },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(`Dialog failed: ${asyncResult.error.message}`));
          return;
        }

        const dialog = asyncResult.value;

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg: unknown) => {
          dialog.close();

          const messageArg = arg as { message: string };
          try {
            const message = JSON.parse(messageArg.message);

            if (message.status === "success") {
              setAuthenticated({
                msAccessToken: message.msAccessToken ?? "",
                authentikAccessToken: message.accessToken,
                authentikRefreshToken: message.refreshToken,
                tokenExpiry: message.expiresIn
                  ? Date.now() + message.expiresIn * 1000
                  : undefined,
                userDisplayName: message.displayName,
                userEmail: message.email,
              });
              resolve();
            } else {
              reject(new Error(message.error ?? "Authentication failed"));
            }
          } catch (err) {
            reject(new Error("Failed to parse auth dialog response"));
          }
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg: unknown) => {
          const eventArg = arg as { error: number };
          dialog.close();
          if (eventArg.error === 12006) {
            reject(new Error("Auth dialog was closed by the user."));
          } else {
            reject(new Error(`Dialog event error: ${eventArg.error}`));
          }
        });
      }
    );
  });
}

/**
 * Sign out: clear shared state and MSAL cache.
 */
export async function signOut(): Promise<void> {
  clearAuth();
  if (msalInstance) {
    const accounts = msalInstance.getAllAccounts();
    for (const account of accounts) {
      try {
        await msalInstance.logout({ account });
      } catch {
        // Best-effort logout
      }
    }
  }
}
