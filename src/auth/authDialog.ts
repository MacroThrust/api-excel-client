/**
 * Auth dialog entry point.
 *
 * This script runs inside the Office Dialog popup. It redirects the user
 * to Authentik's OAuth2 authorize endpoint. After the user authenticates
 * (via their Microsoft Account configured in Authentik), Authentik redirects
 * back to this page with an authorization code. The code is then exchanged
 * for tokens and the result is sent back to the parent via messageParent().
 */

import { getConfig } from "../shared/config";
import { resolveUserProfile } from "./userProfile";

function generateRandomString(length: number): string {
  const array = new Uint8Array(length);
  crypto.getRandomValues(array);
  return Array.from(array, (b) => b.toString(36).padStart(2, "0"))
    .join("")
    .substring(0, length);
}

async function generateCodeChallenge(verifier: string): Promise<string> {
  const encoder = new TextEncoder();
  const data = encoder.encode(verifier);
  const digest = await crypto.subtle.digest("SHA-256", data);
  return btoa(String.fromCharCode(...new Uint8Array(digest)))
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");
}

async function startAuth(): Promise<void> {
  const config = getConfig();
  const params = new URLSearchParams(window.location.search);
  const code = params.get("code");

  if (code) {
    await handleCallback(code, config);
  } else {
    await redirectToAuthentik(config);
  }
}

async function redirectToAuthentik(
  config: ReturnType<typeof getConfig>
): Promise<void> {
  const state = generateRandomString(32);
  const codeVerifier = generateRandomString(64);
  const codeChallenge = await generateCodeChallenge(codeVerifier);

  sessionStorage.setItem("oauth_state", state);
  sessionStorage.setItem("oauth_code_verifier", codeVerifier);

  const redirectUri = `${config.addinHost}/auth-dialog.html`;

  const authUrl = new URL(config.authentikAuthorizeEndpoint);
  authUrl.searchParams.set("client_id", config.authentikClientId);
  authUrl.searchParams.set("response_type", "code");
  authUrl.searchParams.set("redirect_uri", redirectUri);
  authUrl.searchParams.set("scope", config.authentikScopes);
  authUrl.searchParams.set("state", state);
  authUrl.searchParams.set("code_challenge", codeChallenge);
  authUrl.searchParams.set("code_challenge_method", "S256");

  window.location.href = authUrl.toString();
}

async function handleCallback(
  code: string,
  config: ReturnType<typeof getConfig>
): Promise<void> {
  const params = new URLSearchParams(window.location.search);
  const returnedState = params.get("state");
  const savedState = sessionStorage.getItem("oauth_state");
  const codeVerifier = sessionStorage.getItem("oauth_code_verifier");

  if (returnedState !== savedState) {
    sendError("State mismatch — possible CSRF attack.");
    return;
  }

  if (!codeVerifier) {
    sendError("Missing code verifier.");
    return;
  }

  const redirectUri = `${config.addinHost}/auth-dialog.html`;

  try {
    const tokenResponse = await fetch(config.authentikTokenEndpoint, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "authorization_code",
        client_id: config.authentikClientId,
        code,
        redirect_uri: redirectUri,
        code_verifier: codeVerifier,
      }).toString(),
    });

    if (!tokenResponse.ok) {
      const errText = await tokenResponse.text();
      sendError(`Token exchange failed: ${errText}`);
      return;
    }

    const tokenData = (await tokenResponse.json()) as Record<string, unknown>;
    const profile = await resolveUserProfile(tokenData, config);

    Office.context.ui.messageParent(
      JSON.stringify({
        status: "success",
        accessToken: tokenData.access_token,
        refreshToken: tokenData.refresh_token,
        expiresIn: tokenData.expires_in,
        displayName: profile.displayName,
        email: profile.email,
      })
    );
  } catch (err) {
    sendError(`Token exchange error: ${err instanceof Error ? err.message : String(err)}`);
  } finally {
    sessionStorage.removeItem("oauth_state");
    sessionStorage.removeItem("oauth_code_verifier");
  }
}

function sendError(message: string): void {
  Office.context.ui.messageParent(
    JSON.stringify({ status: "error", error: message })
  );
}

Office.onReady(() => {
  startAuth().catch((err) => {
    sendError(`Auth initialization failed: ${err instanceof Error ? err.message : String(err)}`);
  });
});
