/**
 * Resolve display name and email from an Authentik OAuth2 token response.
 *
 * Authentik returns ``id_token`` (JWT) and optionally embeds profile claims in
 * the access token; it does not populate ``id_token_claims`` on the token JSON.
 */

import { type AddinConfig, getConfig } from "../shared/config";

export interface UserProfile {
  displayName: string | null;
  email: string | null;
}

type JwtClaims = Record<string, unknown>;

function decodeJwtPayload(token: string): JwtClaims | null {
  try {
    const parts = token.split(".");
    if (parts.length !== 3) return null;
    const json = atob(parts[1].replace(/-/g, "+").replace(/_/g, "/"));
    return JSON.parse(json) as JwtClaims;
  } catch {
    return null;
  }
}

function pickString(claims: JwtClaims, keys: string[]): string | null {
  for (const key of keys) {
    const value = claims[key];
    if (typeof value === "string" && value.trim()) {
      return value.trim();
    }
  }
  return null;
}

function profileFromClaims(claims: JwtClaims | null): UserProfile {
  if (!claims) {
    return { displayName: null, email: null };
  }

  const given = pickString(claims, ["given_name"]);
  const family = pickString(claims, ["family_name"]);
  const composedName =
    given && family ? `${given} ${family}` : given ?? family ?? null;

  return {
    displayName:
      pickString(claims, ["name", "preferred_username"]) ?? composedName,
    email: pickString(claims, ["email", "preferred_username"]),
  };
}

function profileFromJwt(token: string): UserProfile {
  return profileFromClaims(decodeJwtPayload(token));
}

async function fetchUserinfoProfile(
  accessToken: string,
  config: AddinConfig = getConfig()
): Promise<UserProfile> {
  const userinfoUrl = `${config.authentikBaseUrl}/application/o/userinfo/`;

  try {
    const response = await fetch(userinfoUrl, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json",
      },
    });

    if (!response.ok) return { displayName: null, email: null };

    const claims = (await response.json()) as JwtClaims;
    return profileFromClaims(claims);
  } catch {
    return { displayName: null, email: null };
  }
}

export async function resolveUserProfile(
  tokenData: Record<string, unknown>,
  config: AddinConfig = getConfig()
): Promise<UserProfile> {
  const idTokenClaims = tokenData.id_token_claims;
  if (idTokenClaims && typeof idTokenClaims === "object") {
    const profile = profileFromClaims(idTokenClaims as JwtClaims);
    if (profile.displayName || profile.email) {
      return profile;
    }
  }

  const idToken = tokenData.id_token;
  if (typeof idToken === "string") {
    const profile = profileFromJwt(idToken);
    if (profile.displayName || profile.email) {
      return profile;
    }
  }

  const accessToken = tokenData.access_token;
  if (typeof accessToken === "string") {
    const profile = profileFromJwt(accessToken);
    if (profile.displayName || profile.email) {
      return profile;
    }

    return fetchUserinfoProfile(accessToken, config);
  }

  return { displayName: null, email: null };
}
