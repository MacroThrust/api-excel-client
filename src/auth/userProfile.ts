/**
 * Resolve display name and email from an Authentik OAuth2 token response.
 *
 * Authentik returns ``id_token`` (JWT) and optionally embeds profile claims in
 * the access token; it does not populate ``id_token_claims`` on the token JSON.
 */

import {
  profileFromClaims,
  profileFromJwt,
  type JwtClaims,
  type UserProfile,
} from "@macrothrust/api-client";

import { type AddinConfig, getConfig } from "../shared/config";

export type { UserProfile };

async function fetchUserinfoProfile(
  accessToken: string,
  config: AddinConfig = getConfig(),
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
  config: AddinConfig = getConfig(),
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
