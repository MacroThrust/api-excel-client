/**
 * Centralized configuration for the MT Data Connector Add-in.
 *
 * All configurable values (API endpoint, auth endpoints, client IDs) are
 * defined here with sensible defaults. The API base URL can be overridden
 * at runtime via the taskpane Settings panel, persisted in localStorage.
 */

import { ADDIN_HOST } from "./version";

const STORAGE_KEY_API_BASE = "mt_api_base_url";
const STORAGE_KEY_AUTHENTIK_BASE = "mt_authentik_base_url";

export interface AddinConfig {
  msalClientId: string;
  msalAuthority: string;
  msalScopes: string[];

  authentikBaseUrl: string;
  authentikClientId: string;
  authentikAuthorizeEndpoint: string;
  authentikTokenEndpoint: string;
  authentikScopes: string;

  apiBaseUrl: string;

  addinHost: string;
}

function getStoredOrDefault(key: string, defaultValue: string): string {
  try {
    const stored = localStorage.getItem(key);
    return stored ?? defaultValue;
  } catch {
    return defaultValue;
  }
}

export function getConfig(): AddinConfig {
  const authentikBase = getStoredOrDefault(
    STORAGE_KEY_AUTHENTIK_BASE,
    "https://auth.macrothrust.com"
  );

  return {
    // Microsoft Identity Platform (Azure AD)
    msalClientId: "4305925c-6f37-4f8d-b6db-ef43a636479a",
    msalAuthority: "https://login.microsoftonline.com/common",
    msalScopes: ["User.Read", "openid", "profile", "email"],

    // Authentik OAuth2 Provider
    authentikBaseUrl: authentikBase,
    authentikClientId: "macrothrust-excel",
    authentikAuthorizeEndpoint: `${authentikBase}/application/o/authorize/`,
    authentikTokenEndpoint: `${authentikBase}/application/o/token/`,
    authentikScopes: "openid profile email groups macrothrust-api",

    // Target API
    apiBaseUrl: getStoredOrDefault(STORAGE_KEY_API_BASE, "https://api.macrothrust.com/v1"),

    addinHost: ADDIN_HOST,
  };
}

export function setApiBaseUrl(url: string): void {
  localStorage.setItem(STORAGE_KEY_API_BASE, url.replace(/\/+$/, ""));
}

export function setAuthentikBaseUrl(url: string): void {
  localStorage.setItem(STORAGE_KEY_AUTHENTIK_BASE, url.replace(/\/+$/, ""));
}
