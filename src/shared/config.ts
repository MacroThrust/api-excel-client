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
    "https://authentik.example.com"
  );

  return {
    // Microsoft Identity Platform (Azure AD)
    msalClientId: "YOUR_AZURE_APP_CLIENT_ID",
    msalAuthority: "https://login.microsoftonline.com/common",
    msalScopes: ["User.Read", "openid", "profile", "email"],

    // Authentik OAuth2 Provider
    authentikBaseUrl: authentikBase,
    authentikClientId: "YOUR_AUTHENTIK_CLIENT_ID",
    authentikAuthorizeEndpoint: `${authentikBase}/application/o/authorize/`,
    authentikTokenEndpoint: `${authentikBase}/application/o/token/`,
    authentikScopes: "openid profile email",

    // Target API
    apiBaseUrl: getStoredOrDefault(STORAGE_KEY_API_BASE, "https://api.example.com/v1"),

    addinHost: ADDIN_HOST,
  };
}

export function setApiBaseUrl(url: string): void {
  localStorage.setItem(STORAGE_KEY_API_BASE, url.replace(/\/+$/, ""));
}

export function setAuthentikBaseUrl(url: string): void {
  localStorage.setItem(STORAGE_KEY_AUTHENTIK_BASE, url.replace(/\/+$/, ""));
}
