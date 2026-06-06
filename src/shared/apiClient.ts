/**
 * HTTP client for the target API.
 *
 * All MT custom functions route through this client. It automatically attaches
 * the Authentik bearer token from shared state and targets the configured API
 * base URL. Supports JSON request/response and basic error mapping.
 */

import { getConfig } from "./config";
import { getAuthState, isTokenExpired } from "./state";

export class ApiError extends Error {
  constructor(
    message: string,
    public status: number,
    public body?: unknown
  ) {
    super(message);
    this.name = "ApiError";
  }
}

export interface ApiRequestOptions {
  method?: "GET" | "POST" | "PUT" | "DELETE" | "PATCH";
  path: string;
  params?: Record<string, string | number | boolean | undefined>;
  body?: unknown;
  signal?: AbortSignal;
}

function buildQueryString(params: Record<string, string | number | boolean | undefined>): string {
  const entries = Object.entries(params).filter(
    (entry): entry is [string, string | number | boolean] => entry[1] !== undefined
  );
  if (entries.length === 0) return "";
  const qs = new URLSearchParams();
  entries.forEach(([k, v]) => qs.append(k, String(v)));
  return `?${qs.toString()}`;
}

export async function apiRequest<T = unknown>(options: ApiRequestOptions): Promise<T> {
  const config = getConfig();
  // getAuthState() hydrates from localStorage so custom-function bundles see sign-in.
  const authState = getAuthState();

  if (!authState.isAuthenticated || !authState.authentikAccessToken) {
    throw new ApiError("Not authenticated. Please sign in first.", 401);
  }

  if (isTokenExpired()) {
    throw new ApiError("Session expired. Please sign in again.", 401);
  }

  const url =
    `${config.apiBaseUrl}${options.path}` +
    (options.params ? buildQueryString(options.params) : "");

  const headers: Record<string, string> = {
    Authorization: `Bearer ${authState.authentikAccessToken}`,
    "Content-Type": "application/json",
    Accept: "application/json",
  };

  const response = await fetch(url, {
    method: options.method ?? "GET",
    headers,
    body: options.body ? JSON.stringify(options.body) : undefined,
    signal: options.signal,
  });

  if (!response.ok) {
    let body: unknown;
    try {
      body = await response.json();
    } catch {
      body = await response.text();
    }
    throw new ApiError(
      `API request failed: ${response.status} ${response.statusText}`,
      response.status,
      body
    );
  }

  return response.json() as Promise<T>;
}
