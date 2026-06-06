/**
 * HTTP client for the target API.
 *
 * All MT custom functions route through this client. It automatically attaches
 * the Authentik bearer token from shared state and targets the configured API
 * base URL. Supports JSON request/response and basic error mapping.
 */

import { refreshAccessToken } from "../auth/authConfig";
import { getConfig } from "./config";
import { normalizeApiPath } from "./openApiClient";
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

function formatApiErrorMessage(status: number, statusText: string, body: unknown): string {
  let detail = "";
  if (typeof body === "object" && body !== null && "detail" in body) {
    const raw = (body as { detail: unknown }).detail;
    detail = typeof raw === "string" ? raw : JSON.stringify(raw);
  } else if (typeof body === "string" && body.trim()) {
    detail = body.trim();
  }

  if (status === 401) {
    return detail || "Unauthorized — sign in again from the MT task pane.";
  }

  return detail
    ? `API request failed: ${detail} (${status})`
    : `API request failed: ${status} ${statusText}`;
}

async function ensureFreshAccessToken(): Promise<void> {
  if (!isTokenExpired()) return;

  const refreshed = await refreshAccessToken();
  if (!refreshed) {
    throw new ApiError("Session expired. Please sign in again.", 401);
  }
}

async function performRequest<T>(
  options: ApiRequestOptions,
  retryOnUnauthorized: boolean,
): Promise<T> {
  const config = getConfig();
  const authState = getAuthState();

  if (!authState.isAuthenticated || !authState.authentikAccessToken) {
    throw new ApiError("Not authenticated. Please sign in first.", 401);
  }

  const requestPath = normalizeApiPath(options.path, config.apiBaseUrl);
  const url =
    `${config.apiBaseUrl}${requestPath}` +
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

    if (response.status === 401 && retryOnUnauthorized) {
      const refreshed = await refreshAccessToken();
      if (refreshed) {
        return performRequest(options, false);
      }
    }

    throw new ApiError(
      formatApiErrorMessage(response.status, response.statusText, body),
      response.status,
      body,
    );
  }

  return response.json() as Promise<T>;
}

export async function apiRequest<T = unknown>(options: ApiRequestOptions): Promise<T> {
  // getAuthState() hydrates from localStorage so custom-function bundles see sign-in.
  await ensureFreshAccessToken();
  return performRequest<T>(options, true);
}
