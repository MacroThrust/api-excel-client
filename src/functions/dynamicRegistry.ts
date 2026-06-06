/**
 * Dynamic function registry.
 *
 * At runtime (after auth), fetches the API's OpenAPI spec, determines which
 * endpoints the current user/client can access based on their OAuth2 scopes,
 * and registers an Excel custom function for each permitted endpoint.
 *
 * Technical approach — how visibility & registration work together:
 *
 *   - `functions.json` (metadata) defines the full universe of possible
 *     functions at build time.  Dynamic (OpenAPI-generated) entries are
 *     emitted with `"excludeFromAutoComplete": true` so they start hidden.
 *   - `CustomFunctions.associate(id, impl)` connects each metadata entry
 *     to a JS implementation at runtime.
 *   - `Excel.CustomFunctionManager.setVisibility()` (API 1.20+) can then
 *     show only the permitted functions in autocomplete / Formula Builder,
 *     and hide the denied ones — so users never see functions they can't
 *     call.  On older Excel versions that lack this API, denied functions
 *     simply remain hidden (never associated) and return #NAME?.
 *   - A "reload" function (`reloadFunctions` / =MT.RELOADFUNCTIONS()) triggers re-evaluation.
 */

import { apiRequest, ApiError } from "../shared/apiClient";
import { getConfig } from "../shared/config";
import { getAuthState, isTokenExpired } from "../shared/state";
import {
  type ApiEndpoint,
  type OpenApiSpec,
  fetchOpenApiSpec,
  fetchOpenApiSpecFromUrl,
  parseEndpoints,
  filterEndpointsByScopes,
  getBuiltinFunctionId,
} from "../shared/openApiClient";

/* ------------------------------------------------------------------ */
/*  Helpers (shared with functions.ts — duplicated to avoid circular) */
/* ------------------------------------------------------------------ */

type CellValue = string | number | boolean;

function toMatrix(data: unknown): CellValue[][] {
  if (Array.isArray(data)) {
    if (data.length === 0) return [["(empty)"]];
    if (Array.isArray(data[0])) return data as CellValue[][];
    if (typeof data[0] === "object" && data[0] !== null) {
      const keys = Object.keys(data[0]);
      const header: CellValue[] = keys;
      const rows = data.map((item: Record<string, unknown>) =>
        keys.map((k) => formatCellValue(item[k])),
      );
      return [header, ...rows];
    }
    return data.map((v: unknown) => [formatCellValue(v)]);
  }
  if (typeof data === "object" && data !== null) {
    const entries = Object.entries(data as Record<string, unknown>);
    if (entries.length === 0) return [["(empty object)"]];
    return entries.map(([k, v]) => [k, formatCellValue(v)]);
  }
  return [[formatCellValue(data)]];
}

function formatCellValue(value: unknown): CellValue {
  if (value === null || value === undefined) return "";
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean")
    return value;
  return JSON.stringify(value);
}

/* ------------------------------------------------------------------ */
/*  Registry state                                                    */
/* ------------------------------------------------------------------ */

export interface RegisteredFunction {
  endpoint: ApiEndpoint;
  registered: boolean;
  /** True when a hand-written MT.* function already covers this route. */
  usesBuiltin?: boolean;
}

let cachedSpec: OpenApiSpec | null = null;
let allEndpoints: ApiEndpoint[] = [];
let registeredFunctions: Map<string, RegisteredFunction> = new Map();
let lastLoadTime: number | null = null;
let lastError: string | null = null;

export function getRegistryState() {
  return {
    loaded: cachedSpec !== null,
    endpointCount: allEndpoints.length,
    registeredCount: Array.from(registeredFunctions.values()).filter((f) => f.registered).length,
    lastLoadTime,
    lastError,
    functions: Array.from(registeredFunctions.values()),
  };
}

/* ------------------------------------------------------------------ */
/*  User scope resolution                                             */
/* ------------------------------------------------------------------ */

interface PermissionsResponse {
  scopes?: string[];
  permissions?: string[];
  allowed_scopes?: string[];
  roles?: string[];
}

async function fetchUserScopes(): Promise<string[]> {
  const config = getConfig();

  const permPaths = ["/permissions", "/me/permissions", "/auth/permissions", "/me/scopes"];

  for (const path of permPaths) {
    try {
      const data = await apiRequest<PermissionsResponse>({ path });
      const scopes =
        data.scopes ?? data.permissions ?? data.allowed_scopes ?? data.roles ?? [];
      if (Array.isArray(scopes) && scopes.length > 0) return scopes.map(String);
    } catch {
      /* try next */
    }
  }

  const tokenInfo = parseJwtScopes();
  if (tokenInfo.length > 0) return tokenInfo;

  return [];
}

function parseJwtScopes(): string[] {
  const auth = getAuthState();
  const token = auth.authentikAccessToken;
  if (!token) return [];
  try {
    const parts = token.split(".");
    if (parts.length !== 3) return [];
    const payload = JSON.parse(atob(parts[1].replace(/-/g, "+").replace(/_/g, "/")));
    const scopeStr: string = payload.scope ?? payload.scp ?? "";
    if (scopeStr) return scopeStr.split(/\s+/).filter(Boolean);
    if (Array.isArray(payload.scopes)) return payload.scopes;
    if (Array.isArray(payload.permissions)) return payload.permissions;
    return [];
  } catch {
    return [];
  }
}

/* ------------------------------------------------------------------ */
/*  Dynamic function factory                                          */
/* ------------------------------------------------------------------ */

function buildUrl(pathTemplate: string, pathArgs: Record<string, string>): string {
  let url = pathTemplate;
  for (const [key, val] of Object.entries(pathArgs)) {
    url = url.replace(`{${key}}`, encodeURIComponent(val));
  }
  return url;
}

function createDynamicFunction(
  ep: ApiEndpoint,
): (...args: unknown[]) => Promise<string[][]> {
  return async function dynamicEndpointFunction(
    ...args: unknown[]
  ): Promise<string[][]> {
    const auth = getAuthState();
    if (!auth.isAuthenticated || !auth.authentikAccessToken) {
      return [["#ERROR: Not authenticated. Please sign in first."]];
    }
    if (isTokenExpired()) {
      return [["#ERROR: Session expired. Please sign in again."]];
    }

    try {
      const pathArgs: Record<string, string> = {};
      let argIdx = 0;

      for (const pp of ep.pathParameters) {
        const val = args[argIdx++];
        if (val === undefined || val === null || val === "") {
          if (pp.required !== false) {
            return [[`#ERROR: Missing required parameter: ${pp.name}`]];
          }
          continue;
        }
        pathArgs[pp.name] = String(val);
      }

      const queryParams: Record<string, string | undefined> = {};
      for (const qp of ep.queryParameters) {
        const val = args[argIdx++];
        if (val !== undefined && val !== null && val !== "") {
          queryParams[qp.name] = String(val);
        }
      }

      if (ep.hasRequestBody && args[argIdx] !== undefined) {
        const bodyStr = String(args[argIdx++]);
        let body: unknown;
        try {
          body = JSON.parse(bodyStr);
        } catch {
          body = bodyStr;
        }

        const data = await apiRequest({
          method: ep.method,
          path: buildUrl(ep.path, pathArgs),
          params: Object.keys(queryParams).length > 0 ? queryParams : undefined,
          body,
        });
        return toMatrix(data).map((r) => r.map((c) => String(c)));
      }

      const data = await apiRequest({
        method: ep.method,
        path: buildUrl(ep.path, pathArgs),
        params: Object.keys(queryParams).length > 0 ? queryParams : undefined,
      });
      return toMatrix(data).map((r) => r.map((c) => String(c)));
    } catch (err) {
      if (err instanceof ApiError) {
        return [[`#ERROR: ${err.message} (${err.status})`]];
      }
      return [[`#ERROR: ${err instanceof Error ? err.message : "Unknown error"}`]];
    }
  };
}

function createDeniedFunction(
  ep: ApiEndpoint,
): (...args: unknown[]) => Promise<string[][]> {
  const scopes = ep.requiredScopes.join(", ") || "unknown";
  return async function deniedEndpointFunction(): Promise<string[][]> {
    return [[`#ERROR: Access denied. ${ep.method} ${ep.path} requires scope(s): ${scopes}`]];
  };
}

/* ------------------------------------------------------------------ */
/*  Core: load spec, filter, register                                 */
/* ------------------------------------------------------------------ */

export async function loadAndRegisterFunctions(): Promise<{
  total: number;
  permitted: number;
  denied: number;
  error?: string;
}> {
  const config = getConfig();
  const auth = getAuthState();
  lastError = null;

  try {
    if (!auth.isAuthenticated || !auth.authentikAccessToken) {
      throw new Error("Not authenticated");
    }

    let spec: OpenApiSpec;
    const specUrl = getOpenApiSpecUrl();
    if (specUrl) {
      spec = await fetchOpenApiSpecFromUrl(specUrl, auth.authentikAccessToken);
    } else {
      spec = await fetchOpenApiSpec(config.apiBaseUrl, auth.authentikAccessToken);
    }

    cachedSpec = spec;
    allEndpoints = parseEndpoints(spec);

    const userScopes = await fetchUserScopes();

    let permitted: ApiEndpoint[];
    if (userScopes.length > 0) {
      permitted = filterEndpointsByScopes(allEndpoints, userScopes);
    } else {
      permitted = allEndpoints;
    }

    registeredFunctions.clear();

    for (const ep of allEndpoints) {
      const builtinId = getBuiltinFunctionId(ep.method, ep.path);
      const isPermitted = builtinId ? true : permitted.includes(ep);
      registeredFunctions.set(ep.functionId, {
        endpoint: ep,
        registered: isPermitted,
        usesBuiltin: !!builtinId,
      });
    }

    const permittedSet = new Set(
      permitted
        .filter((ep) => !getBuiltinFunctionId(ep.method, ep.path))
        .map((ep) => ep.functionId),
    );

    for (const ep of allEndpoints) {
      if (getBuiltinFunctionId(ep.method, ep.path)) {
        continue;
      }
      try {
        const impl = permittedSet.has(ep.functionId)
          ? createDynamicFunction(ep)
          : createDeniedFunction(ep);
        CustomFunctions.associate(ep.functionId, impl);
      } catch (err) {
        console.warn(`[MT] Failed to associate ${ep.functionId}:`, err);
      }
    }

    await updateFunctionVisibility(permitted, allEndpoints);

    lastLoadTime = Date.now();

    return {
      total: allEndpoints.length,
      permitted: permitted.length,
      denied: allEndpoints.length - permitted.length,
    };
  } catch (err) {
    const msg = err instanceof Error ? err.message : "Unknown error";
    lastError = msg;
    return { total: 0, permitted: 0, denied: 0, error: msg };
  }
}

/* ------------------------------------------------------------------ */
/*  Visibility control (Excel API 1.20+)                              */
/* ------------------------------------------------------------------ */

let visibilityApiAvailable: boolean | null = null;

function isVisibilityApiSupported(): boolean {
  if (visibilityApiAvailable !== null) return visibilityApiAvailable;
  try {
    visibilityApiAvailable =
      typeof Office !== "undefined" &&
      Office.context?.requirements?.isSetSupported("ExcelApi", "1.20") === true;
  } catch {
    visibilityApiAvailable = false;
  }
  return visibilityApiAvailable;
}

async function updateFunctionVisibility(
  permitted: ApiEndpoint[],
  all: ApiEndpoint[],
): Promise<void> {
  if (!isVisibilityApiSupported()) {
    console.log("[MT] Excel API 1.20 not available — visibility control skipped.");
    return;
  }

  try {
    const permittedIds = permitted
      .filter((ep) => !getBuiltinFunctionId(ep.method, ep.path))
      .map((ep) => ep.functionId);
    const deniedIds = all
      .filter((ep) => !permitted.includes(ep) && !getBuiltinFunctionId(ep.method, ep.path))
      .map((ep) => ep.functionId);

    await Excel.run(async (context) => {
      const options: Excel.CustomFunctionVisibilityOptions = {};
      if (permittedIds.length > 0) {
        options.show = permittedIds;
      }
      if (deniedIds.length > 0) {
        options.hide = deniedIds;
      }
      (Excel as any).CustomFunctionManager.setVisibility(options);
      await context.sync();
    });

    console.log(
      `[MT] Visibility updated: ${permittedIds.length} shown, ${deniedIds.length} hidden.`,
    );
  } catch (err) {
    console.warn("[MT] setVisibility failed (non-fatal):", err);
  }
}

/* ------------------------------------------------------------------ */
/*  Config helpers for OpenAPI spec URL                               */
/* ------------------------------------------------------------------ */

const STORAGE_KEY_OPENAPI_URL = "mt_openapi_spec_url";

export function getOpenApiSpecUrl(): string | null {
  try {
    return localStorage.getItem(STORAGE_KEY_OPENAPI_URL);
  } catch {
    return null;
  }
}

export function setOpenApiSpecUrl(url: string): void {
  try {
    localStorage.setItem(STORAGE_KEY_OPENAPI_URL, url.replace(/\/+$/, ""));
  } catch {
    /* noop */
  }
}

/* ------------------------------------------------------------------ */
/*  Change listeners (for taskpane)                                   */
/* ------------------------------------------------------------------ */

type RegistryChangeListener = () => void;
const changeListeners: RegistryChangeListener[] = [];

export function onRegistryChange(listener: RegistryChangeListener): () => void {
  changeListeners.push(listener);
  return () => {
    const idx = changeListeners.indexOf(listener);
    if (idx >= 0) changeListeners.splice(idx, 1);
  };
}

function notifyRegistryListeners(): void {
  changeListeners.forEach((fn) => {
    try {
      fn();
    } catch {
      /* noop */
    }
  });
}

export async function reloadFunctions(): Promise<{
  total: number;
  permitted: number;
  denied: number;
  error?: string;
}> {
  const result = await loadAndRegisterFunctions();
  notifyRegistryListeners();
  return result;
}
