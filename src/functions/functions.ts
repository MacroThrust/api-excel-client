/**
 * MT Custom Functions for Excel.
 *
 * All UDFs are async and use the shared runtime's auth state to call the
 * configured API endpoint. Functions return 2D string arrays that spill
 * into adjacent cells via Excel's dynamic array support.
 *
 * JSDoc tags are used by custom-functions-metadata-plugin to auto-generate
 * the functions.json metadata file at build time.
 *
 * --- Dynamic OpenAPI-driven functions ---
 *
 * In addition to the hard-coded UDFs below, the add-in can dynamically
 * generate Excel functions from an API's OpenAPI specification. After
 * authentication, call =MT.MTRELOADFUNCTIONS() or use the ribbon
 * "Reload Functions" command to:
 *   1. Fetch the OpenAPI spec from the configured API
 *   2. Determine which endpoints the user/client is permitted to call
 *      (based on OAuth2 scopes in the token vs. security requirements)
 *   3. Register only the permitted endpoints as callable custom functions
 *   4. Use Excel.CustomFunctionManager.setVisibility() (API 1.20+) to
 *      hide denied functions from autocomplete/Formula Builder entirely
 *
 * Functions are named with verb prefixes: mtGet..., mtPost..., mtPut...,
 * mtDelete..., mtPatch... derived from the HTTP method + operationId or path.
 */

import { apiRequest, ApiError } from "../shared/apiClient";
import { ADDIN_VERSION, ADDIN_NAME, BUILD_TIMESTAMP } from "../shared/version";
import {
  reloadFunctions as reloadDynamicFunctions,
  getRegistryState,
  type RegisteredFunction,
} from "./dynamicRegistry";

/* -------------------------------------------------------------------------- */
/*  Helpers                                                                    */
/* -------------------------------------------------------------------------- */

type CellValue = string | number | boolean;

async function safeApiCall(
  fn: () => Promise<CellValue[][]>,
  invocation?: CustomFunctions.CancelableInvocation
): Promise<string[][]> {
  const controller = new AbortController();

  if (invocation?.onCanceled) {
    invocation.onCanceled = () => controller.abort();
  }

  try {
    const result = await fn();
    return result.map((row) => row.map((cell) => String(cell)));
  } catch (err) {
    if (err instanceof ApiError) {
      return [[`#ERROR: ${err.message} (${err.status})`]];
    }
    return [[`#ERROR: ${err instanceof Error ? err.message : "Unknown error"}`]];
  }
}

function toMatrix(data: unknown): CellValue[][] {
  if (Array.isArray(data)) {
    if (data.length === 0) return [["(empty)"]];

    if (Array.isArray(data[0])) {
      return data as CellValue[][];
    }

    if (typeof data[0] === "object" && data[0] !== null) {
      const keys = Object.keys(data[0]);
      const header: CellValue[] = keys;
      const rows = data.map((item: Record<string, unknown>) =>
        keys.map((k) => formatCellValue(item[k]))
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
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return value;
  }
  return JSON.stringify(value);
}

/* -------------------------------------------------------------------------- */
/*  Custom Functions (UDFs)                                                   */
/* -------------------------------------------------------------------------- */

/**
 * Retrieves the list of available data sources from the API.
 * @customfunction mtGetSources
 * @param {string} [filter] Optional filter string to narrow results.
 * @cancelable
 */
async function mtGetSources(
  filter?: string,
  invocation?: CustomFunctions.CancelableInvocation
): Promise<string[][]> {
  return safeApiCall(async () => {
    const data = await apiRequest({
      path: "/sources",
      params: filter ? { filter } : undefined,
    });
    return toMatrix(data);
  }, invocation);
}

/**
 * Fetches records from a specific data source.
 * @customfunction mtGetRecords
 * @param {string} source The data source identifier.
 * @param {number} [limit] Maximum number of records to return.
 * @param {number} [offset] Number of records to skip (for pagination).
 * @param {string} [filter] Optional filter expression.
 * @cancelable
 */
async function mtGetRecords(
  source: string,
  limit?: number,
  offset?: number,
  filter?: string,
  invocation?: CustomFunctions.CancelableInvocation
): Promise<string[][]> {
  return safeApiCall(async () => {
    const data = await apiRequest({
      path: `/sources/${encodeURIComponent(source)}/records`,
      params: {
        limit: limit ?? undefined,
        offset: offset ?? undefined,
        filter: filter ?? undefined,
      },
    });
    return toMatrix(data);
  }, invocation);
}

/**
 * Looks up a single record by ID from the specified source.
 * @customfunction mtGetRecord
 * @param {string} source The data source identifier.
 * @param {string} recordId The unique record identifier.
 * @cancelable
 */
async function mtGetRecord(
  source: string,
  recordId: string,
  invocation?: CustomFunctions.CancelableInvocation
): Promise<string[][]> {
  return safeApiCall(async () => {
    const data = await apiRequest({
      path: `/sources/${encodeURIComponent(source)}/records/${encodeURIComponent(recordId)}`,
    });
    return toMatrix(data);
  }, invocation);
}

/**
 * Retrieves metadata/schema for a given data source.
 * @customfunction mtGetSchema
 * @param {string} source The data source identifier.
 * @cancelable
 */
async function mtGetSchema(
  source: string,
  invocation?: CustomFunctions.CancelableInvocation
): Promise<string[][]> {
  return safeApiCall(async () => {
    const data = await apiRequest({
      path: `/sources/${encodeURIComponent(source)}/schema`,
    });
    return toMatrix(data);
  }, invocation);
}

/**
 * Executes a search/query against the API and returns matching results.
 * @customfunction mtSearch
 * @param {string} query The search query string.
 * @param {string} [source] Optional data source to limit the search.
 * @param {number} [limit] Maximum number of results.
 * @cancelable
 */
async function mtSearch(
  query: string,
  source?: string,
  limit?: number,
  invocation?: CustomFunctions.CancelableInvocation
): Promise<string[][]> {
  return safeApiCall(async () => {
    const data = await apiRequest({
      path: "/search",
      params: {
        q: query,
        source: source ?? undefined,
        limit: limit ?? undefined,
      },
    });
    return toMatrix(data);
  }, invocation);
}

/**
 * Retrieves aggregated/summary statistics for a data source.
 * @customfunction mtGetSummary
 * @param {string} source The data source identifier.
 * @param {string} [metric] Optional metric name (e.g., "count", "sum", "avg").
 * @param {string} [field] Optional field to aggregate on.
 * @param {string} [filter] Optional filter expression.
 * @cancelable
 */
async function mtGetSummary(
  source: string,
  metric?: string,
  field?: string,
  filter?: string,
  invocation?: CustomFunctions.CancelableInvocation
): Promise<string[][]> {
  return safeApiCall(async () => {
    const data = await apiRequest({
      path: `/sources/${encodeURIComponent(source)}/summary`,
      params: {
        metric: metric ?? undefined,
        field: field ?? undefined,
        filter: filter ?? undefined,
      },
    });
    return toMatrix(data);
  }, invocation);
}

/**
 * Returns the current authentication status and user info.
 * Useful for verifying connectivity and active session.
 * @customfunction mtStatus
 */
async function mtStatus(): Promise<string[][]> {
  try {
    const data = await apiRequest<{ status: string; user?: string; version?: string }>({
      path: "/status",
    });
    return toMatrix(data).map((row) => row.map((cell) => String(cell)));
  } catch (err) {
    if (err instanceof ApiError && err.status === 401) {
      return [
        ["status", "Not authenticated"],
        ["message", "Please sign in via the MT Menu."],
      ];
    }
    return [
      ["status", "Error"],
      ["message", err instanceof Error ? err.message : "Unknown error"],
    ];
  }
}

/**
 * Makes a generic API call to a custom endpoint path.
 * This is an escape-hatch for endpoints not covered by the specific mt functions.
 * @customfunction mtApiCall
 * @param {string} path The API path (e.g., "/custom/endpoint").
 * @param {string} [param1Name] First optional query parameter name.
 * @param {string} [param1Value] First optional query parameter value.
 * @param {string} [param2Name] Second optional query parameter name.
 * @param {string} [param2Value] Second optional query parameter value.
 * @param {string} [param3Name] Third optional query parameter name.
 * @param {string} [param3Value] Third optional query parameter value.
 * @cancelable
 */
async function mtApiCall(
  path: string,
  param1Name?: string,
  param1Value?: string,
  param2Name?: string,
  param2Value?: string,
  param3Name?: string,
  param3Value?: string,
  invocation?: CustomFunctions.CancelableInvocation
): Promise<string[][]> {
  return safeApiCall(async () => {
    const params: Record<string, string | undefined> = {};
    if (param1Name) params[param1Name] = param1Value;
    if (param2Name) params[param2Name] = param2Value;
    if (param3Name) params[param3Name] = param3Value;

    const data = await apiRequest({
      path,
      params: Object.keys(params).length > 0 ? params : undefined,
    });
    return toMatrix(data);
  }, invocation);
}

/**
 * Returns the add-in version, name, and build timestamp embedded at compile time.
 * @customfunction mtVersion
 */
async function mtVersion(): Promise<string[][]> {
  return [
    ["name", ADDIN_NAME],
    ["version", ADDIN_VERSION],
    ["built", BUILD_TIMESTAMP],
  ];
}

/* -------------------------------------------------------------------------- */
/*  OpenAPI-driven dynamic functions                                          */
/* -------------------------------------------------------------------------- */

/**
 * Reloads dynamic functions from the API's OpenAPI specification.
 * Fetches the spec, checks user permissions, and registers only the
 * permitted endpoints as callable custom functions. Call this after
 * signing in or when you want to refresh the available functions.
 * @customfunction mtReloadFunctions
 */
async function mtReloadFunctions(): Promise<string[][]> {
  try {
    const result = await reloadDynamicFunctions();
    if (result.error) {
      return [
        ["status", "Error"],
        ["message", result.error],
      ];
    }
    return [
      ["status", "OK"],
      ["total_endpoints", String(result.total)],
      ["permitted", String(result.permitted)],
      ["denied", String(result.denied)],
      ["message", result.permitted > 0
        ? `${result.permitted} function(s) registered. Use =MT.<name>() to call them.`
        : "No endpoints found or no permissions. Check your API and OpenAPI spec."],
    ];
  } catch (err) {
    return [
      ["status", "Error"],
      ["message", err instanceof Error ? err.message : "Unknown error"],
    ];
  }
}

/**
 * Lists all dynamically discovered API endpoints and whether
 * they are available to the current user based on their scopes.
 * Call mtReloadFunctions first to populate the list.
 * @customfunction mtListEndpoints
 */
async function mtListEndpoints(): Promise<string[][]> {
  const state = getRegistryState();
  if (!state.loaded) {
    return [["No OpenAPI spec loaded. Call =MT.MTRELOADFUNCTIONS() first."]];
  }
  if (state.functions.length === 0) {
    return [["No endpoints found in the OpenAPI spec."]];
  }

  const header: string[] = [
    "Function",
    "Method",
    "Path",
    "Permitted",
    "Scopes Required",
    "Description",
  ];

  const rows = state.functions.map((f: RegisteredFunction) => [
    `=MT.${f.endpoint.functionId}`,
    f.endpoint.method,
    f.endpoint.path,
    f.registered ? "Yes" : "No",
    f.endpoint.requiredScopes.join(", ") || "(none)",
    f.endpoint.summary || f.endpoint.description,
  ]);

  return [header, ...rows];
}

/* -------------------------------------------------------------------------- */
/*  Registration                                                              */
/* -------------------------------------------------------------------------- */

CustomFunctions.associate("MTVERSION", mtVersion);
CustomFunctions.associate("MTGETSOURCES", mtGetSources);
CustomFunctions.associate("MTGETRECORDS", mtGetRecords);
CustomFunctions.associate("MTGETRECORD", mtGetRecord);
CustomFunctions.associate("MTGETSCHEMA", mtGetSchema);
CustomFunctions.associate("MTSEARCH", mtSearch);
CustomFunctions.associate("MTGETSUMMARY", mtGetSummary);
CustomFunctions.associate("MTSTATUS", mtStatus);
CustomFunctions.associate("MTAPICALL", mtApiCall);
CustomFunctions.associate("MTRELOADFUNCTIONS", mtReloadFunctions);
CustomFunctions.associate("MTLISTENDPOINTS", mtListEndpoints);
