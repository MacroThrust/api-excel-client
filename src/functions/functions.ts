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
 * authentication, call =MT.RELOADFUNCTIONS() or use the ribbon
 * "Reload Functions" command to:
 *   1. Fetch the OpenAPI spec from the configured API
 *   2. Determine which endpoints the user/client is permitted to call
 *      (based on OAuth2 scopes in the token vs. security requirements)
 *   3. Register only the permitted endpoints as callable custom functions
 *   4. Use Excel.CustomFunctionManager.setVisibility() (API 1.20+) to
 *      hide denied functions from autocomplete/Formula Builder entirely
 *
 * Functions are named with HTTP verb prefixes: get..., post..., put...,
 * delete..., patch... derived from the HTTP method + operationId or path.
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

function optionalNumber(value: number | undefined): number | undefined {
  if (value === undefined || value === null) return undefined;
  const n = Number(value);
  return Number.isFinite(n) ? n : undefined;
}

function optionalBoolean(
  value: boolean | string | undefined,
  defaultValue?: boolean
): boolean | undefined {
  if (value === undefined || value === null) return defaultValue;
  if (typeof value === "boolean") return value;
  if (typeof value === "string") {
    const lower = value.toLowerCase();
    if (lower === "true" || lower === "1") return true;
    if (lower === "false" || lower === "0") return false;
  }
  return defaultValue;
}

/** Unwrap MacroThrust API list/detail envelopes for tabular Excel output. */
function flattenApiResponse(data: unknown): unknown {
  if (typeof data !== "object" || data === null) return data;
  const obj = data as Record<string, unknown>;

  if (Array.isArray(obj.items)) return obj.items;
  if (Array.isArray(obj.dataset_ids)) {
    return obj.dataset_ids.map((id) => ({ dataset_id: id }));
  }
  if (Array.isArray(obj.series_ids)) {
    return obj.series_ids.map((id) => ({ series_id: id }));
  }
  if (Array.isArray(obj.datasets)) return obj.datasets;
  if (Array.isArray(obj.series)) return obj.series;
  if (Array.isArray(obj.observations)) return obj.observations;
  if (obj.source && typeof obj.source === "object" && !Array.isArray(obj.source)) {
    return obj.source;
  }
  if (obj.dataset && typeof obj.dataset === "object" && !Array.isArray(obj.dataset)) {
    return obj.dataset;
  }
  if (obj.series && typeof obj.series === "object" && !Array.isArray(obj.series)) {
    return obj.series;
  }

  return data;
}

/* -------------------------------------------------------------------------- */
/*  Custom Functions (UDFs)                                                   */
/* -------------------------------------------------------------------------- */

/**
 * Retrieves the list of available data sources from the API.
 * Maps to GET /sources.
 * @customfunction getSources
 * @param {number} [offset] Number of sources to skip (pagination).
 * @param {number} [limit] Maximum number of sources to return (1–500).
 * @cancelable
 */
async function mtGetSources(
  offset?: number,
  limit?: number,
  invocation?: CustomFunctions.CancelableInvocation
): Promise<string[][]> {
  return safeApiCall(async () => {
    const data = await apiRequest({
      path: "/sources",
      params: {
        offset: optionalNumber(offset),
        limit: optionalNumber(limit),
      },
    });
    return toMatrix(flattenApiResponse(data));
  }, invocation);
}

/**
 * Looks up a single data source by ID.
 * Maps to GET /sources/{source_id}.
 * @customfunction getSource
 * @param {string} sourceId The data source identifier (e.g. FRED).
 * @cancelable
 */
async function mtGetSource(
  sourceId: string,
  invocation?: CustomFunctions.CancelableInvocation
): Promise<string[][]> {
  return safeApiCall(async () => {
    const data = await apiRequest({
      path: `/sources/${encodeURIComponent(sourceId)}`,
    });
    return toMatrix(flattenApiResponse(data));
  }, invocation);
}

/**
 * Lists datasets registered under a data source.
 * Maps to GET /sources/{source_id}/datasets.
 * @customfunction getSourceDatasets
 * @param {string} sourceId The data source identifier.
 * @param {number} [limit] Maximum number of datasets to return (1–500).
 * @param {number} [offset] Number of datasets to skip (pagination).
 * @param {boolean} [idsOnly] If true, return only dataset IDs.
 * @cancelable
 */
async function mtGetSourceDatasets(
  sourceId: string,
  limit?: number,
  offset?: number,
  idsOnly?: boolean,
  invocation?: CustomFunctions.CancelableInvocation
): Promise<string[][]> {
  return safeApiCall(async () => {
    const data = await apiRequest({
      path: `/sources/${encodeURIComponent(sourceId)}/datasets`,
      params: {
        limit: optionalNumber(limit),
        offset: optionalNumber(offset),
        ids_only: optionalBoolean(idsOnly, false),
      },
    });
    return toMatrix(flattenApiResponse(data));
  }, invocation);
}

/**
 * Lists datasets, optionally filtered by source.
 * Maps to GET /datasets.
 * @customfunction getDatasets
 * @param {string} [sourceId] Filter to datasets under this source.
 * @param {number} [offset] Number of datasets to skip (pagination).
 * @param {number} [limit] Maximum number of datasets to return (1–500).
 * @cancelable
 */
async function mtGetDatasets(
  sourceId?: string,
  offset?: number,
  limit?: number,
  invocation?: CustomFunctions.CancelableInvocation
): Promise<string[][]> {
  return safeApiCall(async () => {
    const data = await apiRequest({
      path: "/datasets",
      params: {
        source_id: sourceId ?? undefined,
        offset: optionalNumber(offset),
        limit: optionalNumber(limit),
      },
    });
    return toMatrix(flattenApiResponse(data));
  }, invocation);
}

/**
 * Looks up a single dataset by ID.
 * Maps to GET /datasets/{dataset_id}.
 * @customfunction getDataset
 * @param {string} datasetId The dataset identifier.
 * @cancelable
 */
async function mtGetDataset(
  datasetId: string,
  invocation?: CustomFunctions.CancelableInvocation
): Promise<string[][]> {
  return safeApiCall(async () => {
    const data = await apiRequest({
      path: `/datasets/${encodeURIComponent(datasetId)}`,
    });
    return toMatrix(flattenApiResponse(data));
  }, invocation);
}

/**
 * Lists series registered under a dataset.
 * Maps to GET /datasets/{dataset_id}/series.
 * @customfunction getDatasetSeries
 * @param {string} datasetId The dataset identifier.
 * @param {number} [limit] Maximum number of series to return (1–500).
 * @param {number} [offset] Number of series to skip (pagination).
 * @param {boolean} [idsOnly] If true, return only series IDs.
 * @cancelable
 */
async function mtGetDatasetSeries(
  datasetId: string,
  limit?: number,
  offset?: number,
  idsOnly?: boolean,
  invocation?: CustomFunctions.CancelableInvocation
): Promise<string[][]> {
  return safeApiCall(async () => {
    const data = await apiRequest({
      path: `/datasets/${encodeURIComponent(datasetId)}/series`,
      params: {
        limit: optionalNumber(limit),
        offset: optionalNumber(offset),
        ids_only: optionalBoolean(idsOnly, false),
      },
    });
    return toMatrix(flattenApiResponse(data));
  }, invocation);
}

/**
 * Looks up a single series by ID.
 * Maps to GET /series/{series_id}.
 * @customfunction getSeries
 * @param {string} seriesId The series identifier.
 * @cancelable
 */
async function mtGetSeries(
  seriesId: string,
  invocation?: CustomFunctions.CancelableInvocation
): Promise<string[][]> {
  return safeApiCall(async () => {
    const data = await apiRequest({
      path: `/series/${encodeURIComponent(seriesId)}`,
    });
    return toMatrix(flattenApiResponse(data));
  }, invocation);
}

/**
 * Fetches observations (time-series values) for a series.
 * Maps to GET /series/{series_id}/observations.
 * @customfunction getObservations
 * @param {string} seriesId The series identifier.
 * @param {string} [startDate] Inclusive start date (YYYY-MM-DD).
 * @param {string} [endDate] Inclusive end date (YYYY-MM-DD).
 * @param {number} [limit] Maximum number of observations (1–10000).
 * @param {number} [offset] Number of observations to skip (pagination).
 * @cancelable
 */
async function mtGetObservations(
  seriesId: string,
  startDate?: string,
  endDate?: string,
  limit?: number,
  offset?: number,
  invocation?: CustomFunctions.CancelableInvocation
): Promise<string[][]> {
  return safeApiCall(async () => {
    const data = await apiRequest({
      path: `/series/${encodeURIComponent(seriesId)}/observations`,
      params: {
        start_date: startDate ?? undefined,
        end_date: endDate ?? undefined,
        limit: optionalNumber(limit),
        offset: optionalNumber(offset),
      },
    });
    return toMatrix(flattenApiResponse(data));
  }, invocation);
}

/**
 * Checks API connectivity and authentication.
 * Maps to GET /health.
 * @customfunction status
 */
async function mtStatus(): Promise<string[][]> {
  try {
    const data = await apiRequest<{ status: string }>({
      path: "/health",
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
 * @customfunction apiCall
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
 * @customfunction version
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
 * @customfunction reloadFunctions
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
 * @customfunction listEndpoints
 */
async function mtListEndpoints(): Promise<string[][]> {
  const state = getRegistryState();
  if (!state.loaded) {
    return [["No OpenAPI spec loaded. Call =MT.RELOADFUNCTIONS() first."]];
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

CustomFunctions.associate("VERSION", mtVersion);
CustomFunctions.associate("GETSOURCES", mtGetSources);
CustomFunctions.associate("GETSOURCE", mtGetSource);
CustomFunctions.associate("GETSOURCEDATASETS", mtGetSourceDatasets);
CustomFunctions.associate("GETDATASETS", mtGetDatasets);
CustomFunctions.associate("GETDATASET", mtGetDataset);
CustomFunctions.associate("GETDATASETSERIES", mtGetDatasetSeries);
CustomFunctions.associate("GETSERIES", mtGetSeries);
CustomFunctions.associate("GETOBSERVATIONS", mtGetObservations);
CustomFunctions.associate("STATUS", mtStatus);
CustomFunctions.associate("APICALL", mtApiCall);
CustomFunctions.associate("RELOADFUNCTIONS", mtReloadFunctions);
CustomFunctions.associate("LISTENDPOINTS", mtListEndpoints);
