/**
 * OpenAPI helpers for Excel custom-function generation.
 *
 * Shared fetch/parse/filter logic lives in @macrothrust/api-client; this module
 * adds Excel-specific function naming (functionId / functionName).
 */

import {
  extractAllDefinedScopes,
  fetchOpenApiSpec,
  fetchOpenApiSpecFromUrl,
  filterEndpointsByScopes as filterEndpointsByScopesShared,
  normalizeApiPath,
  parseEndpoints as parseSharedEndpoints,
  routeKey,
  type ApiEndpoint as SharedApiEndpoint,
  type OpenApiOperation,
  type OpenApiParameter,
  type OpenApiPathItem,
  type OpenApiSchema,
  type OpenApiSecurityRequirement,
  type OpenApiSecurityScheme,
  type OpenApiSpec,
} from "@macrothrust/api-client";

export type {
  OpenApiOperation,
  OpenApiParameter,
  OpenApiPathItem,
  OpenApiSchema,
  OpenApiSecurityRequirement,
  OpenApiSecurityScheme,
  OpenApiSpec,
};

export {
  extractAllDefinedScopes,
  fetchOpenApiSpec,
  fetchOpenApiSpecFromUrl,
  normalizeApiPath,
  routeKey,
};

export interface ApiEndpoint extends SharedApiEndpoint {
  /** e.g. "getUsers" */
  functionName: string;
  /** Uppercase ID for CustomFunctions.associate, e.g. "GETUSERS" */
  functionId: string;
}

const VERB_MAP: Record<string, string> = {
  get: "Get",
  post: "Post",
  put: "Put",
  delete: "Delete",
  patch: "Patch",
};

/** Maps normalized route keys (e.g. GET:/sources) to built-in Excel function IDs. */
export const BUILTIN_FUNCTION_BY_ROUTE: Record<string, string> = {
  "GET:/health": "GETHEALTH",
  "GET:/sources": "GETSOURCES",
  "GET:/sources/{source_id}": "GETSOURCE",
  "GET:/sources/{source_id}/datasets": "GETSOURCEDATASETS",
  "GET:/datasets": "GETDATASETS",
  "GET:/datasets/{dataset_id}": "GETDATASET",
  "GET:/datasets/{dataset_id}/series": "GETDATASETSERIES",
  "GET:/series/{series_id}": "GETSERIES",
  "GET:/series/{series_id}/observations": "GETOBSERVATIONS",
  "GET:/series/{series_id}/observations-detail": "GETOBSERVATIONSDETAIL",
};

function capitalize(s: string): string {
  return s.charAt(0).toUpperCase() + s.slice(1);
}

function snakeToPascalCase(value: string): string {
  return value
    .split("_")
    .filter(Boolean)
    .map(capitalize)
    .join("");
}

function normalizeOperationId(operationId: string): string {
  let name = operationId.replace(/[^a-zA-Z0-9_]/g, "");
  name = name.replace(/_v\d+_[\w]+_(get|post|put|delete|patch)$/i, "");
  name = name.replace(/_(get|post|put|delete|patch)$/i, "");

  const parts = name.split("_");
  const versionIdx = parts.findIndex((part) => /^v\d+$/i.test(part));
  if (versionIdx > 0) {
    name = parts.slice(0, versionIdx).join("_");
  }

  return name;
}

function operationIdToFunctionName(operationId: string, method: string): string {
  const normalized = normalizeOperationId(operationId);
  const verb = VERB_MAP[method.toLowerCase()] ?? capitalize(method.toLowerCase());

  const verbPrefixes: Record<string, string> = {
    list_: "Get",
    get_: "Get",
    create_: "Post",
    update_: "Put",
    delete_: "Delete",
  };

  for (const [prefix, prefixVerb] of Object.entries(verbPrefixes)) {
    if (normalized.startsWith(prefix)) {
      return prefixVerb + snakeToPascalCase(normalized.slice(prefix.length));
    }
  }

  return verb + snakeToPascalCase(normalized);
}

export function getBuiltinFunctionId(method: string, path: string): string | undefined {
  return BUILTIN_FUNCTION_BY_ROUTE[routeKey(method, path)];
}

export function generateFunctionName(
  method: string,
  path: string,
  operationId?: string,
): string {
  const builtinId = getBuiltinFunctionId(method, path);
  if (builtinId) return builtinId;

  const verb = VERB_MAP[method.toLowerCase()] ?? capitalize(method.toLowerCase());

  if (operationId) {
    return operationIdToFunctionName(operationId, method);
  }

  const pathPart = normalizeApiPath(path, "/v1")
    .split("/")
    .filter((s) => s && !s.startsWith("{"))
    .map((s) => s.replace(/[^a-zA-Z0-9]/g, ""))
    .map(capitalize)
    .join("");

  return `${verb}${pathPart}`;
}

export function parseEndpoints(spec: OpenApiSpec): ApiEndpoint[] {
  const base = parseSharedEndpoints(spec);
  const usedIds = new Set<string>();

  return base.map((endpoint) => {
    let funcName = generateFunctionName(
      endpoint.method,
      endpoint.path,
      endpoint.operationId,
    );

    let suffix = 2;
    const baseName = funcName;
    while (usedIds.has(funcName.toUpperCase())) {
      funcName = `${baseName}${suffix++}`;
    }
    usedIds.add(funcName.toUpperCase());

    return {
      ...endpoint,
      functionName: funcName,
      functionId: funcName.toUpperCase(),
    };
  });
}

export function filterEndpointsByScopes(
  endpoints: ApiEndpoint[],
  userScopes: string[],
): ApiEndpoint[] {
  return filterEndpointsByScopesShared(endpoints, userScopes) as ApiEndpoint[];
}
