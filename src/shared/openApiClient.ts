/**
 * OpenAPI 3.x spec fetcher and parser.
 *
 * Retrieves an OpenAPI spec from the target API, parses it into a flat list
 * of endpoint descriptors, and supports filtering by OAuth2 scopes so only
 * permitted operations surface as Excel custom functions.
 */

/* ------------------------------------------------------------------ */
/*  OpenAPI 3.x type subset (enough for function generation)          */
/* ------------------------------------------------------------------ */

export interface OpenApiSpec {
  openapi: string;
  info: { title: string; version: string; description?: string };
  paths: Record<string, OpenApiPathItem>;
  components?: {
    securitySchemes?: Record<string, OpenApiSecurityScheme>;
  };
  security?: OpenApiSecurityRequirement[];
}

export interface OpenApiPathItem {
  get?: OpenApiOperation;
  post?: OpenApiOperation;
  put?: OpenApiOperation;
  delete?: OpenApiOperation;
  patch?: OpenApiOperation;
  head?: OpenApiOperation;
  options?: OpenApiOperation;
  parameters?: OpenApiParameter[];
}

export interface OpenApiOperation {
  operationId?: string;
  summary?: string;
  description?: string;
  parameters?: OpenApiParameter[];
  security?: OpenApiSecurityRequirement[];
  tags?: string[];
  requestBody?: {
    required?: boolean;
    content?: Record<string, { schema?: OpenApiSchema }>;
  };
  responses?: Record<string, unknown>;
  deprecated?: boolean;
}

export interface OpenApiParameter {
  name: string;
  in: "query" | "path" | "header" | "cookie";
  required?: boolean;
  description?: string;
  schema?: OpenApiSchema;
  deprecated?: boolean;
}

export interface OpenApiSchema {
  type?: string;
  format?: string;
  enum?: unknown[];
  description?: string;
  items?: OpenApiSchema;
  properties?: Record<string, OpenApiSchema>;
}

export interface OpenApiSecurityScheme {
  type: string;
  flows?: Record<string, { scopes?: Record<string, string> }>;
  scheme?: string;
  openIdConnectUrl?: string;
}

export type OpenApiSecurityRequirement = Record<string, string[]>;

/* ------------------------------------------------------------------ */
/*  Parsed endpoint descriptor                                        */
/* ------------------------------------------------------------------ */

export interface ApiEndpoint {
  method: "GET" | "POST" | "PUT" | "DELETE" | "PATCH";
  path: string;
  operationId: string | undefined;
  summary: string;
  description: string;
  pathParameters: OpenApiParameter[];
  queryParameters: OpenApiParameter[];
  requiredScopes: string[];
  tags: string[];
  hasRequestBody: boolean;
  deprecated: boolean;
  /** e.g. "getUsers" */
  functionName: string;
  /** Uppercase ID for CustomFunctions.associate, e.g. "GETUSERS" */
  functionId: string;
}

/* ------------------------------------------------------------------ */
/*  Spec fetching                                                     */
/* ------------------------------------------------------------------ */

const WELL_KNOWN_SPEC_PATHS = [
  "/openapi.json",
  "/api/openapi.json",
  "/swagger.json",
  "/api/swagger.json",
  "/docs/openapi.json",
  "/api-docs",
  "/v1/openapi.json",
];

export async function fetchOpenApiSpec(
  baseUrl: string,
  bearerToken?: string,
): Promise<OpenApiSpec> {
  const headers: Record<string, string> = { Accept: "application/json" };
  if (bearerToken) headers["Authorization"] = `Bearer ${bearerToken}`;

  for (const specPath of WELL_KNOWN_SPEC_PATHS) {
    try {
      const res = await fetch(`${baseUrl}${specPath}`, { headers });
      if (res.ok) {
        const spec = (await res.json()) as OpenApiSpec;
        if (spec?.openapi && spec.paths) return spec;
      }
    } catch {
      /* try next */
    }
  }

  throw new Error(
    `Could not find an OpenAPI spec at ${baseUrl}. Tried: ${WELL_KNOWN_SPEC_PATHS.join(", ")}`,
  );
}

export async function fetchOpenApiSpecFromUrl(
  url: string,
  bearerToken?: string,
): Promise<OpenApiSpec> {
  const headers: Record<string, string> = { Accept: "application/json" };
  if (bearerToken) headers["Authorization"] = `Bearer ${bearerToken}`;

  const res = await fetch(url, { headers });
  if (!res.ok) {
    throw new Error(`Failed to fetch OpenAPI spec: ${res.status} ${res.statusText}`);
  }
  const spec = (await res.json()) as OpenApiSpec;
  if (!spec?.paths) {
    throw new Error("Invalid OpenAPI spec: missing 'paths' property");
  }
  return spec;
}

/* ------------------------------------------------------------------ */
/*  Function-name generation                                          */
/* ------------------------------------------------------------------ */

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

/** Strip FastAPI auto-generated suffixes like `_v1_sources_get` from operationIds. */
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

/** Remove a leading /vN segment when the configured API base URL already includes it. */
export function normalizeApiPath(path: string, apiBaseUrl: string): string {
  const normalizedPath = path.startsWith("/") ? path : `/${path}`;
  const base = apiBaseUrl.replace(/\/+$/, "");
  const versionMatch = base.match(/\/v(\d+)$/i);
  if (!versionMatch) return normalizedPath;

  const versionPrefix = `/v${versionMatch[1]}`;
  if (normalizedPath === versionPrefix) return "/";
  if (normalizedPath.startsWith(`${versionPrefix}/`)) {
    return normalizedPath.slice(versionPrefix.length) || "/";
  }
  return normalizedPath;
}

export function routeKey(method: string, path: string): string {
  const normalizedPath = path.replace(/^\/v\d+(?=\/|$)/i, "") || "/";
  return `${method.toUpperCase()}:${normalizedPath}`;
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

/* ------------------------------------------------------------------ */
/*  Spec → ApiEndpoint[]                                              */
/* ------------------------------------------------------------------ */

const HTTP_METHODS = ["get", "post", "put", "delete", "patch"] as const;

function extractRequiredScopes(
  operation: OpenApiOperation,
  globalSecurity: OpenApiSecurityRequirement[] | undefined,
): string[] {
  const secReqs = operation.security ?? globalSecurity ?? [];
  const scopes = new Set<string>();
  for (const req of secReqs) {
    for (const list of Object.values(req)) {
      for (const s of list) scopes.add(s);
    }
  }
  return Array.from(scopes);
}

export function parseEndpoints(spec: OpenApiSpec): ApiEndpoint[] {
  const endpoints: ApiEndpoint[] = [];
  const usedIds = new Set<string>();

  for (const [path, pathItem] of Object.entries(spec.paths)) {
    if (!pathItem) continue;
    const pathLevelParams = pathItem.parameters ?? [];

    for (const method of HTTP_METHODS) {
      const op = pathItem[method];
      if (!op || op.deprecated) continue;

      const allParams = [...pathLevelParams, ...(op.parameters ?? [])];
      const pathParams = allParams.filter((p) => p.in === "path" && !p.deprecated);
      const queryParams = allParams.filter((p) => p.in === "query" && !p.deprecated);

      let funcName = generateFunctionName(method, path, op.operationId);

      let suffix = 2;
      const base = funcName;
      while (usedIds.has(funcName.toUpperCase())) {
        funcName = `${base}${suffix++}`;
      }
      usedIds.add(funcName.toUpperCase());

      endpoints.push({
        method: method.toUpperCase() as ApiEndpoint["method"],
        path,
        operationId: op.operationId,
        summary: op.summary ?? "",
        description: op.description ?? op.summary ?? `${method.toUpperCase()} ${path}`,
        pathParameters: pathParams,
        queryParameters: queryParams,
        requiredScopes: extractRequiredScopes(op, spec.security),
        tags: op.tags ?? [],
        hasRequestBody: !!op.requestBody,
        deprecated: !!op.deprecated,
        functionName: funcName,
        functionId: funcName.toUpperCase(),
      });
    }
  }

  return endpoints;
}

/* ------------------------------------------------------------------ */
/*  Permission filtering                                              */
/* ------------------------------------------------------------------ */

export function filterEndpointsByScopes(
  endpoints: ApiEndpoint[],
  userScopes: string[],
): ApiEndpoint[] {
  const scopeSet = new Set(userScopes);
  return endpoints.filter(
    (ep) =>
      ep.requiredScopes.length === 0 ||
      ep.requiredScopes.every((s) => scopeSet.has(s)),
  );
}

export function extractAllDefinedScopes(spec: OpenApiSpec): Record<string, string> {
  const result: Record<string, string> = {};
  const schemes = spec.components?.securitySchemes;
  if (!schemes) return result;
  for (const scheme of Object.values(schemes)) {
    if (scheme.flows) {
      for (const flow of Object.values(scheme.flows)) {
        if (flow.scopes) Object.assign(result, flow.scopes);
      }
    }
  }
  return result;
}
