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

function capitalize(s: string): string {
  return s.charAt(0).toUpperCase() + s.slice(1);
}

export function generateFunctionName(
  method: string,
  path: string,
  operationId?: string,
): string {
  const verb = VERB_MAP[method.toLowerCase()] ?? capitalize(method.toLowerCase());

  if (operationId) {
    const cleanId = operationId
      .replace(/[^a-zA-Z0-9_]/g, "")
      .replace(/^[a-z]/, (c) => c.toUpperCase());
    return `${verb}${cleanId}`;
  }

  const pathPart = path
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
