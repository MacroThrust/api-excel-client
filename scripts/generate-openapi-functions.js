#!/usr/bin/env node
/**
 * Build-time OpenAPI → functions.json metadata generator.
 *
 * Reads an OpenAPI 3.x spec (from a local file or URL) and emits function
 * metadata entries that can be merged into the functions.json produced by
 * custom-functions-metadata-plugin. This allows ALL possible dynamic
 * functions to appear in Excel's autocomplete/wizard, even before the
 * user signs in — the runtime registry then selectively associates only
 * the permitted subset.
 *
 * Usage:
 *   node scripts/generate-openapi-functions.js <spec-path-or-url> [--output dist/functions.json] [--merge]
 *
 * Flags:
 *   --output <path>   Where to write (default: stdout as JSON)
 *   --merge           If set, reads existing functions.json from --output and
 *                     appends new entries (preserving hand-coded functions).
 */

const fs = require("fs");
const path = require("path");
const https = require("https");
const http = require("http");

/* ------------------------------------------------------------------ */
/*  CLI argument parsing                                               */
/* ------------------------------------------------------------------ */
const args = process.argv.slice(2);
let specSource = null;
let outputPath = null;
let merge = false;

for (let i = 0; i < args.length; i++) {
  if (args[i] === "--output" && args[i + 1]) {
    outputPath = args[++i];
  } else if (args[i] === "--merge") {
    merge = true;
  } else if (!specSource) {
    specSource = args[i];
  }
}

if (!specSource) {
  console.error("Usage: generate-openapi-functions <spec.json | URL> [--output path] [--merge]");
  process.exit(1);
}

/* ------------------------------------------------------------------ */
/*  OpenAPI parsing (mirrors src/shared/openApiClient.ts logic)       */
/* ------------------------------------------------------------------ */

const VERB_MAP = { get: "Get", post: "Post", put: "Put", delete: "Delete", patch: "Patch" };
const HTTP_METHODS = ["get", "post", "put", "delete", "patch"];

const BUILTIN_FUNCTION_BY_ROUTE = {
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

function capitalize(s) {
  return s.charAt(0).toUpperCase() + s.slice(1);
}

function snakeToPascalCase(value) {
  return value
    .split("_")
    .filter(Boolean)
    .map(capitalize)
    .join("");
}

function normalizeOperationId(operationId) {
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

function routeKey(method, pathStr) {
  const normalizedPath = pathStr.replace(/^\/v\d+(?=\/|$)/i, "") || "/";
  return `${method.toUpperCase()}:${normalizedPath}`;
}

function getBuiltinFunctionId(method, pathStr) {
  return BUILTIN_FUNCTION_BY_ROUTE[routeKey(method, pathStr)];
}

function operationIdToFunctionName(operationId, method) {
  const normalized = normalizeOperationId(operationId);
  const verb = VERB_MAP[method] || capitalize(method);

  const verbPrefixes = {
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

function generateFunctionName(method, pathStr, operationId) {
  const builtinId = getBuiltinFunctionId(method, pathStr);
  if (builtinId) return builtinId;

  const verb = VERB_MAP[method] || capitalize(method);
  if (operationId) {
    return operationIdToFunctionName(operationId, method);
  }

  const pathPart = pathStr
    .replace(/^\/v\d+(?=\/|$)/i, "")
    .split("/")
    .filter((s) => s && !s.startsWith("{"))
    .map((s) => s.replace(/[^a-zA-Z0-9]/g, ""))
    .map(capitalize)
    .join("");

  return `${verb}${pathPart}`;
}

function mapSchemaToExcelType(schema) {
  if (!schema) return "string";
  switch (schema.type) {
    case "integer":
    case "number":
      return "number";
    case "boolean":
      return "boolean";
    default:
      return "string";
  }
}

function parseSpec(spec) {
  const functions = [];
  const usedIds = new Set();

  for (const [pathStr, pathItem] of Object.entries(spec.paths || {})) {
    if (!pathItem) continue;
    const pathLevelParams = pathItem.parameters || [];

    for (const method of HTTP_METHODS) {
      const op = pathItem[method];
      if (!op || op.deprecated) continue;

      let funcName = generateFunctionName(method, pathStr, op.operationId);
      if (getBuiltinFunctionId(method, pathStr)) {
        continue;
      }
      let suffix = 2;
      const base = funcName;
      while (usedIds.has(funcName.toUpperCase())) {
        funcName = `${base}${suffix++}`;
      }
      const funcId = funcName.toUpperCase();
      usedIds.add(funcId);

      const allParams = [...pathLevelParams, ...(op.parameters || [])];
      const pathParams = allParams.filter((p) => p.in === "path" && !p.deprecated);
      const queryParams = allParams.filter((p) => p.in === "query" && !p.deprecated);

      const parameters = [];

      for (const pp of pathParams) {
        parameters.push({
          name: pp.name,
          description: pp.description || `Path parameter: ${pp.name}`,
          type: mapSchemaToExcelType(pp.schema),
          optional: pp.required === false,
        });
      }

      for (const qp of queryParams) {
        parameters.push({
          name: qp.name,
          description: qp.description || `Query parameter: ${qp.name}`,
          type: mapSchemaToExcelType(qp.schema),
          optional: true,
        });
      }

      if (op.requestBody) {
        parameters.push({
          name: "jsonBody",
          description: "Request body as JSON string",
          type: "string",
          optional: !op.requestBody.required,
        });
      }

      const description = op.summary || op.description || `${method.toUpperCase()} ${pathStr}`;

      functions.push({
        id: funcId,
        name: funcName,
        description,
        helpUrl: "",
        excludeFromAutoComplete: true,
        result: { type: "string", dimensionality: "matrix" },
        parameters,
        options: { cancelable: true, requiresAddress: false },
      });
    }
  }

  return functions;
}

/* ------------------------------------------------------------------ */
/*  Fetch helper (supports file:// and https:// URLs)                 */
/* ------------------------------------------------------------------ */

function fetchSpec(source) {
  return new Promise((resolve, reject) => {
    if (source.startsWith("http://") || source.startsWith("https://")) {
      const mod = source.startsWith("https") ? https : http;
      mod.get(source, { headers: { Accept: "application/json" } }, (res) => {
        let body = "";
        res.on("data", (chunk) => (body += chunk));
        res.on("end", () => {
          try {
            resolve(JSON.parse(body));
          } catch (e) {
            reject(new Error(`Invalid JSON from ${source}: ${e.message}`));
          }
        });
      }).on("error", reject);
    } else {
      try {
        const content = fs.readFileSync(path.resolve(source), "utf8");
        resolve(JSON.parse(content));
      } catch (e) {
        reject(new Error(`Failed to read ${source}: ${e.message}`));
      }
    }
  });
}

/* ------------------------------------------------------------------ */
/*  Main                                                              */
/* ------------------------------------------------------------------ */

async function main() {
  const spec = await fetchSpec(specSource);

  if (!spec.paths) {
    console.error("Error: spec has no 'paths' property — is this a valid OpenAPI spec?");
    process.exit(1);
  }

  const generated = parseSpec(spec);

  console.error(`[generate-openapi-functions] Generated ${generated.length} function(s) from ${Object.keys(spec.paths).length} path(s)`);

  if (outputPath && merge) {
    let existing = { functions: [] };
    try {
      const raw = fs.readFileSync(outputPath, "utf8");
      existing = JSON.parse(raw);
    } catch {
      /* file doesn't exist yet — start fresh */
    }

    const existingIds = new Set(existing.functions.map((f) => f.id));
    const newFunctions = generated.filter((f) => !existingIds.has(f.id));

    existing.functions.push(...newFunctions);
    fs.writeFileSync(outputPath, JSON.stringify(existing, null, 2));
    console.error(`[generate-openapi-functions] Merged ${newFunctions.length} new function(s) into ${outputPath} (total: ${existing.functions.length})`);
  } else if (outputPath) {
    const output = { functions: generated };
    fs.writeFileSync(outputPath, JSON.stringify(output, null, 2));
    console.error(`[generate-openapi-functions] Wrote ${generated.length} function(s) to ${outputPath}`);
  } else {
    process.stdout.write(JSON.stringify({ functions: generated }, null, 2));
  }
}

main().catch((err) => {
  console.error("Error:", err.message);
  process.exit(1);
});
