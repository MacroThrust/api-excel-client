/**
 * Taskpane entry point.
 *
 * Renders the side panel UI with:
 *  - Authentication status & sign-in/out button
 *  - API endpoint configuration
 *  - Quick-reference list of available mt* custom functions
 *
 * This file also initializes the shared runtime's auth layer on Office.onReady.
 */

import "./taskpane.css";
import { initAuth, signIn, signOut } from "../auth/authConfig";
import { getAuthState, onAuthChange, type AuthState } from "../shared/state";
import { getConfig, setApiBaseUrl, setAuthentikBaseUrl } from "../shared/config";
import { ADDIN_VERSION, ADDIN_HOST, BUILD_TIMESTAMP } from "../shared/version";
import {
  startPeriodicCheck,
  onUpdateAvailable,
  getUpdateInfo,
  getUpdateInstructions,
  type UpdateInfo,
} from "../shared/updateChecker";
import {
  reloadFunctions,
  getRegistryState,
  onRegistryChange,
  setOpenApiSpecUrl,
  getOpenApiSpecUrl,
} from "../functions/dynamicRegistry";

Office.onReady(async () => {
  try {
    await initAuth();
  } catch (err) {
    console.warn("Auth initialization warning:", err);
  }

  renderApp();
  onAuthChange(async (state) => {
    renderApp();
    if (state.isAuthenticated) {
      try {
        await reloadFunctions();
      } catch (err) {
        console.warn("Auto-reload of dynamic functions failed:", err);
      }
      renderApp();
    }
  });
  onRegistryChange(() => renderApp());

  startPeriodicCheck();
  onUpdateAvailable(() => renderApp());
});

function renderApp(): void {
  const root = document.getElementById("app");
  if (!root) return;

  const state = getAuthState();
  const config = getConfig();
  const updateInfo = getUpdateInfo();
  const registryState = getRegistryState();

  root.innerHTML = `
    <div class="taskpane">
      <header class="taskpane-header">
        <h1 class="ms-font-xl">MT Data Connector</h1>
        <p class="ms-font-m taskpane-subtitle">Excel Add-in for MT API Services</p>
      </header>

      ${renderUpdateBanner(updateInfo)}

      <section class="taskpane-section">
        <h2 class="ms-font-l">Authentication</h2>
        ${renderAuthSection(state)}
      </section>

      <section class="taskpane-section">
        <h2 class="ms-font-l">Settings</h2>
        ${renderSettingsSection(config)}
      </section>

      <section class="taskpane-section">
        <h2 class="ms-font-l">API Functions (OpenAPI)</h2>
        ${renderDynamicFunctionsSection(state, registryState)}
      </section>

      <section class="taskpane-section">
        <h2 class="ms-font-l">Built-in Functions</h2>
        ${renderFunctionReference()}
      </section>

      <footer class="taskpane-footer">
        <p class="ms-font-s">
          <a href="${ADDIN_HOST}/docs/user-guide.html" target="_blank" rel="noopener" class="taskpane-link">User guide</a>
          · MT Data Connector v${ADDIN_VERSION}
        </p>
        <p class="ms-font-xs taskpane-build">Built ${BUILD_TIMESTAMP}</p>
      </footer>
    </div>
  `;

  attachEventHandlers();
}

function renderUpdateBanner(info: Readonly<UpdateInfo>): string {
  if (!info.available || !info.latestVersion) return "";

  const releaseLink = info.releaseNotesUrl
    ? `<a href="${info.releaseNotesUrl}" target="_blank" rel="noopener" class="update-banner-link">Release notes</a>`
    : "";

  return `
    <div class="update-banner" role="alert">
      <div class="update-banner-icon">&#x26A0;</div>
      <div class="update-banner-content">
        <p class="ms-font-m-plus update-banner-title">Update Available</p>
        <p class="ms-font-s">
          v${info.latestVersion} is available (you have v${info.currentVersion})
        </p>
        ${releaseLink}
        <button id="btn-update-details" class="btn btn--update">How to Update</button>
      </div>
      <button id="btn-dismiss-update" class="update-banner-dismiss" aria-label="Dismiss">&times;</button>
    </div>
  `;
}

function renderAuthSection(state: Readonly<AuthState>): string {
  if (state.isAuthenticated) {
    return `
      <div class="auth-status auth-status--connected">
        <div class="status-indicator status-indicator--green"></div>
        <div>
          <p class="ms-font-m-plus"><strong>${state.userDisplayName ?? "Connected"}</strong></p>
          <p class="ms-font-s">${state.userEmail ?? ""}</p>
        </div>
      </div>
      <button id="btn-signout" class="btn btn--secondary">Sign Out</button>
    `;
  }

  return `
    <div class="auth-status auth-status--disconnected">
      <div class="status-indicator status-indicator--red"></div>
      <p class="ms-font-m">Not signed in</p>
    </div>
    <button id="btn-signin" class="btn btn--primary">Sign In with Microsoft</button>
  `;
}

function renderSettingsSection(config: ReturnType<typeof getConfig>): string {
  const specUrl = getOpenApiSpecUrl() ?? "";
  return `
    <div class="settings-form">
      <label class="ms-font-s" for="input-api-url">API Base URL</label>
      <input
        id="input-api-url"
        type="url"
        class="input"
        value="${config.apiBaseUrl}"
        placeholder="https://api.example.com/v1"
      />

      <label class="ms-font-s" for="input-authentik-url">Authentik Base URL</label>
      <input
        id="input-authentik-url"
        type="url"
        class="input"
        value="${config.authentikBaseUrl}"
        placeholder="https://auth.macrothrust.com"
      />

      <label class="ms-font-s" for="input-openapi-url">OpenAPI Spec URL <span class="ms-font-xs">(optional — auto-discovered if blank)</span></label>
      <input
        id="input-openapi-url"
        type="url"
        class="input"
        value="${specUrl}"
        placeholder="https://api.example.com/openapi.json"
      />

      <button id="btn-save-settings" class="btn btn--primary">Save Settings</button>
    </div>
  `;
}

function renderDynamicFunctionsSection(
  authState: Readonly<AuthState>,
  registryState: ReturnType<typeof getRegistryState>,
): string {
  if (!authState.isAuthenticated) {
    return `
      <p class="ms-font-s">Sign in to discover API functions from the OpenAPI specification.</p>
    `;
  }

  if (!registryState.loaded) {
    return `
      <p class="ms-font-s">No OpenAPI spec loaded yet.</p>
      <button id="btn-reload-functions" class="btn btn--primary">Load API Functions</button>
      <p class="ms-font-xs" style="margin-top:8px;">
        Or use <code>=MT.RELOADFUNCTIONS()</code> in any cell.
      </p>
    `;
  }

  if (registryState.lastError) {
    return `
      <div class="auth-status auth-status--disconnected">
        <div class="status-indicator status-indicator--red"></div>
        <p class="ms-font-s">${registryState.lastError}</p>
      </div>
      <button id="btn-reload-functions" class="btn btn--primary">Retry</button>
    `;
  }

  const permitted = registryState.functions.filter((f) => f.registered);
  const denied = registryState.functions.filter((f) => !f.registered);

  let html = `
    <div class="auth-status auth-status--connected">
      <div class="status-indicator status-indicator--green"></div>
      <div>
        <p class="ms-font-m-plus"><strong>${registryState.registeredCount}</strong> of ${registryState.endpointCount} endpoints available</p>
        <p class="ms-font-xs">Last loaded: ${registryState.lastLoadTime ? new Date(registryState.lastLoadTime).toLocaleTimeString() : "never"}</p>
      </div>
    </div>
    <button id="btn-reload-functions" class="btn btn--secondary" style="margin-bottom:12px;">Reload Functions</button>
  `;

  if (permitted.length > 0) {
    html += `<div class="function-list">`;
    for (const f of permitted) {
      const params = [
        ...f.endpoint.pathParameters.map((p) => p.required !== false ? p.name : `[${p.name}]`),
        ...f.endpoint.queryParameters.map((p) => `[${p.name}]`),
        ...(f.endpoint.hasRequestBody ? ["[jsonBody]"] : []),
      ].join(", ");
      html += `
        <div class="function-card">
          <code class="function-name">=MT.${f.endpoint.functionId}(${params})</code>
          <p class="ms-font-xs function-desc"><span class="method-badge method-badge--${f.endpoint.method.toLowerCase()}">${f.endpoint.method}</span> ${f.endpoint.path}</p>
          <p class="ms-font-s function-desc">${f.endpoint.summary || f.endpoint.description}</p>
        </div>
      `;
    }
    html += `</div>`;
  }

  if (denied.length > 0) {
    html += `
      <details style="margin-top:8px;">
        <summary class="ms-font-s">${denied.length} endpoint(s) not permitted</summary>
        <ul class="ms-font-xs" style="margin:4px 0 0 16px; color: #888;">
          ${denied.map((f) => `<li>${f.endpoint.method} ${f.endpoint.path} — requires: ${f.endpoint.requiredScopes.join(", ")}</li>`).join("")}
        </ul>
      </details>
    `;
  }

  html += `
    <p class="ms-font-xs" style="margin-top:8px;">
      Use <code>=MT.LISTENDPOINTS()</code> to see all endpoints in a cell.
    </p>
  `;

  return html;
}

function renderFunctionReference(): string {
  const functions = [
    {
      name: "=MT.GETSOURCES",
      args: "([offset], [limit])",
      desc: "List data sources (GET /sources).",
    },
    {
      name: "=MT.GETSOURCE",
      args: "(sourceId)",
      desc: "Look up one data source (GET /sources/{id}).",
    },
    {
      name: "=MT.GETSOURCEDATASETS",
      args: "(sourceId, [limit], [offset], [idsOnly])",
      desc: "List datasets in a source (GET /sources/{id}/datasets).",
    },
    {
      name: "=MT.GETDATASETS",
      args: "([sourceId], [offset], [limit])",
      desc: "List datasets (GET /datasets).",
    },
    {
      name: "=MT.GETDATASET",
      args: "(datasetId)",
      desc: "Look up one dataset (GET /datasets/{id}).",
    },
    {
      name: "=MT.GETDATASETSERIES",
      args: "(datasetId, [limit], [offset], [idsOnly])",
      desc: "List series in a dataset (GET /datasets/{id}/series).",
    },
    {
      name: "=MT.GETSERIES",
      args: "(seriesId)",
      desc: "Look up one series (GET /series/{id}).",
    },
    {
      name: "=MT.GETOBSERVATIONS",
      args: "(seriesId, [startDate], [endDate], [limit], [offset])",
      desc: "Fetch series observations (GET /series/{id}/observations).",
    },
    {
      name: "=MT.STATUS",
      args: "()",
      desc: "Check API health and auth (GET /health).",
    },
    {
      name: "=MT.VERSION",
      args: "()",
      desc: "Returns add-in version and build info.",
    },
    {
      name: "=MT.APICALL",
      args: "(path, [p1Name], [p1Val], ...)",
      desc: "Generic API call to any endpoint.",
    },
  ];

  return `
    <div class="function-list">
      ${functions
        .map(
          (f) => `
        <div class="function-card">
          <code class="function-name">${f.name}${f.args}</code>
          <p class="ms-font-s function-desc">${f.desc}</p>
        </div>
      `
        )
        .join("")}
    </div>
  `;
}

function attachEventHandlers(): void {
  const btnSignIn = document.getElementById("btn-signin");
  const btnSignOut = document.getElementById("btn-signout");
  const btnSave = document.getElementById("btn-save-settings");
  const btnUpdateDetails = document.getElementById("btn-update-details");
  const btnDismissUpdate = document.getElementById("btn-dismiss-update");
  const btnReloadFunctions = document.getElementById("btn-reload-functions");

  btnSignIn?.addEventListener("click", async () => {
    const btn = btnSignIn as HTMLButtonElement;
    btn.disabled = true;
    btn.textContent = "Opening sign-in…";
    try {
      await signIn();
    } catch (err) {
      console.error("Sign in failed:", err);
      alert(`Sign in failed: ${err instanceof Error ? err.message : "Unknown error"}`);
    } finally {
      renderApp();
    }
  });

  btnSignOut?.addEventListener("click", async () => {
    try {
      await signOut();
    } catch (err) {
      console.error("Sign out failed:", err);
    } finally {
      renderApp();
    }
  });

  btnSave?.addEventListener("click", () => {
    const apiInput = document.getElementById("input-api-url") as HTMLInputElement;
    const authentikInput = document.getElementById("input-authentik-url") as HTMLInputElement;
    const openapiInput = document.getElementById("input-openapi-url") as HTMLInputElement;

    if (apiInput?.value) {
      setApiBaseUrl(apiInput.value);
    }
    if (authentikInput?.value) {
      setAuthentikBaseUrl(authentikInput.value);
    }
    if (openapiInput) {
      setOpenApiSpecUrl(openapiInput.value);
    }

    alert("Settings saved. Functions will use the new endpoint on next recalculation.");
    renderApp();
  });

  btnUpdateDetails?.addEventListener("click", () => {
    const info = getUpdateInfo();
    if (info.latestVersion) {
      alert(getUpdateInstructions(info.latestVersion));
    }
  });

  btnDismissUpdate?.addEventListener("click", () => {
    const banner = document.querySelector(".update-banner");
    if (banner) {
      (banner as HTMLElement).style.display = "none";
    }
  });

  btnReloadFunctions?.addEventListener("click", async () => {
    btnReloadFunctions.setAttribute("disabled", "true");
    btnReloadFunctions.textContent = "Loading…";
    try {
      const result = await reloadFunctions();
      if (result.error) {
        alert(`Failed to load functions: ${result.error}`);
      }
    } catch (err) {
      alert(`Error: ${err instanceof Error ? err.message : "Unknown error"}`);
    } finally {
      renderApp();
    }
  });
}
