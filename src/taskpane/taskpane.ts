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

Office.onReady(async () => {
  try {
    await initAuth();
  } catch (err) {
    console.warn("Auth initialization warning:", err);
  }

  renderApp();
  onAuthChange(() => renderApp());
});

function renderApp(): void {
  const root = document.getElementById("app");
  if (!root) return;

  const state = getAuthState();
  const config = getConfig();

  root.innerHTML = `
    <div class="taskpane">
      <header class="taskpane-header">
        <h1 class="ms-font-xl">MT Data Connector</h1>
        <p class="ms-font-m taskpane-subtitle">Excel Add-in for MT API Services</p>
      </header>

      <section class="taskpane-section">
        <h2 class="ms-font-l">Authentication</h2>
        ${renderAuthSection(state)}
      </section>

      <section class="taskpane-section">
        <h2 class="ms-font-l">Settings</h2>
        ${renderSettingsSection(config)}
      </section>

      <section class="taskpane-section">
        <h2 class="ms-font-l">Available Functions</h2>
        ${renderFunctionReference()}
      </section>

      <footer class="taskpane-footer">
        <p class="ms-font-s">MT Data Connector v1.0.0</p>
      </footer>
    </div>
  `;

  attachEventHandlers();
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
        placeholder="https://authentik.example.com"
      />

      <button id="btn-save-settings" class="btn btn--primary">Save Settings</button>
    </div>
  `;
}

function renderFunctionReference(): string {
  const functions = [
    {
      name: "=MT.MTGETSOURCES",
      args: '([filter])',
      desc: "List available data sources.",
    },
    {
      name: "=MT.MTGETRECORDS",
      args: '(source, [limit], [offset], [filter])',
      desc: "Fetch records from a source.",
    },
    {
      name: "=MT.MTGETRECORD",
      args: '(source, recordId)',
      desc: "Look up a single record by ID.",
    },
    {
      name: "=MT.MTGETSCHEMA",
      args: '(source)',
      desc: "Get schema for a data source.",
    },
    {
      name: "=MT.MTSEARCH",
      args: '(query, [source], [limit])',
      desc: "Search across data sources.",
    },
    {
      name: "=MT.MTGETSUMMARY",
      args: '(source, [metric], [field], [filter])',
      desc: "Aggregate statistics for a source.",
    },
    {
      name: "=MT.MTSTATUS",
      args: '()',
      desc: "Check connection and auth status.",
    },
    {
      name: "=MT.MTAPICALL",
      args: '(path, [p1Name], [p1Val], ...)',
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

  btnSignIn?.addEventListener("click", async () => {
    btnSignIn.setAttribute("disabled", "true");
    btnSignIn.textContent = "Signing in…";
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

    if (apiInput?.value) {
      setApiBaseUrl(apiInput.value);
    }
    if (authentikInput?.value) {
      setAuthentikBaseUrl(authentikInput.value);
    }

    alert("Settings saved. Functions will use the new endpoint on next recalculation.");
    renderApp();
  });
}
