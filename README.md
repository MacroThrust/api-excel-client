# MT Data Connector — Excel Add-in

An Excel Office Add-in that connects to a third-party API protected by [Authentik](https://goauthentik.io/) OAuth2, using Microsoft Account authentication. Designed for Office 365 Online (Excel on the web) with a shared runtime architecture.

The add-in is hosted as a static site on **GitHub Pages** and deployed automatically via **GitHub Actions** on every push to `main`.

## Documentation

| Audience | Where to read |
|----------|----------------|
| **End users** (install, sign in, formulas) | [docs/user-guide.md](docs/user-guide.md) — HTML is generated at build time and published at `https://<owner>.github.io/<repo>/docs/user-guide.html` |
| **Developers / admins** (build, config, CI) | This README |

The user guide is copied into `dist/docs/` on every build, so it ships on GitHub Pages next to `manifest.xml` and the add-in bundles. The task pane includes a **User guide** link to that URL.

## Architecture

```
┌─────────────────────────────────────────────────────────┐
│  Excel (Office 365 Online / Desktop)                    │
│                                                         │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐  │
│  │  Taskpane UI  │  │ Custom Funcs │  │ Ribbon Menu  │  │
│  │  (Settings,   │  │ (MT.* UDFs)  │  │ (Sign In,    │  │
│  │   Auth, Ref)  │  │              │  │  Refresh)    │  │
│  └──────┬───────┘  └──────┬───────┘  └──────┬───────┘  │
│         │                 │                  │          │
│         └────────┬────────┴──────────────────┘          │
│                  │  Shared Runtime                       │
│         ┌────────┴────────┐                             │
│         │   Auth State    │                             │
│         │   API Client    │                             │
│         │   Config        │                             │
│         │  Update Checker │                             │
│         └────────┬────────┘                             │
└──────────────────┼──────────────────────────────────────┘
                   │ HTTPS
         ┌─────────┴─────────┐
         │    Authentik       │
         │  (OAuth2 / OIDC)   │
         └─────────┬─────────┘
                   │
         ┌─────────┴─────────┐
         │   Your API         │
         │  (Protected)       │
         └───────────────────┘

Hosted on GitHub Pages:
  https://<owner>.github.io/<repo>/
    ├── taskpane.html          ← Shared runtime entry
    ├── taskpane.*.js          ← Taskpane + functions + commands bundle
    ├── functions.*.js         ← Custom functions bundle
    ├── functions.json         ← Custom function metadata (auto-generated)
    ├── auth-dialog.html       ← OAuth2 dialog popup
    ├── auth-dialog.*.js       ← Auth dialog bundle
    ├── version.json           ← Version manifest (for update detection)
    ├── manifest.xml           ← Office Add-in manifest (URLs rewritten at build)
    ├── docs/                  ← User guide (HTML + markdown)
    └── assets/                ← Icons
```

## Features

- **Microsoft Account SSO** via MSAL.js (Nested App Authentication with Dialog API fallback)
- **Authentik OAuth2 integration** — token exchange or authorization code flow with PKCE
- **Dynamic OpenAPI-driven functions** — auto-generates Excel custom functions from an API's OpenAPI spec, filtered by user/client OAuth2 scopes
- **Permission-based function visibility** — only endpoints the user's token permits are registered; unauthorized endpoints return `#NAME?`
- **Catalog-aligned custom functions** — built-in UDFs map to MacroThrust API drill-down: sources → datasets → series → observations
- **Shared runtime** — auth tokens are shared between taskpane, custom functions, and ribbon commands
- **Configurable API endpoint** — change the target API URL at runtime via the Settings panel
- **Ribbon menu** with Sign In, Sign Out, Settings, Refresh Data, and Reload API Functions commands
- **Automatic update notifications** — detects new deployments and shows user instructions
- **CI/CD via GitHub Actions** — builds and deploys to GitHub Pages on every push to `main`

## Custom Functions

All functions are available under the `MT` namespace in Excel:

### Built-in Functions

| Function | Description |
|---|---|
| `=MT.GETSOURCES([offset], [limit])` | List data sources (`GET /sources`) |
| `=MT.GETSOURCE(sourceId)` | Look up one data source (`GET /sources/{id}`) |
| `=MT.GETSOURCEDATASETS(sourceId, [limit], [offset], [idsOnly])` | Datasets under a source (`GET /sources/{id}/datasets`) |
| `=MT.GETDATASETS([sourceId], [offset], [limit])` | List datasets (`GET /datasets`) |
| `=MT.GETDATASET(datasetId)` | Look up one dataset (`GET /datasets/{id}`) |
| `=MT.GETDATASETSERIES(datasetId, [limit], [offset], [idsOnly])` | Series under a dataset (`GET /datasets/{id}/series`) |
| `=MT.GETSERIES(seriesId)` | Look up one series (`GET /series/{id}`) |
| `=MT.GETOBSERVATIONS(seriesId, [startDate], [endDate], [limit], [offset])` | Time-series observations (`GET /series/{id}/observations`) |
| `=MT.STATUS()` | Check API health and auth (`GET /health`) |
| `=MT.VERSION()` | Returns embedded add-in name, version, and build timestamp |
| `=MT.APICALL(path, [p1Name], [p1Val], [p2Name], [p2Val], [p3Name], [p3Val])` | Generic API call |
| `=MT.RELOADFUNCTIONS()` | Reload dynamic functions from OpenAPI spec |
| `=MT.LISTENDPOINTS()` | List all discovered endpoints and permissions |

All functions are async, support dynamic array spilling, use cancelable invocations, and return inline error messages instead of `#VALUE!`.

### Dynamic OpenAPI Functions

After signing in, the add-in can automatically generate custom functions from the API's **OpenAPI 3.x specification**. These functions are permission-aware: only endpoints that the user's OAuth2 token scopes allow will be registered.

**How it works:**

1. **Sign in** — the add-in acquires an OAuth2 token with specific scopes
2. **Reload** — call `=MT.RELOADFUNCTIONS()`, click "Reload API Functions" in the ribbon, or click the button in the taskpane (also triggers automatically after sign-in)
3. The add-in fetches the OpenAPI spec from the API (auto-discovers `/openapi.json`, `/swagger.json`, etc., or uses a configured URL)
4. Each endpoint in the spec is mapped to a function named with a **verb prefix**:
   - `GET /users` → `=MT.GETUSERS()`
   - `POST /users` → `=MT.POSTUSERS(jsonBody)`
   - `DELETE /users/{id}` → `=MT.DELETEUSERS(id)`
   - If the endpoint has an `operationId`, it is used: `GET /users` with `operationId: listUsers` → `=MT.GETLISTUSERS()`
5. The spec's `security` requirements are compared against the user's token scopes — only matching endpoints become callable
6. **Denied endpoints are removed from the Excel UI** — on Excel builds that support API 1.20+, `Excel.CustomFunctionManager.setVisibility()` hides unauthorized functions from autocomplete and the Formula Builder entirely. On older Excel versions, denied functions remain hidden from autocomplete (via `excludeFromAutoComplete` in metadata) and return an "Access denied" error if typed manually

**Function naming convention:**

| HTTP Method | Prefix | Example |
|---|---|---|
| GET | `get...` | `=MT.GETUSERS()` |
| POST | `post...` | `=MT.POSTUSERS(jsonBody)` |
| PUT | `put...` | `=MT.PUTUSERS(id, jsonBody)` |
| DELETE | `delete...` | `=MT.DELETEUSERS(id)` |
| PATCH | `patch...` | `=MT.PATCHUSERS(id, jsonBody)` |

**Parameters** are derived from the OpenAPI spec:
- Path parameters become required function arguments
- Query parameters become optional function arguments
- Request bodies become an optional `jsonBody` string argument (pass JSON)

**OpenAPI spec discovery:** The add-in tries these paths in order: `/openapi.json`, `/api/openapi.json`, `/swagger.json`, `/api/swagger.json`, `/docs/openapi.json`, `/api-docs`, `/v1/openapi.json`. You can also configure an explicit URL in the Settings panel.

### Build-time Pre-generation (Optional)

For functions to appear in Excel's autocomplete/Insert Function wizard, they must be declared in `functions.json` at build time. To pre-generate metadata for all possible endpoints:

```bash
# From a local OpenAPI spec file:
node scripts/generate-openapi-functions.js openapi.json --output dist/functions.json --merge

# From a URL:
node scripts/generate-openapi-functions.js https://api.example.com/openapi.json --output dist/functions.json --merge
```

The `--merge` flag preserves the built-in functions already in `functions.json` and adds the OpenAPI-derived ones. Without pre-generation, dynamic functions still work but won't appear in Excel's autocomplete until invoked.

---

## Deploying to GitHub Pages

This section explains how to serve the add-in from GitHub Pages and what the CI pipeline does.

### What Gets Deployed

GitHub Pages serves the contents of the `dist/` folder produced by `npm run build`. The build is parameterized by the `ADDIN_HOST` environment variable, which controls all URLs inside `manifest.xml` and in the runtime configuration. The CI pipeline sets this automatically to `https://<owner>.github.io/<repo>/`.

The deployed site contains:

| File | Purpose |
|---|---|
| `taskpane.html` | Shared runtime HTML — loads all JavaScript bundles |
| `taskpane.*.js` | Main bundle (taskpane UI + custom functions + ribbon commands) |
| `functions.*.js` | Custom functions bundle |
| `functions.json` | Auto-generated function metadata for Excel registration |
| `auth-dialog.html` | OAuth2 popup dialog HTML |
| `auth-dialog.*.js` | Auth dialog bundle |
| `version.json` | Version manifest for automatic update detection |
| `manifest.xml` | Office Add-in manifest with all URLs pointing to GitHub Pages |
| `docs/user-guide.html` | End-user guide HTML (generated from markdown at build time) |
| `docs/user-guide.md` | User guide source (edit this file) |
| `assets/` | Add-in icons (16, 32, 64, 80px) |

### Prerequisites

1. A GitHub repository with this code on the `main` branch
2. GitHub Pages enabled for the repository (see setup steps below)
3. The following configuration values updated in the source (see [Configuration](#configuration))

### One-Time GitHub Pages Setup

1. Go to your repository on GitHub
2. Navigate to **Settings** > **Pages**
3. Under **Build and deployment**, set **Source** to **GitHub Actions**
4. That's it — the workflow will deploy on the next push to `main`

The deployed site will be available at:

```
https://<your-github-username>.github.io/<your-repo-name>/
```

For example, if the repo is `MacroThrust/excel-api-client`, the URL would be:

```
https://macrothrust.github.io/excel-api-client/
```

### How the CI Pipeline Works

The workflow file is at `.github/workflows/deploy.yml`. Here is what happens on every push to `main` (or manual trigger via `workflow_dispatch`):

```
Push to main
     │
     ▼
┌─────────────────────────────────────────────────┐
│  Job: build                                      │
│                                                  │
│  1. Checkout repository                          │
│  2. Setup Node.js 20 with npm cache              │
│  3. npm ci (install locked dependencies)         │
│  4. npm run build                                │
│     └─ ADDIN_HOST is set to the GitHub Pages URL │
│     └─ manifest.xml URLs are rewritten           │
│     └─ version.json is emitted                   │
│  5. Upload dist/ as a Pages artifact             │
└──────────────────┬──────────────────────────────┘
                   │
                   ▼
┌─────────────────────────────────────────────────┐
│  Job: deploy                                     │
│                                                  │
│  1. Deploy artifact to GitHub Pages              │
│  2. Site is live at the Pages URL                │
└─────────────────────────────────────────────────┘
```

Key details:

- **`ADDIN_HOST`** is computed automatically from `github.repository_owner` and `github.event.repository.name`, so no manual URL configuration is needed in the workflow.
- **`manifest.xml`** is transformed at build time: every occurrence of `https://localhost:3000` is replaced with the GitHub Pages URL.
- **`version.json`** is emitted alongside the bundle so the update checker in running add-in instances can detect new deployments.
- **Concurrency** is set so only one deployment runs at a time; subsequent pushes cancel in-progress deploys.
- The workflow uses the recommended `actions/upload-pages-artifact` + `actions/deploy-pages` pattern for GitHub Pages.

### The Manifest on GitHub Pages

After a successful deployment, the manifest is available at:

```
https://<owner>.github.io/<repo>/manifest.xml
```

This is the URL you provide when sideloading or deploying the add-in (see [Installing the Add-in](#installing-the-add-in-in-excel)).

---

## Configuration

Before the add-in will work, you must update the following values in the source code. These are compile-time settings baked into the build.

### `src/shared/config.ts`

| Value | What to set |
|---|---|
| `msalClientId` | Your Azure AD application (client) ID |
| `msalAuthority` | Your Azure AD authority URL (default: `https://login.microsoftonline.com/common`) |
| `authentikClientId` | Your Authentik OAuth2 provider client ID |
| `authentikBaseUrl` (default) | Your Authentik instance URL (e.g., `https://authentik.yourcompany.com`) |
| `apiBaseUrl` (default) | Your API base URL (e.g., `https://api.yourcompany.com/v1`) |

### `manifest.xml`

| Value | What to set |
|---|---|
| `<Id>` | A unique GUID for your add-in (generate at https://www.guidgenerator.com) |
| `<ProviderName>` | MacroThrust |
| `<AppDomains>` | Replace `auth.macrothrust.com` with your Authentik domain |

The `https://localhost:3000` URLs in the manifest are **automatically rewritten** at build time using the `ADDIN_HOST` environment variable, so you do not need to change them manually.

### `.env.example`

Copy to `.env` for local development reference. These values are not used at build time — they exist as documentation for the values in `config.ts`.

---

## Installing the Add-in in Excel

> **End users:** Step-by-step install and first-run instructions are in the [user guide](docs/user-guide.md).

### Excel on the Web (Office 365 Online) — Sideloading

1. Open Excel Online at https://www.office.com/launch/excel
2. Open or create a workbook
3. Go to **Insert** > **Office Add-ins** > **Upload My Add-in**
4. Enter the manifest URL:
   ```
   https://<owner>.github.io/<repo>/manifest.xml
   ```
   Or download the manifest from that URL and upload the file directly.

### Excel on the Web — Direct Manifest URL

Some organizations allow providing the manifest URL directly. Point it at:

```
https://<owner>.github.io/<repo>/manifest.xml
```

### Excel Desktop (Windows/Mac)

1. Download `manifest.xml` from your GitHub Pages URL
2. Place it in a local folder or network share
3. In Excel, go to **File** > **Options** > **Trust Center** > **Trusted Add-in Catalogs**
4. Add the folder path as a trusted catalog
5. Restart Excel, then **Insert** > **My Add-ins** > **Shared Folder**

### Organization-wide Deployment (Microsoft 365 Admin Center)

For deploying to all users in your organization:

1. Go to the [Microsoft 365 admin center](https://admin.microsoft.com)
2. Navigate to **Settings** > **Integrated apps** > **Upload custom apps**
3. Provide the manifest URL:
   ```
   https://<owner>.github.io/<repo>/manifest.xml
   ```
4. Assign the add-in to the appropriate users or groups

---

## Automatic Update Notifications

When a new version is deployed to GitHub Pages, users running the previous version will see a notification banner in the taskpane:

- The add-in checks `version.json` on the hosting server every 30 minutes
- If the server version is newer than the version embedded in the running bundle, a dismissible warning banner appears
- The "How to Update" button shows platform-specific instructions:
  - **Excel Online**: Hard-refresh the browser (`Ctrl+Shift+R`)
  - **Excel Desktop**: Clear the Office web cache and restart
  - **Admin-deployed**: Contact IT to update the deployment

To bump the version for a new release, update the `version` field in `package.json` and push to `main`. The CI pipeline will build and deploy with the new version, and running instances will detect it automatically.

---

## Project Structure

```
├── .github/
│   └── workflows/
│       └── deploy.yml            # GitHub Actions: build + deploy to Pages
├── manifest.xml                  # Office Add-in XML manifest (source; URLs rewritten at build)
├── package.json                  # Dependencies and scripts
├── webpack.config.js             # Build configuration (DefinePlugin + manifest rewrite)
├── tsconfig.json                 # TypeScript configuration
├── .env.example                  # Environment variable reference
├── assets/                       # Add-in icons (16, 32, 64, 80 px)
├── docs/
│   ├── user-guide.md             # End-user guide (source; edit this)
│   ├── user-guide.html           # Generated by npm run build:docs (gitignored)
│   └── index.html                # Redirects to user-guide.html
└── src/
    ├── auth/
    │   ├── authConfig.ts         # MSAL init, NAA + Dialog API auth flows
    │   ├── authDialog.ts         # Dialog popup: Authentik OAuth2 + PKCE
    │   └── auth-dialog.html      # Dialog popup HTML shell
    ├── commands/
    │   └── commands.ts           # Ribbon button/menu command handlers
    ├── functions/
    │   └── functions.ts          # Custom function definitions (catalog UDFs)
    ├── shared/
    │   ├── apiClient.ts          # HTTP client with auth headers
    │   ├── config.ts             # Centralized configuration (uses build-injected host)
    │   ├── state.ts              # Shared auth state (module-level singleton)
    │   ├── updateChecker.ts      # Periodic update detection + instructions
    │   └── version.ts            # Build-embedded constants (version, name, host, timestamp)
    └── taskpane/
        ├── taskpane.html         # Taskpane HTML (shared runtime entry)
        ├── taskpane.ts           # Taskpane UI logic
        └── taskpane.css          # Taskpane styles (Fluent-inspired)
```

---

## Local Development

### Prerequisites

- **Node.js** >= 18
- **npm** >= 9

### Setup

```bash
# Install dependencies
npm install

# Start dev server (https://localhost:3000)
npm run dev
```

The dev server uses `https://localhost:3000` with a self-signed certificate. For local development, the manifest URLs point to localhost by default (no `ADDIN_HOST` override needed).

To sideload locally in Excel Online, upload `manifest.xml` (from the repo root, not `dist/`) in **Insert** > **Office Add-ins** > **Upload My Add-in**.

### Build Commands

| Command | Description |
|---|---|
| `npm run dev` | Start webpack dev server with hot reload |
| `npm run build` | Production build to `dist/` |
| `npm run build:dev` | Development build to `dist/` |
| `npm run validate` | Validate `manifest.xml` against Office schemas |

### Building for a Custom Host

To build for a host other than GitHub Pages or localhost:

```bash
ADDIN_HOST=https://your-custom-domain.com npm run build
```

This rewrites all `https://localhost:3000` URLs in `manifest.xml` and injects the host into the runtime configuration.

---

## Authentication Flow

```
User clicks "Sign In"
        │
        ├─── NAA Supported? ───► acquireTokenSilent (MSAL)
        │         │                      │
        │         │               Success │ Fail
        │         │                 │     └──► acquireTokenPopup
        │         │                 │                │
        │         │                 ▼                ▼
        │         │          MS Access Token ◄───────┘
        │         │                 │
        │         │      Exchange with Authentik
        │         │       (token exchange grant)
        │         │                 │
        │         │          Authentik Token
        │         │                 │
        │         └──────► Store in Shared State
        │
        └─── NAA Not Supported ──► Open Dialog API
                                        │
                                  Redirect to Authentik
                                  (OAuth2 + PKCE)
                                        │
                                  User authenticates
                                  (via Microsoft IdP
                                   configured in Authentik)
                                        │
                                  Authorization Code
                                        │
                                  Exchange for Token
                                        │
                                  messageParent(token)
                                        │
                                  Store in Shared State
```

---

## Releasing a New Version

1. Update the `version` field in `package.json` (e.g., `"1.0.0"` → `"1.1.0"`)
2. Commit and push to `main`
3. The GitHub Actions workflow will automatically build and deploy
4. Running add-in instances will detect the new version within 30 minutes and show an update banner

Optionally, add a `homepage` field to `package.json` to include a release notes URL in the update notification:

```json
{
  "homepage": "https://github.com/<owner>/<repo>/releases"
}
```

---

## License

Private — All rights reserved.
