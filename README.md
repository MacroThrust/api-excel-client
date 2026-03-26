# MT Data Connector — Excel Add-in

An Excel Office Add-in that connects to a third-party API protected by [Authentik](https://goauthentik.io/) OAuth2, using Microsoft Account authentication. Designed for Office 365 Online (Excel on the web) with a shared runtime architecture.

## Architecture

```
┌─────────────────────────────────────────────────────────┐
│  Excel (Office 365 Online / Desktop)                    │
│                                                         │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐  │
│  │  Taskpane UI  │  │ Custom Funcs │  │ Ribbon Menu  │  │
│  │  (Settings,   │  │ (mt* UDFs)   │  │ (Sign In,    │  │
│  │   Auth, Ref)  │  │              │  │  Refresh)    │  │
│  └──────┬───────┘  └──────┬───────┘  └──────┬───────┘  │
│         │                 │                  │          │
│         └────────┬────────┴──────────────────┘          │
│                  │  Shared Runtime                       │
│         ┌────────┴────────┐                             │
│         │   Auth State    │                             │
│         │   API Client    │                             │
│         │   Config        │                             │
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
```

## Features

- **Microsoft Account SSO** via MSAL.js (Nested App Authentication with Dialog API fallback)
- **Authentik OAuth2 integration** — token exchange or authorization code flow with PKCE
- **8 async custom functions** (UDFs) prefixed with `mt`, all supporting dynamic array spilling
- **Shared runtime** — auth tokens are shared between taskpane, custom functions, and ribbon commands
- **Configurable API endpoint** — change the target API URL at runtime via the Settings panel
- **Ribbon menu** with Sign In, Sign Out, Settings, and Refresh Data commands

## Custom Functions

All functions are available under the `MT` namespace in Excel:

| Function | Description |
|---|---|
| `=MT.MTGETSOURCES([filter])` | List available data sources |
| `=MT.MTGETRECORDS(source, [limit], [offset], [filter])` | Fetch records from a source |
| `=MT.MTGETRECORD(source, recordId)` | Look up a single record by ID |
| `=MT.MTGETSCHEMA(source)` | Get schema/metadata for a source |
| `=MT.MTSEARCH(query, [source], [limit])` | Search across data sources |
| `=MT.MTGETSUMMARY(source, [metric], [field], [filter])` | Aggregated statistics |
| `=MT.MTSTATUS()` | Check connection and auth status |
| `=MT.MTAPICALL(path, [p1Name], [p1Val], [p2Name], [p2Val], [p3Name], [p3Val])` | Generic API call |

All functions:
- Are **async** — they fetch data from the API and return when the response arrives
- Support **dynamic arrays** — results spill into adjacent cells automatically
- Use **cancelable invocations** — pending requests are aborted if the cell is recalculated
- Return inline **error messages** instead of `#VALUE!` for better user experience

## Project Structure

```
├── manifest.xml                  # Office Add-in XML manifest
├── package.json                  # Dependencies and scripts
├── webpack.config.js             # Build configuration
├── tsconfig.json                 # TypeScript configuration
├── .env.example                  # Environment variable template
├── assets/                       # Add-in icons
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-64.png
│   └── icon-80.png
└── src/
    ├── auth/
    │   ├── authConfig.ts         # MSAL init, NAA + Dialog API auth flows
    │   ├── authDialog.ts         # Dialog popup: Authentik OAuth2 + PKCE
    │   └── auth-dialog.html      # Dialog popup HTML shell
    ├── commands/
    │   └── commands.ts           # Ribbon button/menu command handlers
    ├── functions/
    │   └── functions.ts          # Custom function definitions (UDFs)
    ├── shared/
    │   ├── apiClient.ts          # HTTP client with auth headers
    │   ├── config.ts             # Centralized configuration
    │   └── state.ts              # Shared auth state (module-level singleton)
    └── taskpane/
        ├── taskpane.html         # Taskpane HTML (shared runtime entry)
        ├── taskpane.ts           # Taskpane UI logic
        └── taskpane.css          # Taskpane styles (Fluent-inspired)
```

## Prerequisites

- **Node.js** >= 18
- **npm** >= 9
- An **Azure AD App Registration** (for MSAL client ID)
- An **Authentik** instance with an OAuth2/OIDC provider configured
- Your **API** with bearer token validation

## Setup

1. **Clone and install:**

   ```bash
   npm install
   ```

2. **Configure credentials** — copy `.env.example` to `.env` and fill in:
   - `MSAL_CLIENT_ID` — your Azure AD application (client) ID
   - `AUTHENTIK_CLIENT_ID` — your Authentik OAuth2 provider client ID
   - `AUTHENTIK_BASE_URL` — your Authentik instance URL
   - `API_BASE_URL` — your target API base URL

3. **Update `manifest.xml`** — replace placeholder URLs with your hosting domain:
   - Search for `https://localhost:3000` and replace with your production URL
   - Replace `authentik.example.com` in `<AppDomains>` with your Authentik domain
   - Update the `<Id>` GUID with a unique identifier (generate one at https://www.guidgenerator.com)

4. **Update `src/shared/config.ts`** — replace placeholder client IDs and URLs.

## Development

```bash
# Start dev server with hot reload
npm run dev

# Build for production
npm run build
```

The dev server runs at `https://localhost:3000` with a self-signed certificate.

## Sideloading for Testing

### Excel on the Web (Office 365 Online)

1. Open Excel Online at https://www.office.com/launch/excel
2. Open or create a workbook
3. Go to **Insert** > **Office Add-ins** > **Upload My Add-in**
4. Upload `manifest.xml` (or the built copy from `dist/manifest.xml`)

### Excel Desktop

1. Place `manifest.xml` in a network share or local folder
2. In Excel, go to **File** > **Options** > **Trust Center** > **Trusted Add-in Catalogs**
3. Add the folder path as a trusted catalog
4. Restart Excel, then **Insert** > **My Add-ins** > **Shared Folder**

## Deployment

For production deployment:

1. Build the project: `npm run build`
2. Host the `dist/` folder on any HTTPS web server (Azure Static Web Apps, Netlify, Vercel, etc.)
3. Update all URLs in `manifest.xml` to point to the hosted location
4. Deploy the manifest via Microsoft 365 Admin Center for organization-wide availability

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

## License

Private — All rights reserved.
