# MT Data Connector — User Guide

This guide is for **people who use the add-in in Excel**. If you build or host the add-in, see the [main README](../README.md) instead.

---

## What this add-in does

MT Data Connector lets you pull data from your organization’s API directly into Excel using formulas (for example `=MT.GETSOURCES()`). After you sign in, the add-in can also register extra functions from your API’s OpenAPI specification—only for endpoints your account is allowed to call.

---

## Before you start

You need:

- **Microsoft Excel** (Excel on the web or Excel desktop with Office Add-ins support)
- Access to the add-in’s **manifest** (usually provided by your IT team or listed on your organization’s internal docs)
- Permission to sign in with your **Microsoft account** and complete any **Authentik** login your organization uses

Your admin may install the add-in for everyone (Microsoft 365 admin center) or ask you to sideload it once using a manifest URL.

---

## Install or open the add-in

### If IT deployed it for you

1. Open Excel and create or open a workbook.
2. Go to **Insert** → **Office Add-ins** (or **My Add-ins**).
3. Find **MT Data Connector** and open it.

### If you sideload it yourself (Excel on the web)

1. Open [Excel on the web](https://www.office.com/launch/excel).
2. Open a workbook.
3. Go to **Insert** → **Office Add-ins** → **Upload My Add-in**.
4. Paste the manifest URL your admin gave you, for example:
   ```
   https://<your-org>.github.io/<repo-name>/manifest.xml
   ```
   Or upload a downloaded `manifest.xml` file.

### Excel desktop (Windows / Mac)

1. Download `manifest.xml` from the same URL with `/manifest.xml` at the end.
2. Follow your organization’s steps for **Trusted Add-in Catalogs** or sideloading from a shared folder.
3. Restart Excel if required, then open the add-in from **Insert** → **My Add-ins**.

---

## First-time setup (about 2 minutes)

### 1. Open the task pane

With the add-in loaded, open **MT Data Connector** from the ribbon (or **MT Menu**). The side panel shows authentication, settings, and function help.

### 2. Sign in

1. In the task pane, click **Sign In**.
2. Complete the Microsoft sign-in window.
3. If prompted, finish sign-in on your organization’s Authentik page.

When sign-in succeeds, the task pane shows your account as authenticated. The add-in usually **reloads API functions automatically** after sign-in.

### 3. Reload API functions (if needed)

If you do not see API-specific functions yet:

- Click **Reload API Functions** in the task pane or ribbon, **or**
- In any cell, enter:
  ```
  =MT.RELOADFUNCTIONS()
  ```

Wait until the formula returns a status of **OK** with a count of registered functions.

### 4. Try a formula

| Goal | Example formula |
|------|-----------------|
| Check you are connected | `=MT.STATUS()` |
| See add-in version | `=MT.VERSION()` |
| List data sources | `=MT.GETSOURCES()` |
| See which API endpoints you can use | `=MT.LISTENDPOINTS()` |

Formulas **spill** results into neighboring cells when Excel supports dynamic arrays.

---

## Everyday use

### Ribbon and menu

Use the **MT Data Connector** group on the ribbon for:

- **Sign In** / **Sign Out**
- **Settings** (opens the task pane)
- **Refresh Data** — recalculates cells that use MT functions
- **Reload API Functions** — refreshes functions from the OpenAPI spec

### Built-in functions

All built-in functions live under the **`MT`** namespace:

| Function | What it does |
|----------|----------------|
| `=MT.GETSOURCES([filter])` | List data sources |
| `=MT.GETRECORDS(source, [limit], [offset], [filter])` | Fetch records |
| `=MT.GETRECORD(source, recordId)` | One record by ID |
| `=MT.GETSCHEMA(source)` | Schema for a source |
| `=MT.SEARCH(query, [source], [limit])` | Search |
| `=MT.GETSUMMARY(source, [metric], [field], [filter])` | Aggregates |
| `=MT.STATUS()` | Connection / auth status |
| `=MT.VERSION()` | Add-in name and version |
| `=MT.APICALL(path, …)` | Generic API path |
| `=MT.RELOADFUNCTIONS()` | Reload OpenAPI functions |
| `=MT.LISTENDPOINTS()` | Endpoints and permissions |

### API functions from OpenAPI

After reload, endpoints from your API appear as functions such as `=MT.GETUSERS()` or `=MT.POSTORDERS(jsonBody)`. Names follow the HTTP method (`GET`, `POST`, etc.) plus the path or `operationId` from the spec.

Only endpoints your token allows are registered. Others may be hidden or return an access error.

---

## Troubleshooting

| Problem | What to try |
|---------|-------------|
| `#NAME?` on a formula | Confirm the add-in is loaded. For API functions, run `=MT.RELOADFUNCTIONS()` and sign in first. |
| “Not authenticated” / 401 | Sign in again from the task pane or ribbon. |
| No API functions after reload | Check with IT that your account has the right API scopes and that the OpenAPI URL is configured. |
| Old behavior after an update | **Excel on the web:** hard refresh (`Ctrl+Shift+R` / `Cmd+Shift+R`). **Desktop:** clear Office cache and restart Excel. The task pane may show an **Update available** banner. |
| Formula shows an error message in cells | Read the spilled text; the add-in returns readable errors instead of only `#VALUE!`. |

---

## Where to get more help

- **In Excel:** use the **Built-in Functions** and **API Functions** sections in the task pane.
- **Online (same host as the add-in):** open the [User guide (HTML)](user-guide.html) shipped with the add-in site—see below.
- **Developers / admins:** [README](../README.md) (configuration, deployment, authentication architecture).

---

## Viewing this guide online

| Location | URL / path |
|----------|------------|
| **GitHub repository** | `docs/user-guide.md` in the project (this file) |
| **GitHub Pages** (after deploy) | `https://<owner>.github.io/<repo>/docs/user-guide.html` |

Replace `<owner>` and `<repo>` with your organization’s GitHub Pages host. The HTML version is published automatically with each production build.
