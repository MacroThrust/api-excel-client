# Authentik test accounts for Microsoft certification

> **Audience:** MacroThrust admins preparing test logins for [Microsoft Marketplace certification](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/add-in-submission-guide).  
> **Related:** [microsoft-marketplace-publication.md](./microsoft-marketplace-publication.md)

---

## How sign-in works in this add-in

Microsoft’s reviewers must be able to sign in **without contacting you**. Your add-in does **not** accept a username/password inside Excel. The flow is:

```
Excel → Sign In
  → Office auth dialog (auth-dialog.html)
  → Authentik authorization page (https://auth.macrothrust.com/...)
  → User signs in with Microsoft (via your Authentik federation)
  → Authentik returns an access token (scopes include macrothrust-api)
  → Add-in calls the API
```

Relevant settings in this repo:

| Setting | Value |
|---------|--------|
| Authentik OAuth client | `macrothrust-excel` |
| Requested scopes | `openid profile email groups macrothrust-api` |
| Redirect URI (production) | `https://macrothrust.github.io/api-excel-client/auth-dialog.html` |

**Important:** Reviewers only need **one set of credentials** — a **Microsoft account** that Authentik accepts and that is authorized for API access. There is no separate “Authentik password” in the Excel UI unless your Authentik login page offers one.

---

## What you need to give Microsoft

In Partner Center → **Review and publish** → **Notes for certification**, provide:

1. **Microsoft test account** — email and password (create a dedicated account; do not use your personal login).
2. **Step-by-step sign-in** — what the reviewer clicks on the Authentik page (e.g. “Sign in with Microsoft”).
3. **Confirmation** that the account already has API access (member of the right Authentik group).
4. **Custom function test** — e.g. run `=MT.STATUS()` after sign-in.

Microsoft will **not** email you for credentials. If anything is missing, the submission fails.

---

## Recommended approach: one Microsoft test account

This is the simplest path when users normally sign in with Microsoft through Authentik.

### Step 1 — Create a Microsoft test account

Create a **new** account reserved for certification, for example:

- A new **@outlook.com** personal account, or
- A user in a **Microsoft 365 developer/test tenant**

Use a strong password and store it in your password manager. Example label: `MT Marketplace certification`.

### Step 2 — Ensure Authentik federates Microsoft sign-in

In your Authentik admin UI (separate project):

1. Go to **Directory → Federation and Social login** (or **Sources**).
2. Confirm you have a **Microsoft / Entra ID** source configured.  
   See [Authentik Entra ID OAuth docs](https://docs.goauthentik.io/users-sources/sources/social-logins/entra-id/oauth/).
3. Confirm that source is part of your **default authentication flow** (the flow shown when the add-in opens Authentik).

On the Authentik login screen, reviewers should see a way to continue with Microsoft (wording depends on your flow).

### Step 3 — Pre-authorize the test account for the API

Signing in with Microsoft is not enough by itself. The Authentik token must include the **`macrothrust-api`** scope (and your API must accept that token).

Do **one** of the following **before** submission:

#### Option A — Pre-create the user and add to a group (recommended)

1. In Authentik: **Directory → Users → Create**.
2. Set **email** to the same address as the Microsoft test account (e.g. `mt-cert-test@outlook.com`).
3. Add the user to the group that grants API access (see Step 4).
4. Have the test user sign in once via a browser (optional but useful):  
   open your Authentik authorize URL or sign in through the Excel add-in to link the Microsoft identity.

#### Option B — Let the first Microsoft login create the user, then authorize

1. Sign in once with the test Microsoft account (via the add-in or Authentik directly).
2. Authentik creates the user from the Microsoft source.
3. In **Directory → Users**, find the new user.
4. Add them to the API access group (Step 4).

Option A is safer for certification: you know the account is ready before Microsoft tests.

### Step 4 — Grant the `macrothrust-api` scope

How this is wired depends on your Authentik app setup. Typically:

1. **Applications → Applications** → open **`macrothrust-excel`** (OAuth2/OpenID provider used by the add-in).
2. Check **Redirect URIs** include:
   ```
   https://macrothrust.github.io/api-excel-client/auth-dialog.html
   ```
   (Add `https://localhost:3000/auth-dialog.html` only for local dev.)
3. Confirm requested scopes include **`macrothrust-api`** (and `openid`, `profile`, `email`, `groups` as in `src/shared/config.ts`).
4. Ensure a **group** (e.g. `api-users` or `marketplace-testers`) is:
   - bound to the application, **or**
   - mapped to the `macrothrust-api` scope via a **scope mapping** / property mapping.

**Verify:** Sign in as the test user, decode the access token at [jwt.io](https://jwt.io), and confirm `macrothrust-api` (or your API audience) appears in `scope` or claims. Then run `=MT.STATUS()` in Excel.

### Step 5 — Provision API data for the test user

The test account must be able to call the API:

- If the API checks Authentik groups or scopes, ensure the test user has them.
- If the API needs a tenant or license, provision that for the test email.
- Seed at least one data source so `=MT.GETSOURCES()` returns something meaningful (or document the expected empty/error state).

### Step 6 — End-to-end test as a reviewer would

1. Use a clean browser profile or InPrivate window.
2. Sideload or open the add-in in [Excel on the web](https://www.office.com/launch/excel).
3. **MT Menu → Sign In**.
4. Complete Authentik → **Microsoft** with the **test** account (not your admin account).
5. Run `=MT.STATUS()` and `=MT.GETSOURCES()`.

If this works externally (not on your corporate VPN only), it is ready for certification notes.

---

## Alternative: Authentik local password (optional)

You can add a **username/password** stage to an Authentik flow for internal testing. **Marketplace reviewers will still see whatever your production authentication flow shows.** If production is Microsoft-only, give them Microsoft credentials, not Authentik passwords.

To create a local Authentik user for **your own** QA:

1. **Directory → Users → Create**
2. Set username and email.
3. **Directory → Users → [user] → Set password**
4. Add user to the API access group.
5. Ensure your authentication flow includes a **Identification** / password stage (not only Microsoft).

This is useful for debugging Authentik without Microsoft; it is usually **not** what you put in Partner Center unless that password login is what production users see.

---

## What to paste into Partner Center

Replace the placeholders with your real test account (example structure):

```
SIGN-IN (required)
1. Open the add-in in Excel on the web.
2. MT Menu → Sign In (or task pane → Sign In).
3. An authentication window opens and redirects to Authentik (auth.macrothrust.com).
4. Click "Sign in with Microsoft" (or equivalent on our Authentik login page).
5. Sign in with the test Microsoft account below.

TEST CREDENTIALS (Microsoft — used via Authentik)
Email: mt-cert-test@outlook.com
Password: <your test password>

This account is pre-authorized for API access. No separate Authentik password is required.

CUSTOM FUNCTION TEST
1. After sign-in, enter =MT.STATUS() in cell A1 — expect a health/auth message.
2. Enter =MT.GETSOURCES() in cell A1 — expect a list of data sources.

If sign-in fails, confirm pop-ups are allowed for Excel and login.microsoftonline.com.
```

Check **“Uses Microsoft Entra ID / SSO”** on the Partner Center product setup page.

---

## Checklist before submit

- [ ] Dedicated Microsoft test account created (not your personal account)
- [ ] Test account added to Authentik group that grants `macrothrust-api`
- [ ] Redirect URI on `macrothrust-excel` matches production `auth-dialog.html`
- [ ] Test user can complete sign-in from Excel on the web in a fresh browser
- [ ] `=MT.STATUS()` and `=MT.GETSOURCES()` work for that user
- [ ] Credentials and steps pasted into **Notes for certification**
- [ ] **Additional purchases** checkbox set correctly if API access is paid/separate

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---------|----------------|-----|
| Authentik login page never opens | Pop-up blocked | Allow pop-ups for Excel / Office domains |
| Microsoft login works but `=MT.STATUS()` shows auth error | Token missing `macrothrust-api` scope | Add user to correct Authentik group / scope mapping |
| “Redirect URI mismatch” | OAuth client misconfigured | Add exact `auth-dialog.html` URL to `macrothrust-excel` |
| User created on first login but no API access | Group not assigned | Add user to API group in Authentik admin |
| Works for you, fails for Microsoft | Account not pre-provisioned; VPN-only API | Pre-create user; ensure API is public HTTPS |

---

## Revision log

| Date | Notes |
|------|--------|
| 2026-06-16 | Initial guide for Marketplace certification test accounts |
