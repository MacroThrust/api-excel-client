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

### Step 1 — Create a test user in your Azure tenant (recommended)

If you already registered the add-in’s app in **Microsoft Entra ID** (same place as `msalClientId` in `src/shared/config.ts`), create the certification user **in that same tenant**. That keeps sign-in consistent with your app registration and Authentik Microsoft federation.

#### Where to go (same portal as your app registration)

**Option A — Microsoft Entra admin center (usual)**

1. Open [https://entra.microsoft.com](https://entra.microsoft.com) and sign in as a tenant admin.
2. Go to **Identity → Users → All users**  
   (same tenant where you have **Applications → App registrations** for the Excel add-in).
3. Click **New user → Create new user**.
4. Fill in:
   - **User principal name:** e.g. `mt-cert-test@<your-tenant>.onmicrosoft.com`  
     (or your verified custom domain, e.g. `mt-cert-test@macrothrust.com` if configured).
   - **Display name:** e.g. `Marketplace Certification Test`
   - **Password:** choose **Let me create the password** and set a strong password you can share in Partner Center notes.
   - Uncheck **Require this user to change their password when they first sign in** so Microsoft’s reviewers are not forced through a password-change flow.
5. Click **Create**.
6. Copy the **User principal name** and password — you will paste these into **Notes for certification**.

**Option B — Azure portal**

1. Open [https://portal.azure.com](https://portal.azure.com).
2. Search for **Microsoft Entra ID** (or **Azure Active Directory**) and select your tenant.
3. **Manage → Users → New user → Create new user** — same fields as above.

**Option C — Microsoft 365 Developer Program sandbox**

If your app registration lives in a **developer sandbox tenant** from the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program):

1. Sign in to [https://developer.microsoft.com/microsoft-365/dev-program](https://developer.microsoft.com/microsoft-365/dev-program) and open your sandbox admin portal.
2. Go to **Microsoft Entra ID → Users → New user → Create new user** (same steps as Option A).
3. Sandbox tenants often include sample users; you can still create a dedicated `mt-cert-test@...` user so certification does not share your admin account.

#### Link the test user to your app registration

You do **not** need to add the test user to the app registration’s user list for this add-in (it uses delegated sign-in via Authentik, not per-user app assignment). Do confirm:

| Item | Where to check |
|------|----------------|
| App registration | **Entra → Applications → App registrations** → your app (`4305925c-6f37-4f8d-b6db-ef43a636479a` or your client ID) |
| Supported account types | **Authentication** → if only your org, the test user **must** be in that tenant |
| API permissions | **API permissions** → `User.Read`, `openid`, `profile`, `email` (as in `config.ts`) — admin consent if required |

#### Authentik must trust the same tenant

Your Authentik **Microsoft / Entra ID** source should use the **same directory (tenant) ID** as this Entra tenant (unless you intentionally use `/common` for multi-tenant). See [Authentik Entra ID OAuth](https://docs.goauthentik.io/users-sources/sources/social-logins/entra-id/oauth/).

When the test user signs in on the Authentik page with **Sign in with Microsoft**, Entra authenticates `mt-cert-test@<your-tenant>.onmicrosoft.com` and Authentik maps that identity to an Authentik user.

#### Optional: assign a Microsoft 365 license

Certification testers usually use **Excel on the web** with their own environment; your test account mainly needs to **authenticate**. If you want the test user to open Excel as that identity:

1. **Microsoft 365 admin center** → **Users → Active users** → select the test user.
2. **Licenses and apps** → assign a license that includes Excel (e.g. Microsoft 365 Business Basic).

This is optional for Marketplace notes if reviewers only need your credentials for the **Authentik → Microsoft** step inside the add-in.

#### Alternative: personal @outlook.com account

You can still use a new **@outlook.com** account instead of a tenant user. That works when your Entra app registration allows personal Microsoft accounts (**Accounts in any organizational directory and personal Microsoft accounts**) and Authentik’s Microsoft source is configured for `/common`. A **tenant user in the same directory as your app registration** is usually simpler to control and audit.

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
Email: mt-cert-test@<your-tenant>.onmicrosoft.com
Password: <password you set in Entra → Users → Create new user>

(User created in the same Microsoft Entra tenant as the app registration.)

This account is pre-authorized for API access. No separate Authentik password is required.

CUSTOM FUNCTION TEST
1. After sign-in, enter =MT.STATUS() in cell A1 — expect a health/auth message.
2. Enter =MT.GETSOURCES() in cell A1 — expect a list of data sources.

If sign-in fails, confirm pop-ups are allowed for Excel and login.microsoftonline.com.
```

Check **“Uses Microsoft Entra ID / SSO”** on the Partner Center product setup page.

---

## Checklist before submit

- [ ] Dedicated test user created in **Entra → Users** (same tenant as app registration)
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
| 2026-06-16 | Added Azure Entra tenant test user steps (same portal as app registration) |
