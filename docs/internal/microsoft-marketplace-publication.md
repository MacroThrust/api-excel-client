# Microsoft Marketplace publication — internal checklist

> **Audience:** MacroThrust team preparing to publish **MT Data Connector** (Excel Office Add-in) on [Microsoft Marketplace](https://marketplace.microsoft.com/) for public availability.
>
> **Last reviewed against Microsoft docs:** June 2026  
> **Official references:**
> - [Publish Office Add-ins to Microsoft Marketplace](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish-office-add-ins-to-appsource)
> - [Microsoft 365 app publishing checklist](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/checklist)
> - [Step-by-step submission guide](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/add-in-submission-guide)
> - [Certification policies (Office Add-ins §1120)](https://learn.microsoft.com/en-us/legal/marketplace/certification-policies#1120-office-add-ins-word-excel-powerpoint-and-outlook)
> - [Microsoft Publisher Agreement](https://learn.microsoft.com/en-us/legal/marketplace/msft-publisher-agreement)

---

## Summary

This add-in is an **Excel task pane add-in with custom functions** (`MT.*` namespace), hosted on **GitHub Pages**, using **Microsoft Account (MSAL)** plus **Authentik OAuth2** to reach a protected API.

Publishing for **everyone** (public Marketplace listing) requires:

1. A **Microsoft Partner Center** account enrolled in the **Microsoft 365 and Copilot program**
2. A **production-hosted manifest** (HTTPS, no localhost) — already automated via GitHub Actions
3. **Legal pages**: privacy policy, support URL, EULA (yours or Microsoft’s standard contract)
4. **Certification test notes** with working demo credentials (reviewers cannot email you for access)
5. **Store listing** metadata, screenshots, and accurate disclosures about required external services

**Timeline expectation:** Microsoft states first-time submissions often take **3–6 weeks** and may require multiple resubmissions.

---

## Repo readiness (what is already in place)

| Item | Status | Location |
|------|--------|----------|
| Office Add-in manifest (XML) | Done | `manifest.xml` |
| Manifest schema validation (`npm run validate`) | Done | `package.json` |
| Production manifest validation (`npm run validate:production`) | Done | builds `dist/manifest.xml` first |
| Icons (16, 32, 64, 80 px) | Done | `assets/` |
| Custom functions + ribbon commands | Done | `src/functions/`, `src/commands/` |
| User guide (end users) | Done | `docs/user-guide.md` |
| Support page | Done | `docs/support.md` → `…/docs/support.html` |
| Privacy policy | Done | `docs/privacy-policy.md` → `…/docs/privacy-policy.html` |
| CI deploy to GitHub Pages | Done | `.github/workflows/deploy.yml` |
| HTTPS hosting | Done | GitHub Pages |

---

## ACTION REQUIRED — complete before submitting

Use this section as your working checklist. Search for `<!-- TODO:` in this repo to find every placeholder.

### 1. Partner Center & publisher identity

| Field | Current value in repo | What you must do |
|-------|----------------------|------------------|
| Publisher / `ProviderName` | `MacroThrust` (`manifest.xml`) | **<!-- TODO: Confirm this exactly matches the Publisher name in Partner Center (must be identical or very similar). -->** |
| Partner Center enrollment | Not in repo | **<!-- TODO: Enroll in [Microsoft 365 and Copilot program](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/open-a-developer-account). -->** |
| Payout / tax profile | Not in repo | **<!-- TODO: Complete if you plan to monetize; not required for free offers but account must be in good standing. -->** |
| Microsoft Publisher Agreement | Not in repo | **<!-- TODO: Accept current [Microsoft Publisher Agreement](https://learn.microsoft.com/en-us/legal/marketplace/msft-publisher-agreement) in Partner Center. -->** |

### 2. Manifest & hosting (production)

| Field | Current value | What you must do |
|-------|---------------|------------------|
| Add-in ID (`<Id>`) | `8269bd0d-f5bf-49bc-9c0d-e5dfedf8cbdb` | Keep stable across releases; do not change after first publish |
| Display name | `MT Data Connector` | Must match Partner Center offer name (identical or very similar) |
| Version | `1.0.0.0` in manifest / `1.0.0` in `package.json` | Bump both together on each Marketplace release |
| Production host | `https://macrothrust.github.io/api-excel-client/` | **<!-- TODO: Confirm this is the permanent production URL, or replace `ADDIN_HOST` in CI if you move to a custom domain. -->** |
| Support URL | `…/docs/support.html` | Verify deployed URL returns 200 after merge (contact: mattange@gmail.com) |
| `AppDomains` | `login.microsoftonline.com`, `auth.macrothrust.com`, `api.macrothrust.com` | **<!-- TODO: Confirm all auth/API domains are production-ready and listed in Azure/Authentik redirect URIs. -->** |

**Pre-submit command (run locally or in CI):**

```bash
ADDIN_HOST=https://macrothrust.github.io/api-excel-client npm run validate:production
```

Upload **`dist/manifest.xml`** (not the source `manifest.xml`) to Partner Center → **Packages**.

### 3. Legal & compliance URLs (Partner Center → Properties)

All links must be **valid `https://` URLs** (not email addresses), publicly reachable, and **no login wall**.

| Requirement | Draft in repo | Final URL for Partner Center |
|-------------|---------------|------------------------------|
| **Privacy policy** | `docs/privacy-policy.md` → `…/docs/privacy-policy.html` | `https://macrothrust.github.io/api-excel-client/docs/privacy-policy.html` |
| **Support page** | `docs/support.md` → `…/docs/support.html` | `https://macrothrust.github.io/api-excel-client/docs/support.html` |
| **EULA** | Not in repo | **<!-- TODO: Choose Microsoft Standard Contract in Partner Center OR host your own EULA at `https://_______________` -->** |

**Privacy policy must include (Microsoft will reject otherwise):**

- How you handle users’ personal information
- Reference to **this app** by name (`MT Data Connector`), not only your corporate website
- A description of the service
- A URL that does not 404

A Terms of Use page **does not** substitute for a privacy policy.

### 4. External services & disclosures

This add-in depends on services outside Microsoft. You **must** disclose these in the Marketplace **long description** and in **Notes for certification**.

| Service | Purpose | Disclosure needed |
|---------|---------|-------------------|
| Microsoft Entra ID / MSAL | User sign-in | Check **“Uses Microsoft Entra ID / SSO”** in Partner Center product setup |
| Authentik (`auth.macrothrust.com`) | OAuth2 / token exchange | **<!-- TODO: Describe whether customers need their own Authentik instance or use MacroThrust-hosted auth. -->** |
| MacroThrust API (`api.macrothrust.com`) | Data source for formulas | **<!-- TODO: State if API access is included, requires separate subscription, or is enterprise-only. -->** |
| GitHub Pages | Add-in static hosting | Usually not listed to end users; ensure SLA is acceptable |

**Additional purchases checkbox:** If users need a **paid API subscription**, Authentik tenant, or other license, check **“Requires additional purchases”** and provide test credentials below.

**Enterprise-only option:** If the app is **not** for individual consumers, state that explicitly in **Notes for certification** per [Marketplace FAQ](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/checklist).

### 5. Azure AD app registration (MSAL)

Current compile-time values in `src/shared/config.ts`:

| Setting | Current value |
|---------|---------------|
| `msalClientId` | `4305925c-6f37-4f8d-b6db-ef43a636479a` |
| `msalAuthority` | `https://login.microsoftonline.com/common` |

**<!-- TODO: Complete Azure portal checklist: -->**

- [ ] App registration owned by MacroThrust publisher account
- [ ] Redirect URIs include production add-in origin (`https://macrothrust.github.io/api-excel-client/*` or custom domain)
- [ ] Redirect URIs include `https://auth.macrothrust.com/...` if using dialog flow
- [ ] API permissions: `User.Read`, `openid`, `profile`, `email` (and any others you use)
- [ ] **Publisher verification** completed in Azure (recommended for Marketplace trust)
- [ ] Separate **multi-tenant** vs **single-tenant** decision documented: **<!-- TODO: `common` / specific tenant ID: ___________ -->**

> Do not change `src/shared/config.ts` unless IDs or authority URLs are final. Rebuild and redeploy after any change.

### 6. Authentik OAuth2 provider

Current compile-time values in `src/shared/config.ts`:

| Setting | Current value |
|---------|---------------|
| `authentikClientId` | `macrothrust-excel` |
| Default `authentikBaseUrl` | `https://auth.macrothrust.com` |
| Scopes | `openid profile email groups macrothrust-api` |

**How to create test logins for Microsoft certification:** see **[authentik-test-accounts.md](./authentik-test-accounts.md)**.

Summary: create a **dedicated Entra test user** in the same tenant as your app registration (Azure / Entra admin center → Users), ensure they can sign in through **Authentik’s Microsoft federation**, and **pre-add the user to the Authentik group** that grants the `macrothrust-api` scope. Paste that user’s email/password into Partner Center **Notes for certification**.

**<!-- TODO: Authentik production checklist: -->**

- [ ] OAuth2 provider `macrothrust-excel` exists in production Authentik
- [ ] Redirect URIs include `https://macrothrust.github.io/api-excel-client/auth-dialog.html`
- [ ] Certification test user in API access group (`macrothrust-api` scope)
- [ ] End-to-end test completed (see authentik-test-accounts.md)

### 7. Certification test notes (Partner Center → Review and publish)

Copy the block below into **Notes for certification**. Replace every `<!-- TODO: -->` before submitting.

---

#### Certification notes — draft (paste into Partner Center)

```
App name: MT Data Connector
Manifest ID: 8269bd0d-f5bf-49bc-9c0d-e5dfedf8cbdb
Production manifest URL: https://macrothrust.github.io/api-excel-client/manifest.xml

PLATFORMS TO TEST
- Excel on the web (primary)
- Excel on Windows (Microsoft 365)
- Excel on Mac (Microsoft 365)

SIGN-IN FLOW
1. Install/open the add-in in Excel.
2. Home ribbon → MT Data Connector → MT Menu → Sign In (or task pane Sign In).
3. An auth dialog opens and redirects to Authentik (auth.macrothrust.com).
4. On the Authentik page, choose Sign in with Microsoft.
5. Sign in with the test Microsoft account below.

See docs/internal/authentik-test-accounts.md for how to create and authorize test users in Authentik.

TEST CREDENTIALS (required — reviewers cannot contact us)
Microsoft account (via Authentik federation):
  Email: mt-cert-test@<your-tenant>.onmicrosoft.com
  Password: ___________

(Create this user in Microsoft Entra admin center → Identity → Users → Create new user,
 in the same tenant as your app registration. Uncheck "require password change at first sign-in".)

This account must already be in the Authentik group that grants macrothrust-api scope.
No separate Authentik password is required unless your login flow shows one.

CUSTOM FUNCTION TEST (required for add-ins with custom functions)
1. After sign-in, wait for automatic reload or click MT Menu → Reload API Functions.
2. In cell A1 enter: =MT.STATUS()
   Expected: health/auth status string (not #NAME?).
3. In cell A1 enter: =MT.GETSOURCES()
   Expected: spilled array of data sources (or clear error if API empty).
4. Optional: =MT.VERSION() shows add-in version.

<!-- TODO: Add one formula that returns meaningful sample data in the certification environment. -->

SETTINGS PANEL
- API base URL and Authentik URL can be changed in task pane Settings (for enterprise deployments).
- Defaults: https://api.macrothrust.com/v1 and https://auth.macrothrust.com

ADDITIONAL PURCHASES
<!-- TODO: Yes/No — if Yes, explain what reviewers need (API key, subscription, tenant invite) and provide it above. -->

CRYPTOGRAPHY
<!-- TODO: Yes/No — if Yes, describe use (TLS only / token signing / etc.) per export compliance questions in Partner Center. -->

KNOWN LIMITATIONS
- Requires network access to Authentik and MacroThrust API.
- Dynamic OpenAPI functions depend on user OAuth scopes; unauthorized endpoints return errors or are hidden.
```

---

Optional: attach a PDF with screenshots to **Additional certification info** (Partner Center allows images in PDF).

### 8. Store listing (Partner Center → Marketplace listings)

Prepare **before** submission. Content is not stored in this repo except drafts below.

| Asset | Requirement | Status |
|-------|-------------|--------|
| App name | Same as manifest `DisplayName` | `MT Data Connector` |
| Short description | ≤256 chars, clear value prop | **<!-- TODO: Write final copy -->** |
| Long description (HTML) | Features, setup, external service names + links | **<!-- TODO: Write final HTML in `docs/internal/store-listing-draft.md` or external doc -->** |
| Search keywords | Optional | **<!-- TODO: ___________ -->** |
| Categories | 1–3 required | Suggested: **Productivity**, **Data + analytics** — **<!-- TODO: confirm -->** |
| Screenshots | At least 1; recommended 3–5 | **<!-- TODO: Capture from Excel web + desktop; no sensitive data -->** |
| Store icons | 300×300 px PNG (separate from manifest icons) | **<!-- TODO: Export from `assets/icon-master.png` -->** |
| Video demo | Optional but recommended | **<!-- TODO: URL ___________ -->** |
| Supported languages | Must match manifest locales | Currently `en-GB` only — **<!-- TODO: Add locales if localizing -->** |

**First-run experience (policy §1120):** Users must understand the value **before** sign-in. Ensure screenshots/description explain what data/formulas they get after authentication.

### 9. Accessibility & quality

Microsoft requires apps to work on all supported platforms without errors.

- [ ] Test sign-in, `=MT.STATUS()`, and one data formula on **Excel on the web**
- [ ] Test same on **Excel Windows** and **Excel Mac**
- [ ] Task pane readable with keyboard navigation (see [Office Add-ins accessibility](https://learn.microsoft.com/en-us/office/dev/add-ins/design/accessibility))
- [ ] No unexpected document modifications (policy: add-in must not alter workbook without user action)
- [ ] Pop-ups only on explicit user action (Sign In opens auth dialog — OK)

### 10. Licensing & monetization

| Decision | Your choice |
|----------|-------------|
| Marketplace price | **<!-- TODO: Free / Paid / Contact Me -->** |
| API pricing | **<!-- TODO: Included / BYOL / Subscription URL ___________ -->** |
| License in repo | Currently “Private — All rights reserved” in README | **<!-- TODO: Choose SPDX/license for public distribution if required by legal -->** |

### 11. Post-publication

- [ ] Monitor Partner Center **Action Center** for certification feedback
- [ ] Bump `package.json` version + `manifest.xml` `<Version>` for updates; redeploy GitHub Pages before submitting manifest update
- [ ] Consider [Microsoft 365 App Compliance Program](https://learn.microsoft.com/en-us/microsoft-365-app-certification/overview) for enterprise trust (optional)
- [ ] Plan for re-certification when changing pricing, fulfillment URLs, or major features

---

## Submission workflow (quick reference)

```
1. Partner Center → Microsoft 365 and Copilot → + New offer → Office Add-in
2. Product setup (SSO, additional purchases, CRM)
3. Packages → upload dist/manifest.xml (after validate:production passes)
4. Properties → categories, privacy, support, EULA
5. Marketplace listings → descriptions, icons, screenshots
6. Availability → schedule publish date (cannot change schedule after first publish)
7. Review and publish → Notes for certification (test accounts!)
8. Wait 3–5 business days (up to 4 weeks first time)
```

---

## Files to edit before go-live

| File | Why |
|------|-----|
| `docs/support.md` | Support contact (mattange@gmail.com) — update only if contact changes |
| `docs/privacy-policy.md` | Privacy policy — update only if data practices change |
| `docs/internal/microsoft-marketplace-publication.md` | This checklist — complete all TODOs |
| `manifest.xml` | Only if changing `ProviderName`, domains, or display strings |
| `src/shared/config.ts` | Only when production client IDs/URLs are finalized |
| `package.json` | Version bumps for releases |

**Do not change** `src/functions/functions.ts`, `src/auth/*` logic, or ribbon handlers unless a certification issue requires it.

---

## Internal contacts

| Role | Name | Email |
|------|------|-------|
| Publication owner | **<!-- TODO: ___________ -->** | **<!-- TODO: ___________ -->** |
| Legal (EULA in Partner Center) | **<!-- TODO: ___________ -->** | **<!-- TODO: ___________ -->** |
| Azure / Entra admin | **<!-- TODO: ___________ -->** | **<!-- TODO: ___________ -->** |
| Authentik admin | **<!-- TODO: ___________ -->** | **<!-- TODO: ___________ -->** |
| API / backend owner | **<!-- TODO: ___________ -->** | **<!-- TODO: ___________ -->** |

---

## Revision log

| Date | Author | Notes |
|------|--------|-------|
| **<!-- TODO: ___________ -->** | **<!-- TODO: ___________ -->** | Initial publication prep |
