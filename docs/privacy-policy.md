# Privacy Policy — MT Data Connector

**Effective date:** <!-- TODO: YYYY-MM-DD -->  
**Provider:** MacroThrust (“we”, “us”)  
**Application:** **MT Data Connector** — an Excel Office Add-in that connects workbooks to MacroThrust API services.

> **Internal note:** This is a draft template for Microsoft Marketplace certification.  
> **Legal review required** before publishing. A Terms of Use page does **not** replace this privacy policy.

---

## 1. Scope

This policy describes how **MT Data Connector** handles information when you use the add-in in Microsoft Excel (web, Windows, or Mac). It applies to the add-in and related authentication flows, not only to the MacroThrust corporate website.

---

## 2. Information we collect

When you use MT Data Connector, we may process:

| Category | Examples | Purpose |
|----------|----------|---------|
| Microsoft account identifiers | Display name, email, object ID from Microsoft Entra ID / MSAL | Sign-in and session management |
| OAuth tokens | Microsoft access token, Authentik-issued API token | Authenticate API requests on your behalf |
| API request metadata | Endpoints called, timestamps, error codes | Provide formula data and troubleshoot issues |
| Workbook interaction | Cell formulas you enter (e.g. `=MT.GETSOURCES()`) | Execute requested data operations |
| Local storage | API base URL, Authentik URL preferences | Remember your settings in the browser |

<!-- TODO: Confirm whether you log IP addresses, device info, or workbook contents server-side. -->

We do **not** intentionally collect workbook content beyond what your formulas request from the API.

---

## 3. How we use information

We use the information above to:

- Authenticate you with Microsoft and **<!-- TODO: Authentik / your IdP -->**
- Fetch data from **<!-- TODO: MacroThrust API or customer API -->** for Excel formulas
- Operate, secure, and improve the add-in
- Respond to support requests

<!-- TODO: Add any analytics, crash reporting, or telemetry if used. -->

---

## 4. How we share information

We may share information with:

| Recipient | Why |
|-----------|-----|
| Microsoft | Sign-in via Microsoft identity platform |
| **<!-- TODO: Authentik host, e.g. auth.macrothrust.com -->** | OAuth2 authorization and token issuance |
| **<!-- TODO: API host, e.g. api.macrothrust.com -->** | Data requested by your formulas |
| Infrastructure providers | Hosting the add-in static files (**<!-- TODO: e.g. GitHub Pages / Azure -->**) |

We do not sell personal information.

<!-- TODO: Add subprocessors list or link if required by your jurisdiction. -->

---

## 5. Data retention

<!-- TODO: Specify retention periods, e.g.:

- OAuth tokens: session / until sign-out / local storage only
- API logs: X days
- Support tickets: X years
-->

---

## 6. Security

We use HTTPS for all add-in hosting and API communication. Tokens are stored in the add-in runtime according to Office shared-runtime practices.

<!-- TODO: Link to corporate security page if available. -->

---

## 7. Your rights and choices

Depending on your location, you may have rights to access, correct, delete, or restrict processing of your personal data.

- **Sign out:** Use **MT Menu → Sign Out** to clear the session in the add-in.
- **Remove the add-in:** Uninstall via Excel or your Microsoft 365 administrator.
- **Contact us:** **<!-- TODO: privacy@macrothrust.com -->**

<!-- TODO: Add GDPR / UK GDPR / CCPA-specific sections after legal review. -->

---

## 8. Children

MT Data Connector is not directed at children under 13 (or the age required in your jurisdiction).

---

## 9. Changes to this policy

We may update this policy. We will post the new version at:

**<!-- TODO: https://macrothrust.github.io/api-excel-client/docs/privacy-policy.html -->**

and update the effective date above.

---

## 10. Contact

**MacroThrust**  
**<!-- TODO: Registered address -->**  
Privacy inquiries: **<!-- TODO: privacy@macrothrust.com -->**  
Support: [support page](support.html)

---

*This document must remain publicly accessible at an `https://` URL with no login required for Microsoft Marketplace certification.*
