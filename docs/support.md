# MT Data Connector — Support

**Provider:** MacroThrust  
**App:** MT Data Connector (Excel Office Add-in)

---

## Contact

<!-- TODO: Replace with your production support channel before Marketplace submission. -->

| Channel | Details |
|---------|---------|
| Email | **<!-- TODO: support@macrothrust.com -->** |
| Web | **<!-- TODO: https://macrothrust.com/support -->** |
| Hours | **<!-- TODO: e.g. Mon–Fri 09:00–17:00 GMT -->** |

For security issues, contact: **<!-- TODO: security@macrothrust.com -->**

---

## Before you contact support

1. Confirm you are signed in (**MT Menu → Sign In** or the task pane **Sign In** button).
2. Run `=MT.STATUS()` in a cell — note the message shown.
3. Run `=MT.VERSION()` — note the version and build time.
4. If formulas show `#NAME?`, open the task pane and click **Reload API Functions**, or run `=MT.RELOADFUNCTIONS()`.

See the [user guide](user-guide.html) for install and first-run steps.

---

## Common issues

### Sign-in fails or loops

- Allow pop-ups for Excel and `login.microsoftonline.com`.
- Complete both Microsoft sign-in and any Authentik prompt your organization uses.
- If your organization uses a custom Authentik URL, set it in the task pane **Settings** panel.

### Formulas return errors or “Access denied”

- Your account may lack API scopes for that endpoint. Run `=MT.LISTENDPOINTS()` to see permitted operations.
- Click **Reload API Functions** after sign-in or permission changes.

### Add-in does not update after a new release

- **Excel on the web:** Hard-refresh the browser (`Ctrl+Shift+R` / `Cmd+Shift+R`).
- **Excel desktop:** Clear the Office cache and restart Excel. See [Clear the Office cache](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/clear-cache).

### IT deployed the add-in but I cannot see it

- Ask your administrator to confirm deployment in the Microsoft 365 admin center (**Settings → Integrated apps**).
- Open **Insert → Office Add-ins → My Add-ins**.

---

## Documentation

- [User guide](user-guide.html) — install, sign-in, and formulas
- [Privacy policy](privacy-policy.html)
- [Developer README on GitHub](https://github.com/MacroThrust/api-excel-client)

---

## Service dependencies

This add-in requires:

- Microsoft Excel with Office Add-ins support
- Network access to Microsoft sign-in (`login.microsoftonline.com`)
- **<!-- TODO: Describe your Authentik instance, e.g. auth.macrothrust.com -->**
- **<!-- TODO: Describe your API, e.g. api.macrothrust.com — and how customers obtain access -->**

---

*Last updated: <!-- TODO: YYYY-MM-DD -->*
