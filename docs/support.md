# MT Data Connector — Support

**Provider:** MacroThrust  
**App:** MT Data Connector (Excel Office Add-in)

---

## Contact

For help with this add-in, email **[mattange@gmail.com](mailto:mattange@gmail.com)**.

Please include:

- A short description of the problem
- The output of `=MT.VERSION()` if you can run it
- Whether you are using Excel on the web or desktop

---

## Before you email

1. Sign in via **MT Menu → Sign In** or the task pane.
2. Try `=MT.STATUS()` in a cell and note the result.
3. If formulas show `#NAME?`, click **Reload API Functions** or run `=MT.RELOADFUNCTIONS()`.

See the [user guide](user-guide.html) for install and setup steps.

---

## Common issues

**Sign-in fails** — Allow pop-ups for Excel and `login.microsoftonline.com`, then try again.

**Formulas return errors** — Sign in, reload API functions, and confirm your account has access to the API.

**Add-in looks outdated** — Hard-refresh Excel on the web, or clear the Office cache on desktop. See [Clear the Office cache](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/clear-cache).

---

## More information

- [User guide](user-guide.html)
- [Privacy policy](privacy-policy.html)
