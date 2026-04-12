# Contributing to MCP-Access

## Submitting Changes via Pull Request

This project is hosted at [github.com/unmateria/MCP-Access](https://github.com/unmateria/MCP-Access).
The standard open-source contribution flow is: **fork → branch → commit → push → PR**.

---

### Step 1 — Fork the repo on GitHub

Go to `https://github.com/unmateria/MCP-Access` and click **Fork** (top-right).
This creates a copy under your own GitHub account.

---

### Step 2 — Connect your local copy to your fork

Your local repo currently only points to Unmateria's remote. Add your fork as `origin`
and keep Unmateria's as `upstream`:

```bash
cd "C:/Users/JuanSoto/MCP-Access"
git remote add origin https://github.com/<your-github-username>/MCP-Access.git
git remote add upstream https://github.com/unmateria/MCP-Access.git
```

| Remote | Points to |
|--------|-----------|
| `origin` | Your fork — where you push |
| `upstream` | Unmateria's repo — source of truth |

---

### Step 3 — Create a feature branch

```bash
git checkout -b feature/bypass-startup-on-open
```

---

### Step 4 — Commit your changes

```bash
git add access_mcp_server.py
git commit -m "Bypass AutoExec and StartupForm when opening via COM automation

- Set AutomationSecurity=3 (msoAutomationSecurityForceDisable) before
  OpenCurrentDatabase to suppress AutoExec macros.
- Use DAO.DBEngine.36 to temporarily blank the StartupForm database
  property before Access opens the file, then restore it after close.
- Manual opens are unaffected: AutomationSecurity is a COM session
  property (not stored in the DB), and StartupForm is restored before
  the user can open the file.
- Added _dao_suppress_startup() and _dao_restore_startup() helpers.
- _force_cleanup() now clears _saved_startup_form state."
```

---

### Step 5 — Push to your fork

```bash
git push origin feature/bypass-startup-on-open
```

---

### Step 6 — Open a Pull Request

1. Go to `https://github.com/<your-github-username>/MCP-Access`
2. GitHub will show a banner **"Compare & pull request"** — click it
3. Set the target:
   - **base repository:** `unmateria/MCP-Access` → `main`
   - **head:** `<your-github-username>/MCP-Access` → `feature/bypass-startup-on-open`
4. Write a description explaining:
   - **Problem:** Databases with AutoExec macros or startup forms block the COM session,
     causing errors and preventing the MCP server from working
   - **Solution:** `AutomationSecurity = 3` suppresses both AutoExec and startup forms
     during COM automation; manual opens are completely unaffected
   - **Tested with:** `ITImpactOrderDemoBeta.accdb` — startup form (`frmSplash`) was
     suppressed during MCP session and fully intact after close

---

### Keeping your fork up to date

Before starting future contributions, sync with upstream first:

```bash
git fetch upstream
git checkout main
git merge upstream/main
git push origin main
```
