# clasp setup — GESI assets project

Push `Assets-GESI.gs` straight to the bound Apps Script project instead of copy-pasting.

- **Project (script) ID:** `1Ln7wFngV38LVNtD4v8hvlAZLAd8gmCGLPqlCgIXBkfaBdXhJ-yXSsbZb`
- **Bound sheet:** `1f8svzMOn0JNsbn3VeQQF108j7LmAGWeFr3Xlb7H-q8Y`
- clasp is already installed (`clasp 3.3.0`).
- The GESI library dependency (id `1KjnRVV…u0s3`, v53, symbol `GESI`) is **already in `appsscript.json`**, so `clasp push` carries it — no editor step needed.

## One-time

1. **Enable the Apps Script API** for your account: <https://script.google.com/home/usersettings> → *Apps Script API* = ON.
2. **Log in:** `clasp login` (opens a browser; use the GESI-authorized Google account).

## Push code

From this folder (note the spaces — quote the path):

```powershell
cd "C:\Users\colt2\Documents\GitHub\EveSheetsCode\2026 Inventory Automation"
clasp push -f
```

`-f` forces overwrite of the remote manifest. Safe here: our local `appsscript.json` already includes the GESI dependency and the web-app config, so nothing is lost.

> If you ever bump the GESI version in the editor, run `clasp pull` once to sync the new version number back into the local `appsscript.json`.

## Set the shared token (once)

The token is **not** in source. Set it in both places to the same value:

- **Server:** Apps Script editor → Project Settings → *Script Properties* → add `GESI_TOKEN`.
- **PC:** `setx GESI_TOKEN "<same value>"` then reopen the shell (so `Refresh-Inventory.ps1` reads it).

`clasp` cannot set Script Properties — do this in the editor UI.

## Deploy as Web App

```powershell
clasp deploy -d "gesi-assets"        # first deploy → note the Deployment ID
clasp deployments                    # list IDs
```

- Callable URL: `https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec`.
- Paste it into `$WebAppUrl` in `Refresh-Inventory.ps1` (lives in the **eve-windows-helpers** repo, under `2026 Inventory Automation\`).
- Web-app behavior comes from `appsscript.json`: `executeAs USER_DEPLOYING` (runs as you → GESI's tokens) and `access ANYONE_ANONYMOUS` (the PowerShell POST needs no Google auth; the shared `token` is the gate).
- **Keep the URL stable on updates** — redeploy the *same* deployment:
  ```powershell
  clasp deploy -i <DEPLOYMENT_ID> -d "gesi-assets"
  ```

## Everyday loop

```powershell
clasp push -f
clasp deploy -i <DEPLOYMENT_ID> -d "gesi-assets"   # same /exec URL
```
