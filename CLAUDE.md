# CLAUDE.md — Bench Dashboard

Internal EPAM tool for delivery managers to track bench associates: course completion, client pipeline, and placement outcomes. React SPA hosted on GitHub Pages; all data in manager's EPAM OneDrive via Microsoft Graph API.

## Tech Stack

- React 19 + Vite 8 + Tailwind CSS 3
- `@azure/msal-browser` + `@azure/msal-react` for Azure AD OAuth 2.0 + PKCE
- Microsoft Graph API for OneDrive read/write
- Chart.js 4 (analytics panel + PDF charts)
- jsPDF + jspdf-autotable (PDF export)
- GitHub Pages (hosting) + GitHub Actions (CI/CD)

## Key Files

| File | Role |
|---|---|
| `src/App.jsx` | Entire application — single React component, ~900 lines |
| `src/main.jsx` | MSAL `PublicClientApplication` setup + React root |
| `vite.config.js` | `base: './'` — relative paths required for GitHub Pages |
| `.github/workflows/deploy.yml` | Build + deploy on push to main |
| `.env` | `VITE_AAD_CLIENT_ID` + `VITE_AAD_TENANT_ID` (not committed) |
| `.env.example` | Template for env vars |

## Architecture in One Paragraph

The app is a pure client-side SPA. On load, `main.jsx` creates an MSAL `PublicClientApplication` with the redirect URI set to `window.location.origin + window.location.pathname`. `App.jsx` checks `USE_CLOUD = !!CLIENT_ID` — if true, shows a Microsoft login screen. After login, a `useEffect` fires, acquires a Bearer token via `acquireTokenSilent`, and GETs `bench-data.json` from the manager's OneDrive via `https://graph.microsoft.com/v1.0/me/drive/root:/BenchDashboard/bench-data.json:/content`. A second `useEffect` debounces writes (1200ms) back to the same endpoint via PUT whenever `associates` or `coursePool` state changes.

## Critical State Patterns — Read Before Editing

### skipSave + cloudReady refs
Two refs prevent a race condition where the initial cloud load triggers an immediate save (echo):

```js
const skipSave = useRef(false);  // set true once after load; persist effect skips once
const cloudReady = useRef(false); // set true after load completes; persist effect skips until then
```

**In the load effect:** `if (record) skipSave.current = true;` — only set on actual data, not on 404.
**Do NOT** set `skipSave` when the load returns 404 (no file yet) — it would drop the first real save.

### getToken
```js
const getToken = useCallback(async () => {
  try {
    const res = await instance.acquireTokenSilent({ scopes: TOKEN_SCOPE, account: accounts[0] });
    return res.accessToken;
  } catch {
    const res = await instance.acquireTokenPopup({ scopes: TOKEN_SCOPE });
    return res.accessToken;
  }
}, [instance, accounts]);
```
Silent first, popup fallback. `getToken` is in both effect dependency arrays — do not remove it.

### ID-based state (never use array index as ID)
Associate IDs are assigned at creation: `id: Math.max(0, ...associates.map(a => a.id)) + 1`
All updates use `.map(a => a.id === id ? {...a, ...changes} : a)` — never splice by index.

### loginRedirect, not loginPopup
Popup gets stuck showing the app's login screen inside the popup (redirect loops). Always use:
```js
instance.loginRedirect({ scopes: TOKEN_SCOPE })
```

### Chart.js cleanup
Analytics panel stores chart instances in refs (`chartInst1`, `chartInst2`). Before creating a new chart, the effect always calls `chartInst.current.destroy()` if the ref is non-null. Without this, Chart.js throws "Canvas is already in use".

## Data Schema

```json
{
  "associates": [{
    "id": 1,
    "name": "Jane Smith",
    "joinDate": "2025-01-15",
    "archived": false,
    "courses": { "GenAI": 85, "Claude": 60 },
    "clients": [{
      "client": "Acme Corp",
      "projectCode": "PRJ-001",
      "status": "Client Round",
      "proposedDate": "2025-02-01",
      "selectedDate": "",
      "rejectedDate": ""
    }]
  }],
  "coursePool": ["GenAI", "Claude"]
}
```

`migrate()` adds `archived: false` to any record missing the field (backward compat).

## Environment Variables

| Var | What it is | Secret? |
|---|---|---|
| `VITE_AAD_CLIENT_ID` | Azure AD App Registration Client ID | No — identifies the app, not a credential |
| `VITE_AAD_TENANT_ID` | Azure AD Directory (tenant) ID | No — identifies the tenant |

Both are baked into the JS bundle at build time by Vite. This is safe — they don't grant access without the user's own EPAM credentials.

## Running Locally

```bash
npm install
# set VITE_AAD_CLIENT_ID and VITE_AAD_TENANT_ID in .env
npm run dev   # http://localhost:5173
```

`http://localhost:5173` must be registered as a redirect URI in the Azure AD App Registration.

## Deploying

Push to `main` — GitHub Actions handles build and deploy automatically (~2 min).
GitHub Secrets required: `VITE_AAD_CLIENT_ID`, `VITE_AAD_TENANT_ID`.

## What NOT to Do

- **Do not add a router** — single view with conditional rendering is intentional; no navigation needed
- **Do not switch to localStorage-only** — `USE_CLOUD = !!CLIENT_ID` already handles local fallback; OneDrive is the production path
- **Do not use `loginPopup`** — confirmed broken (popup stays open with login screen)
- **Do not move env vars to a backend** — there is no backend; this is a static SPA by design
- **Do not commit `.env`** — it's gitignored; use GitHub Secrets for CI/CD
- **Do not split into multiple components prematurely** — the single-component approach is deliberate given the app's scope

## Common Tasks

**Add a new field to associates:** Update the state type, the add handler, the render in Manager View, the Team View render, and the PDF export table columns. Also update `migrate()` if a default value is needed for existing records.

**Add a new chart:** Create a canvas ref + instance ref pair. In the analytics `useEffect`, destroy the old instance if it exists, then `new Chart(ref.current, config)`. Capture the instance in the ref.

**Change the debounce delay:** Find `setTimeout` in the persist `useEffect` — single number to change.

**Add a new course status colour:** Find the inline style/class logic on course progress bars in the associate expand panel.
