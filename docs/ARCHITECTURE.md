# Architecture — Bench Dashboard

## System Overview

```
┌─────────────────────────────────────────────────────────┐
│                    EPAM Employee Browser                 │
│                                                         │
│  ┌─────────────────────────────────────────────────┐   │
│  │         React SPA (GitHub Pages)                │   │
│  │  - Loads HTML/JS/CSS from GitHub's CDN          │   │
│  │  - Zero sensitive data in static files          │   │
│  └──────────────┬──────────────────┬───────────────┘   │
│                 │                  │                    │
│         [1] Auth Request    [2] Data Read/Write         │
│                 │                  │                    │
└─────────────────┼──────────────────┼────────────────────┘
                  │                  │
     ┌────────────▼──────┐  ┌────────▼──────────────────┐
     │  Azure AD (EPAM)  │  │  Microsoft Graph API      │
     │  login.microsoft  │  │  graph.microsoft.com      │
     │  online.com       │  │                           │
     │                   │  │  OneDrive (EPAM tenant)   │
     │  Issues Bearer    │  │  /me/drive/root:          │
     │  Token (JWT)      │  │  /BenchDashboard/         │
     └───────────────────┘  │  bench-data.json          │
                            └───────────────────────────┘
```

**Key property:** GitHub Pages (external) serves only code — HTML, CSS, JavaScript. All actual bench data lives inside Microsoft's infrastructure (EPAM OneDrive). An attacker who gains access to GitHub gets zero business data.

---

## Authentication Flow

The app uses **OAuth 2.0 Authorization Code Flow with PKCE** (Proof Key for Code Exchange), the industry standard for browser-based SPAs.

```
1. User clicks "Sign in with Microsoft"
       │
2. MSAL generates code_verifier + code_challenge (PKCE)
   Saves state to sessionStorage
       │
3. Browser redirects to:
   https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize
   ?client_id={app_id}
   &scope=Files.ReadWrite openid profile
   &code_challenge={PKCE}
   &redirect_uri=https://saurabhsuman1-droid.github.io/bench-dashboard/
       │
4. EPAM Azure AD authenticates user (SSO if already logged in)
       │
5. Microsoft redirects back to app with:
   https://saurabhsuman1-droid.github.io/bench-dashboard/#code={auth_code}&state={...}
       │
6. MSAL processes auth code, verifies state + PKCE
   Exchanges code for Access Token + Refresh Token at token endpoint
   Stores tokens in sessionStorage
       │
7. App calls acquireTokenSilent() for subsequent requests
   (auto-refreshes via Refresh Token — user never sees login again)
```

### Why This Is Secure

- **PKCE** prevents authorization code interception attacks
- **State parameter** prevents CSRF attacks
- **sessionStorage** — tokens cleared when browser tab closes
- **No client secret** — SPAs cannot safely store secrets; PKCE replaces the secret
- **Scope `Files.ReadWrite`** — token only grants access to files the signed-in user has permission to access
- **Single-tenant** — only EPAM Azure AD accounts can authenticate; personal Microsoft accounts are rejected

---

## Data Flow

### Read (on login)

```
isAuthenticated → true
       │
acquireTokenSilent({ scopes: ["Files.ReadWrite"] })
       │
GET https://graph.microsoft.com/v1.0/me/drive/root:/BenchDashboard/bench-data.json:/content
Authorization: Bearer {access_token}
       │
Response: JSON blob → parsed → React state updated
(404 on first use → app starts empty, creates file on first save)
```

### Write (on any data change)

```
User edits data (add associate / update course / etc.)
       │
React state updates (immutable, ID-based)
       │
Persist useEffect fires (deps: [associates, coursePool, getToken])
       │
Debounce timer: 1200ms (clears + resets on rapid changes)
       │
acquireTokenSilent() → fresh Bearer token
       │
PUT https://graph.microsoft.com/v1.0/me/drive/root:/BenchDashboard/bench-data.json:/content
Authorization: Bearer {access_token}
Content-Type: application/json
Body: { "associates": [...], "coursePool": [...] }
       │
200 OK → setSyncing(false) → header shows "· synced"
Error → setSyncError(true) → header shows "· ⚠ sync failed"
```

---

## Data Schema

Stored as a single JSON file at `OneDrive/BenchDashboard/bench-data.json`:

```json
{
  "associates": [
    {
      "id": 1,
      "name": "Jane Smith",
      "joinDate": "2025-01-15",
      "archived": false,
      "courses": {
        "GenAI": 85,
        "Claude": 60,
        "AWS Cloud": 0
      },
      "clients": [
        {
          "client": "Acme Corp",
          "projectCode": "PRJ-001",
          "status": "Client Round",
          "proposedDate": "2025-02-01",
          "selectedDate": "",
          "rejectedDate": ""
        }
      ]
    }
  ],
  "coursePool": ["GenAI", "Claude", "AWS Cloud"]
}
```

**Design decisions:**
- Single flat JSON file — simple, no joins, easy to inspect/backup
- `id` field is immutable (set at creation via `Math.max(...ids) + 1`) — avoids array-index mutation bugs
- `archived: false` default added via `migrate()` for backward compatibility with old records
- `coursePool` is the global list of available courses; `courses` on each associate is a `{ name: percentage }` map

---

## Frontend Component Structure

The entire UI is one component: `src/App.jsx`. No routing library is used.

### State

| State | Type | Purpose |
|---|---|---|
| `associates` | `Associate[]` | All associates (active + archived) |
| `coursePool` | `string[]` | Available courses |
| `loading` | `bool` | True while initial OneDrive fetch is in-flight |
| `syncing` | `bool` | True while debounced save is in-flight |
| `syncError` | `bool` | True if last save failed |
| `isAdmin` | `bool` | Shows admin panel |
| `managerView` | `bool` | Manager vs Team view toggle |
| `showAnalytics` | `bool` | Analytics panel open/closed |
| `expandedId` | `number\|null` | Which associate row is expanded |

### Refs (not state — don't trigger re-renders)

| Ref | Purpose |
|---|---|
| `saveTimer` | Holds the debounce timeout ID for cloud saves |
| `skipSave` | Prevents echoing the initial OneDrive load back as a save |
| `cloudReady` | Guards the persist effect until initial fetch completes |
| `chart1Ref`, `chart2Ref` | DOM refs for analytics Chart.js canvases |
| `chartInst1`, `chartInst2` | Chart.js instance refs (for cleanup before re-render) |

### Key Effects

```
useEffect([isAuthenticated, getToken])
  → Loads data from OneDrive after login

useEffect([associates, coursePool, getToken])
  → Debounced save to OneDrive on every data change
  → Skips if cloudReady is false (still loading)
  → Skips exactly once after initial load (skipSave flag)

useEffect([showAnalytics, associates, coursePool])
  → Destroys and recreates Chart.js instances for analytics panel
```

---

## Build & Deployment Pipeline

```
Developer pushes to main branch
       │
GitHub Actions: .github/workflows/deploy.yml
       │
  ├── actions/checkout@v4
  ├── actions/setup-node@v4 (Node 20)
  ├── npm ci
  ├── npm run build
  │     ├── Vite reads VITE_AAD_CLIENT_ID + VITE_AAD_TENANT_ID from GitHub Secrets
  │     ├── Inlines env vars into JS bundle at build time
  │     └── Outputs to dist/ (index.html + assets/ with content hashes)
  ├── actions/configure-pages@v4
  ├── actions/upload-pages-artifact@v3 (uploads dist/)
  └── actions/deploy-pages@v4 (publishes to github.io CDN)
       │
Live at: https://saurabhsuman1-droid.github.io/bench-dashboard/
```

**Note on env vars:** `VITE_AAD_CLIENT_ID` and `VITE_AAD_TENANT_ID` are baked into the JavaScript bundle at build time. This is safe — Azure AD Client IDs and Tenant IDs are **not secrets**. They identify the app to Azure AD but cannot be used to authenticate without the user's own EPAM credentials. The actual sensitive tokens are obtained at runtime via MSAL and stored only in the user's sessionStorage.

---

## Security Summary

| Concern | Mitigation |
|---|---|
| Who can access the app? | Anyone with a valid EPAM Microsoft account |
| Where is bench data stored? | EPAM OneDrive (Microsoft tenant) — never leaves Microsoft |
| Can GitHub see the data? | No — GitHub only serves static JS/HTML/CSS files |
| What if someone copies the GitHub JS bundle? | They get code but no data; Graph API rejects requests without a valid EPAM Bearer token |
| Token leakage? | Tokens in sessionStorage only; cleared on tab close |
| Can an unauthenticated user read data? | No — every Graph API call requires a valid Bearer token |
| What if the manager leaves EPAM? | Their OneDrive file becomes inaccessible; transfer ownership by copying `bench-data.json` to a new manager's OneDrive before offboarding |

---

## File Map

```
bench-dashboard/
├── src/
│   ├── App.jsx          ← Entire application (single component, ~900 lines)
│   ├── main.jsx         ← MSAL PublicClientApplication setup + React root
│   └── index.css        ← Tailwind directives only
├── .github/
│   └── workflows/
│       └── deploy.yml   ← GitHub Actions CI/CD pipeline
├── docs/
│   ├── USER_GUIDE.md
│   ├── ARCHITECTURE.md  ← this file
│   └── SETUP.md
├── CLAUDE.md            ← AI assistant context (read this first)
├── README.md
├── vite.config.js       ← base: './' for relative asset paths
├── tailwind.config.js
├── postcss.config.js
├── staticwebapp.config.json  ← Azure SWA SPA routing (if migrated to Azure later)
├── netlify.toml         ← Netlify config (kept for reference, not active)
├── .env.example         ← Template for required environment variables
└── .gitignore           ← Includes .env to prevent secret commit
```
